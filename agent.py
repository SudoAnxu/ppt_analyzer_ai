import os
import google.generativeai as genai
from PIL import Image
import json
from pathlib import Path
import logging
import time
import win32com.client
import argparse
from dotenv import load_dotenv

# --- Setup Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
load_dotenv()

# --- Secure Configuration ---
try:
    genai.configure(api_key=os.environ['GOOGLE_API_KEY'])
    
    logging.info("Google GenAI configured successfully.")
except KeyError:
    logging.error("🚨 Critical Error: GOOGLE_API_KEY environment variable not set.")
    exit()

# --- Model Configuration ---
generation_config = genai.GenerationConfig(temperature=0.0)
model = genai.GenerativeModel('gemini-2.5-flash', generation_config=generation_config)

# --- Prompt 1: For the initial, simple extraction ---
EXTRACTION_PROMPT = """
You are an expert data analyst AI. Your task is to analyze the provided presentation slide and extract all factual claims (textual and numerical) into a structured JSON format.

It is crucial that you also **semantically categorize** each metric so that data from different slides can be compared, even if the wording is different. Do NOT normalize or change any values yet.

Your output MUST be a JSON object containing a list called "elements". For each piece of data, create a JSON object with the following schema:
- "metric_category": A standardized, snake_case name for the data's category (e.g., "total_productivity_savings_usd", "time_saved_per_slide", "competitor_time_savings", "total_time_savings_claim", "feature_time_savings_breakdown", "qualitative_claim").
- "feature_name": If the category is a breakdown item, specify the feature's name here (e.g., "Automated Formatting"). Otherwise, null.
- "text_content": The raw, extracted text.
- "numerical_value": The original number, if one exists (e.g., for "$2M", this is 2000000). Null if no number.
- "unit": The original unit, if one exists (e.g., "USD", "hours", "minutes"). Null if no unit.

Analyze the slide and generate ONLY the JSON output.
"""

# --- Prompt 2: For the second, context-aware normalization pass ---
AI_NORMALIZING_PROMPT_TEMPLATE = """
You are a precise data normalization engine. I have provided you with a list of JSON objects below. These objects all belong to the SAME semantic category but were extracted from different slides.

Your task is to analyze the list and normalize all the items to a single, consistent unit. Follow these rules:
1.  Examine all the units present in the list (e.g., "hours", "minutes").
2.  Choose the most logical and common base unit for comparison (e.g., for time, 'hours' is usually best).
3.  Rewrite the entire list of JSON objects, adding two new keys: "normalized_value" and "normalized_unit".
4.  Perform the mathematical conversions accurately. For time, 60 minutes = 1 hour. For currency, "$2M" = 2000000.
5.  Return ONLY the updated list of JSON objects.

Here is the list of data to normalize:
{json_list}
"""

# --- Prompt 3: For the final logical analysis ---
AI_ANALYSIS_PROMPT = """
You are a world-class logical and factual analyst. I will provide you with a complete, grouped, and normalized set of data extracted from a multi-slide presentation.

Your sole task is to meticulously analyze this entire dataset to find ALL factual and logical inconsistencies.

Look for these specific types of problems:
1.  **Direct Numerical Contradictions:** When the same metric (e.g., 'total_productivity_savings_usd') has different numerical values across slides.
2.  **Incorrect Summations:** When a claimed total (e.g., 'total_time_savings_claim') does not equal the sum of its component parts (e.g., 'feature_time_savings_breakdown').
3.  **Logical Contradictions:** When a qualitative claim (e.g., "Tool X is superior") is contradicted by quantitative data (e.g., Tool X saves less time than Tool Y).
4.  **Omissions & Incompleteness:** When a list of features or benefits on one slide is different from a list on another, indicating an incomplete picture.

Your output MUST be a single JSON object containing a list called "findings". Each object in the list must have the following schema:
- "type_of_inconsistency": A string describing the category of the error (e.g., "Incorrect Summation").
- "description": A clear, human-readable paragraph explaining the inconsistency.
- "evidence": A list of the specific text snippets and slide numbers that prove the inconsistency.

Now, analyze the following complete dataset and generate your findings.

{grouped_data_json}
"""
def convert_pptx_to_images(pptx_path, output_folder):
    """
    Converts each slide of a .pptx file into a .jpeg image using
    Windows COM automation to control PowerPoint.
    Returns the path to the folder where images were saved.
    """
    powerpoint = None # Initialize to ensure it's defined for the finally block
    try:
        # Created a temporary folder for the images
        if not output_folder.exists():
            output_folder.mkdir(parents=True)

        logging.info(f"Attempting to convert '{pptx_path.name}' using PowerPoint...")
        
        # Connecting to the PowerPoint Application COM object
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        # Opening the presentation
        presentation = powerpoint.Presentations.Open(str(pptx_path), WithWindow=False)
        
        # Saving each slide as a high-quality JPEG
        for i, slide in enumerate(presentation.Slides):
            image_name = f"slide_{i+1}.jpg"
            image_path = output_folder / image_name
            slide.Export(str(image_path), "JPG")
            logging.info(f"  - Saved {image_name}")

        # Closing the presentation
        presentation.Close()
        logging.info("Conversion successful.")
        return str(output_folder)

    except Exception as e:
        logging.error(f"  ❌ PowerPoint conversion failed. Please ensure PowerPoint is installed.")
        logging.error(f"     Error details: {e}")
        return None
    
    finally:
        # Ensure PowerPoint is closed, even if an error occurred
        if powerpoint:
            powerpoint.Quit()
            
class AiReasoningAgent:
    def __init__(self, slide_folder_path):
        self.slide_folder_path = Path(slide_folder_path)
        self.raw_data = []
        self.grouped_data = {}
        self.inconsistencies = []

    def run_analysis(self):
        logging.info("🚀 Starting Two-Pass AI Presentation Analyzer...")
        self._extract_raw_data()
        if not self.raw_data: return
        self._group_raw_data()
        self._normalize_groups_with_ai()
        self._analyze_for_inconsistencies() # To performs all checks
        self._generate_report()

    def _extract_raw_data(self):
        # This function from the previous version is correct.
        slide_paths = sorted([p for p in self.slide_folder_path.glob("*") if p.suffix.lower() in [".jpg", ".jpeg", ".png"]])
        logging.info(f"Found {len(slide_paths)} slides for Pass 1 Extraction.")
        for i, slide_path in enumerate(slide_paths):
            logging.info(f"📄 Pass 1 - Extracting from {slide_path.name}...")
            try:
                img = Image.open(slide_path)
                response = model.generate_content([EXTRACTION_PROMPT, img], request_options={"timeout": 120})
                slide_data = json.loads(response.text.strip().replace("```json", "").replace("```", ""))
                for element in slide_data.get("elements", []):
                    element['slide'] = i + 1
                self.raw_data.extend(slide_data.get("elements", []))
                print(f"Element in slide {i + 1}: {len(slide_data.get('elements', []))} items extracted.")
            except Exception as e:
                logging.error(f"  ⚠️ Error during Pass 1 on {slide_path.name}: {e}")

    def _group_raw_data(self):
        # This function is also correct.
        logging.info("📊 Grouping extracted data by category...")
        for item in self.raw_data:
            group_key = item.get('metric_category')
            if group_key:
                if group_key not in self.grouped_data: self.grouped_data[group_key] = []
                self.grouped_data[group_key].append(item)

    def _normalize_groups_with_ai(self):
        # This function is also correct.
        logging.info("🧠 Starting Pass 2 - AI-led Contextual Normalization...")
        normalized_groups = {}
        for category, items in self.grouped_data.items():
            if len(items) > 1 and any(item.get('numerical_value') is not None for item in items):
                logging.info(f"  - Normalizing category: '{category}'...")
                try:
                    json_string_of_items = json.dumps(items, indent=2)
                    prompt = AI_NORMALIZING_PROMPT_TEMPLATE.format(json_list=json_string_of_items)
                    response = model.generate_content(prompt, request_options={"timeout": 240})
                    normalized_items = json.loads(response.text.strip().replace("```json", "").replace("```", ""))
                    normalized_groups[category] = normalized_items
                    time.sleep(1) 
                except Exception as e:
                    logging.error(f"  ⚠️ Error during Pass 2 on category '{category}': {e}")
                    normalized_groups[category] = items
            else:
                normalized_groups[category] = items
        self.grouped_data = normalized_groups
        print(f"Normalization complete. {len(self.grouped_data)} categories processed.")
        
    def _analyze_for_inconsistencies(self):
        """
        This is the new "AI Brain". It offloads all reasoning to the LLM.
        """
        logging.info("🧠 Handing off full dataset to AI Reasoning Engine for final analysis...")
        
        try:
            # Convert the entire "case file" of grouped data into a string.
            grouped_data_string = json.dumps(self.grouped_data, indent=2)
            
            # Creating the final, single prompt for the AI.
            prompt = AI_ANALYSIS_PROMPT.format(grouped_data_json=grouped_data_string)
            logging.info("  - Sending data to AI for logical analysis...")
            # Giving it more time for this complex task
            response = model.generate_content(prompt, request_options={"timeout": 300}) 
            
            # The AI will return a JSON object with a "findings" list
            analysis_result = json.loads(response.text.strip().replace("```json", "").replace("```", ""))
            logging.info("  - AI analysis complete. Processing findings...")
                      
            # Store the AI's findings in the agent's state.
            self.inconsistencies = analysis_result.get("findings", [])

        except Exception as e:
            logging.error(f"  ⚠️ A critical error occurred during the AI reasoning phase: {e}")

    def _generate_report(self):
        """This report generator is now simpler, as it prints the AI's findings directly."""
        print("\n\n" + "="*30)
        print("  AI REASONING ENGINE REPORT")
        print("="*30 + "\n")
        if not self.inconsistencies:
            print("✅ The AI Reasoning Engine reported no inconsistencies.")
            return

        for i, issue in enumerate(self.inconsistencies, 1):
            print(f"**{i}. {issue.get('type_of_inconsistency', 'Unknown Issue')}**")
            print(f"   - **AI's Analysis:** {issue.get('description', 'No description provided.')}")
            if issue.get('evidence'):
                print("   - **Evidence Cited by AI:**")
                for ev in issue['evidence']:
                    print(f"     - {ev}")
            print("\n")
        
        print("--- END OF REPORT ---")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Analyze a presentation for inconsistencies.")
    parser.add_argument("input_path", type=str, help="Path to a .pptx file or a folder of slide images.")
    
    args = parser.parse_args()
    input_path = Path(args.input_path).resolve() ### Use absolute path

    analysis_folder = None

    if not input_path.exists():
        logging.error(f"The specified path does not exist: {input_path}")
    
    # --- INTELLIGENT PATH HANDLING ---
    elif input_path.is_file() and input_path.suffix.lower() == '.pptx':
        # Check if we are on Windows
        if os.name == 'nt':
            # Create a temporary folder for the images next to the pptx
            temp_image_folder = input_path.parent / f"{input_path.stem}_images"
            analysis_folder = convert_pptx_to_images(input_path, temp_image_folder)
        else:
            # On other OS, guide the user as automatic conversion isn't supported.
            logging.warning("A .pptx file was provided on a non-Windows OS.")
            print("\n========================= ACTION REQUIRED =========================")
            print("Automatic .pptx conversion is only supported on Windows with PowerPoint installed.")
            print(f"To analyze '{input_path.name}', please first export your slides as images into a folder.")
            print(f"Then, run the agent again on that new folder.")
            print("===================================================================")

    elif input_path.is_dir():
        # The path is already a directory of images
        analysis_folder = str(input_path)
    
    else:
        logging.error(f"Invalid input path. Please provide a path to a valid directory or a .pptx file.")

    # --- RUNNING THE ANALYSIS ---
    # Only run the analyzer if we have a valid folder path to work with
    if analysis_folder:
        analyzer = AiReasoningAgent(analysis_folder) 
        analyzer.run_analysis()
