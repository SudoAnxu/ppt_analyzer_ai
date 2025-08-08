# Presentation Inconsistency Detection Agent

This is a Python-based AI agent that analyzes presentation slides to find factual and logical inconsistencies. The agent can process a folder of slide images (JPEG, PNG) or a `.pptx` file directly (on Windows with PowerPoint installed). It produces a clear, structured report in the terminal detailing any contradictions found.

## How It Works: The Architectural Approach

The agent's intelligence comes from a **Three-Pass AI Architecture**. This approach was chosen to maximize accuracy and reasoning capability by using the Large Language Model (LLM) for what it does best (understanding language and context) while ensuring reliability.

### Pass 1: Raw Data Extraction
The agent first iterates through each slide image. It uses a detailed prompt (`EXTRACTION_PROMPT`) to instruct the Gemini AI model to act as a data analyst. It extracts every piece of factual data and, most importantly, assigns a **semantic category** to each fact (e.g., `total_productivity_savings_usd`, `feature_time_savings_breakdown`). It does *not* attempt to normalize or analyze at this stage, focusing only on high-quality, structured data capture.

### Pass 2: Contextual Normalization
After gathering all the raw data, the agent groups it by the AI-generated `metric_category`. It then sends each group of related facts back to the AI. This is the crucial step. By providing the AI with the full context of all related claims at once (e.g., all claims about "time saved per slide"), it can make an intelligent decision on how to normalize the data into a single, common unit (e.g., converting all `minutes` to `hours`). This is far more robust than a single-pass approach.

### Pass 3: The AI Reasoning Engine
With a complete, cleaned, and normalized dataset, the agent performs the final analysis. It bundles the entire case file into a single, final prompt (`AI_ANALYSIS_PROMPT`). This prompt instructs the AI to act as a world-class analyst, empowering it to find all types of inconsistencies on its own:
-   **Direct Numerical Contradictions**
-   **Incorrect Summations**
-   **Logical Contradictions** (e.g., qualitative claims vs. quantitative data)
-   **Omissions & Incompleteness**

This final step delegates the entire reasoning task to the AI, allowing it to find complex errors that would be difficult to pre-program with simple `if/else` logic.

## How to Run the Agent

### Prerequisites
*   Python 3.8+
*   Windows Operating System (for automatic `.pptx` conversion)
*   Microsoft PowerPoint installed (for automatic `.pptx` conversion)
*   A Google AI API Key

### 1. Clone the Repository
```bash
git clone <https://github.com/SudoAnxu/ppt_analyzer_ai>
cd <ppt_analyzer_ai>
```

### 2. Install Dependencies
```bash
pip install google-generativeai pillow pywin32 dotenv
```

### 3. Set Your API Key
For security, the script reads your API key from an environment variable. Do not hardcode it.

**On Windows (Command Prompt):**
```cmd
set GOOGLE_API_KEY="YOUR_API_KEY_HERE"
```

**On macOS/Linux:**
```bash
export GOOGLE_API_KEY="YOUR_API_KEY_HERE"
```

### 4. Execute the Script
Run the agent from your terminal, providing the path to your presentation file or image folder.

**To analyze a `.pptx` file directly (Windows only):**
```bash
python agent.py "C:\Users\YourName\Documents\My Presentation.pptx"
```
*(The agent will create a temporary folder named `temp_My Presentation_images` to store the converted slides.)*

**To analyze a folder of images:**
```bash
python agent.py ./path/to/your/slide_images_folder/
```

## Limitations
*   **Automatic `.pptx` conversion** is only supported on Windows with an active PowerPoint installation due to its reliance on COM automation. On other operating systems, the agent will guide the user to manually export images.
*   **LLM Reliability:** While using a `temperature` of `0.0` makes the AI's output more deterministic, its reasoning is still based on statistical patterns. It may occasionally miss a very subtle inconsistency or misinterpret a highly ambiguous claim.
