🎓 Academic PPT Agent (demo)
An AI-driven automation tool powered by the DeepSeek API and python-pptx.
This Agent understands academic requirements, parses local documents, and autonomously generates and executes Python code to create professional, research-oriented PowerPoint presentations.
⚠️ Important Notice (Demo Constraints)
Demo Purpose Only
This project is intended as a functional demonstration.
Format Support
Currently supports .docx and .txt files only.
No Iterative Modification
The tool is designed for one-shot generation. It does not support follow-up interactions or iterative refinement within the same session.
Environment Recommendation
It is strongly recommended to run this project inside a virtual environment (venv) to avoid dependency conflicts.
🌟 Key Features
🔹 Dual Input Modes
Generate slides from a simple topic
Or analyze local .docx / .txt files as input
🔹 Intelligent Outlining
Simulates an academic mentor’s logic
Automatically structures:
Title Slide
Agenda
Core Research Content
Conclusion
🔹 Autonomous Code Execution
The Agent writes its own python-pptx script
Executes it locally in a sandbox environment
Outputs a ready-to-use .pptx file
🔹 Self-Healing Mechanism
If generated code fails:
Captures traceback
Rewrites code automatically
Retries execution
🛠️ Requirements & Installation
1. Create Virtual Environment
# Create a virtual environment
python -m venv .venv

# Activate (macOS / Linux)
source .venv/bin/activate

# Activate (Windows)
.venv\Scripts\activate
2. Install Dependencies
pip install python-pptx python-docx openai
📦 Package Overview
Package	Description
python-pptx	Core engine for creating PowerPoint files
python-docx	Extracts text from Word documents
openai	Client for interacting with DeepSeek API
🚀 Quick Start
1. Configure API Key
Replace with your actual DeepSeek key:
self.client = openai.OpenAI(
    api_key="YOUR_DEEPSEEK_API_KEY",
    base_url="https://api.deepseek.com"
)
2. Run the Script
python your_script_name.py
3. Interaction Flow
🧭 Mode Selection
Enter y → Upload a file
Enter n → Start from a topic
📄 File Input
Drag & drop file into terminal
Or paste absolute path
💬 Command Input
Example:
Summarize the key findings and generate an academic presentation
📂 Workflow Architecture
1. Input Parsing
Cleans file paths
Extracts text from .docx / .txt
2. Planning Engine (Stage 1)
LLM generates structured academic outline
3. Rendering Engine (Stage 2)
Agent writes python-pptx code
Applies safety patches (e.g., removing MSO_ANCHOR)
4. Sandbox Execution
Executes code using isolated exec() environment
5. Output Reporting
Generates final file:
output_presentation.pptx
💡 Notes
Best suited for:
Academic presentations
Research summaries
Thesis defenses
Not recommended for:
Highly visual design-heavy slides
Interactive editing workflows
