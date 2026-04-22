# 🎓 Academic PPT Agent (AutoPPT)

An AI-driven automation tool powered by the DeepSeek API and `python-pptx`. This Agent understands academic requirements, parses local documents, and autonomously writes and executes Python code to generate professional, research-oriented PowerPoint presentations.

---

## ⚠️ Important Notice (Demo Constraints)

- **Demo Purpose Only**: This project is a functional demo.  
- **Format Support**: It currently supports `.docx` and `.txt` files only.  
- **No Iterative Modification**: The tool is designed for one-shot generation. It does not support follow-up questions or iterative modifications to the generated PPT within the same session.  
- **Environment**: It is highly recommended to run this script within a virtual environment (`venv`) to avoid dependency conflicts.  

---

## 🌟 Key Features

- **Dual Input Modes**: Generate slides from a simple topic or by analyzing local `.docx` and `.txt` reference files.  
- **Intelligent Outlining**: Simulates an academic mentor's logic to structure content (Title slide, Agenda, Core Research, and Conclusion).  
- **Autonomous Code Execution**: The Agent writes its own `python-pptx` script and runs it in a local sandbox to render the `.pptx` file.  
- **Self-Healing Mechanism**: If the generated code encounters an error, the Agent analyzes the traceback and re-writes the code to fix the issue.  

---

## 🛠️ Requirements & Installation

It is recommended to set up a virtual environment first:

```bash
# Create a virtual environment
python -m venv .venv

# Activate it (macOS/Linux)
source .venv/bin/activate

# Activate it (Windows)
.venv\Scripts\activate

# Install dependencies
pip install python-pptx python-docx openai

# 🚀 Quick Start

## 1. Configure API Key
Replace the placeholder in the `AutoPPTAgent` class with your actual DeepSeek key:

```python
self.client = openai.OpenAI(
    api_key="YOUR_DEEPSEEK_API_KEY",
    base_url="https://api.deepseek.com"
)
````

## 2. Run the Script

```bash
python your_script_name.py
```

## 3. Interaction Flow

* **Reference Mode**: Type `y` to upload a file or `n` to start from a topic
* **File Path**: Drag and drop your file into the terminal or paste the absolute path
* **Command**: Enter your specific requirements (e.g., *"Summarize the key findings"*)

---

# 📂 Workflow Architecture

## 1. Input Parsing

Cleans file paths and extracts text content from local documents.

## 2. Planning Engine (Stage 1)

The LLM generates a structured academic outline.

## 3. Rendering Engine (Stage 2)

The agent writes `python-pptx` code and applies safety patches (e.g., `MSO_ANCHOR` filtering).

## 4. Sandbox Execution

Runs the generated code in an isolated `exec()` environment.

## 5. Reporting

Outputs the final file as:

```
output_presentation.pptx
```

```
```
