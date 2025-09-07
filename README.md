# SharePoint PnP Provisioning XML Generator

A Python CLI tool that converts natural language descriptions into SharePoint PnP Provisioning XML templates.

## Features

- 🤖 **ChatGPT Integration**: Uses OpenAI's GPT-3.5 for intelligent parsing of site requirements
- 🔄 **Fallback Parsing**: Works without ChatGPT using basic regex patterns
- ✅ **Proper XML**: Generates valid PnP Provisioning XML with correct namespaces
- 📋 **Smart Detection**: Automatically detects site types, lists, libraries, and fields
- 🧭 **Navigation**: Includes appropriate navigation structure

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Set up ChatGPT integration (optional but recommended):
   
   **Option A: Using .env file (recommended)**
   ```bash
   # Copy the example file
   cp .env.example .env
   
   # Edit .env and add your OpenAI API key
   OPENAI_API_KEY=sk-proj-your-actual-api-key-here
   ```
   
   **Option B: Using environment variables**
   ```bash
   # Windows
   set OPENAI_API_KEY=your-openai-api-key-here

   # Linux/Mac
   export OPENAI_API_KEY=your-openai-api-key-here
   ```

3. Get your OpenAI API key from: https://platform.openai.com/api-keys

## Usage

### Command Line
```bash
python generate_template.py "Create a team site with a document library called 'Project Files'"
```

### Interactive Mode
```bash
python generate_template.py
# Then enter your description when prompted
```

### Examples

**Team Site with Custom Library:**
```bash
python generate_template.py "Create a team site for project management with a document library called 'Project Files'"
```

**Communication Site for HR:**
```bash
python generate_template.py "Create a communication site for HR policies with document library and announcements"
```

**Educational Hub:**
```bash
python generate_template.py "Educational hub with class materials, events calendar and student resources"
```

## Output

Generated XML files are saved in the `generated-templates/` directory with timestamps.

## ChatGPT vs Fallback

| Feature | With ChatGPT (GPT-4) | Fallback Mode |
|---------|----------------------|---------------|
| Site Title Extraction | ✅ Advanced with context | ✅ Basic pattern matching |
| List/Library Detection | ✅ Intelligent parsing | ✅ Regex-based |
| Field Generation | ✅ Context-aware | ✅ Standard fields |
| Content Types | ✅ Smart detection | ❌ Limited |
| Complex Scenarios | ✅ Excellent | ⚠️ Basic |
| Natural Language | ✅ Understands intent | ⚠️ Keywords only |
| Quoted Names | ✅ Perfect extraction | ✅ Good extraction |

## Requirements

- Python 3.7+
- lxml (for XML validation)
- openai (optional, for ChatGPT integration)