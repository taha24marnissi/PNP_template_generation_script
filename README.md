# SharePoint PnP Provisioning XML Generator

🚀 **AI-powered SharePoint template generator** that converts natural language descriptions into PnP Provisioning XML templates with advanced field types and ## 📋 Requirements

**Core Dependencies:**
```
openai>=1.0.0          # AI integration
lxml>=4.9.0            # XML validation and CAML processing
python-dotenv>=1.0.0   # Configuration management
```iew generation.

## 🌟 Key Features

### 🤖 **AI Integration**
- **GPT-4 Intelligence**: Natural language parsing of enterprise requirements
- **Smart Field Detection**: Identifies and creates appropriate SharePoint field types
- **Intelligent Views**: AI creates custom views with CAML queries based on context

### 🏗️ **Template Generation**
- **Multiple Site Types**: Team Sites, Communication Sites, Hub Sites
- **Advanced Field Types**: Text, Choice, DateTime, Boolean, Number, Currency, User/Person, Note/Multiline
- **List Structures**: Document Libraries and Custom Lists with proper versioning
- **Enterprise Navigation**: Multi-level navigation with proper structure
- **Custom Views**: AI-generated views with complex CAML queries for filtering, sorting, and grouping

### 🎯 **View Generation**
- **Multiple View Types**: HTML, CALENDAR views supported
- **Smart CAML Queries**: Complex filtering with Where, OrderBy, GroupBy clauses
- **Field Name Mapping**: Automatic resolution of field conflicts
- **Real-Time Validation**: Proper XML formatting for SharePoint compatibility

### ✅ **Production-Ready**
- **XSD Validation**: All templates validated against official PnP Schema
- **Real Deployment Testing**: Tested with actual SharePoint Online environments
- **Scalable Architecture**: Handles templates with 25+ fields across multiple lists

## 🚀 Installation

### Prerequisites
- Python 3.7+
- OpenAI API Key

### Quick Setup
1. **Install Dependencies**:
```bash
pip install -r requirements.txt
```

2. **Set OpenAI API Key**:
```bash
# Create .env file in project root
OPENAI_API_KEY=sk-proj-your-actual-api-key-here
```

3. **Get API Key**: https://platform.openai.com/api-keys

## 💻 Usage

```bash
# Simple example
python generate_template.py "Create a project management site with document tracking"

# Complex enterprise example
python generate_template.py "Create an enterprise legal contracts library with approval workflows and vendor management"

# With views
python generate_template.py "Create task management site with Priority, Status, DueDate fields. Create 'High Priority Tasks' view and calendar view for due dates."
```

## 🎯 Examples

### **Project Management**
```bash
python generate_template.py "Create a project management site with document library, task tracking, and team collaboration features"
```

### **Event Management with Views**
```bash
python generate_template.py "Create an event management site with an events list having Title, EventDate, Category, Priority fields. Create calendar view for EventDate field and high priority events view."
```

### **Enterprise Knowledge Management**
```bash
python generate_template.py "Create an enterprise knowledge management communication site with: policy documents library, training materials library, company updates list, FAQ management list."
```

## 🔍 View Generation

### **Supported View Types**
- **HTML Views**: Standard list views with filtering and sorting
- **CALENDAR Views**: Date-based calendar display for events and schedules

### **CAML Query Features**
- **Filtering**: Complex Where clauses with Eq, Lt, Gt, Contains operators
- **Sorting**: OrderBy with Ascending/Descending
- **Grouping**: GroupBy with collapse and limits
- **Logic**: And/Or operators for sophisticated filtering
- **Date Functions**: Today(), Now() for dynamic comparisons

### **Example CAML Query**
```xml
<Query>
  <Where>
    <And>
      <Eq>
        <FieldRef Name="Priority"/>
        <Value Type="Choice">High</Value>
      </Eq>
      <Lt>
        <FieldRef Name="DueDate"/>
        <Value Type="DateTime"><Today/></Value>
      </Lt>
    </And>
  </Where>
  <OrderBy>
    <FieldRef Name="DueDate" Ascending="TRUE"/>
  </OrderBy>
</Query>
```

## 📊 Supported Field Types

| Field Type | SharePoint Type | Use Cases | Status |
|------------|----------------|-----------|---------|
| **Text** | Text | Names, IDs, short descriptions | ✅ Full Support |
| **Choice** | Choice | Dropdowns, status fields, categories | ✅ Full Support |
| **DateTime** | DateTime | Dates, deadlines, timestamps | ✅ Full Support |
| **Boolean** | Boolean | Yes/No, Active/Inactive flags | ✅ Full Support |
| **Number** | Number | Quantities, scores, percentages | ✅ Full Support |
| **Currency** | Currency | Budgets, costs, financial values | ✅ Full Support |
| **User/Person** | User | Assignees, managers, contacts | ✅ Full Support |
| **Note** | Note | Rich text, multiline descriptions | ✅ Full Support |

## 🏗️ Generated Template Structure

### **Enterprise Template Components**
- **Site Configuration**: Title, description, template type
- **Multiple Lists/Libraries**: Document libraries, custom lists
- **Advanced Fields**: All major SharePoint field types
- **Navigation Structure**: Multi-level navigation menus
- **Versioning & Security**: Proper versioning and permission settings

### **XML Quality Standards**
- ✅ **PnP Schema 2022-09 Compliant**
- ✅ **Microsoft Pattern Implementation**
- ✅ **Real Deployment Tested**
- ✅ **XSD Validated**
- ✅ **Enterprise Scalable**

## 📁 Output Structure

```
generated-templates/
├── enhanced_mock_YYYYMMDD_HHMMSS.xml    # Generated template
├── ...

validation-reports/
├── comprehensive_report_YYYYMMDD_HHMMSS.txt    # Validation report
├── ...

Scriptlogs/
├── TraceOutput-YYYYMMDD-HHMMSS.txt    # Deployment logs
├── ...
```

## 🧪 Deployment

The included `invoke-template.ps1` script provides automated SharePoint deployment:

```powershell
.\script\invoke-template.ps1 -TemplateFilePath ".\generated-templates\your-template.xml"
```

## 🎯 Proven Results

✅ **Tested with SharePoint Online**  
✅ **Complex templates with 25+ fields**  
✅ **CAML queries and calendar views working**  
✅ **Enterprise-scale deployments successful**

## 🔧 Technical Architecture

### **AI Processing Pipeline**
1. **Natural Language Input** → GPT-4 Analysis
2. **Structure Generation** → JSON Schema with Views
3. **Field & View Processing** → Automatic conflict resolution and CAML generation
4. **XML Generation** → PnP Template Building
5. **Validation** → XSD Schema Validation

### **Key Innovations**
- **AI-Powered View Creation**: LLM generates appropriate views based on context
- **CAML Query Intelligence**: Complex filtering, sorting, and grouping logic
- **Field Name Resolution**: Automatic mapping and conflict resolution
- **XML Parsing Engine**: Proper CAML XML formatting (no HTML encoding)

## � Prompt Engineering Guide

### **Be Specific with View Requirements**
```bash
# ✅ Good
"Create task list with Priority, Status, DueDate fields. Create 'High Priority Tasks' view and calendar view for due dates."

# ❌ Vague  
"Create a task list with some views"
```

### **Use Descriptive Field Names**
- Include field names in view descriptions
- Use Choice fields for filtering (Status, Priority, Category)
- Use Date fields for calendar views
- Use Person fields for "My Items" views

## �📋 Requirements

### **Core Dependencies**
```
openai>=1.0.0          # AI integration
lxml>=4.9.0            # XML validation and CAML processing
python-dotenv>=1.0.0   # Configuration management
```

### **Optional Components**
- **PnP PowerShell Module**: For template deployment
- **SharePoint Online**: For testing and validation
- **VS Code**: For development and debugging

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Test with real SharePoint environments  
4. Submit pull request with validation reports

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

**Built with ❤️ for Enterprise SharePoint Solutions**