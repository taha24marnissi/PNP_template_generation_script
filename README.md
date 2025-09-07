# SharePoint PnP Provisioning XML Generator

🚀 **Enterprise-grade AI-powered SharePoint template generator** that converts natural language descriptions into comprehensive PnP Provisioning XML templates with advanced field types, complex structures, and enterprise features.

## 🌟 Key Features

### 🤖 **Advanced AI Integration**
- **GPT-4 Turbo Intelligence**: Sophisticated parsing of complex enterprise requirements
- **Natural Language Processing**: Understands business context and intent
- **Enterprise Pattern Recognition**: Automatically detects industry-standard structures
- **Smart Field Type Detection**: Identifies and creates appropriate SharePoint field types

### 🏗️ **Comprehensive Template Generation**
- **Multiple Site Types**: Team Sites, Communication Sites, Hub Sites
- **Advanced Field Types**: Text, Choice, DateTime, Boolean, Number, Currency, User/Person, Note/Multiline
- **Complex List Structures**: Document Libraries, Lists, with proper versioning and content types
- **Enterprise Navigation**: Multi-level navigation with proper structure
- **Microsoft-Compliant XML**: Follows official PnP Provisioning Schema 2022-09

### ✅ **Production-Ready Quality**
- **XSD Validation**: All templates validated against official PnP Schema
- **Real Deployment Testing**: Tested with actual SharePoint Online environments
- **Error-Free Field Creation**: Resolved "Column does not exist" issues using Microsoft's proven patterns
- **Scalable Architecture**: Handles templates with 25+ fields across multiple lists

### 🔧 **Advanced Capabilities**
- **Field Distribution Logic**: Automatically assigns relevant fields to appropriate lists
- **Enterprise Security**: Supports confidentiality levels, approval workflows
- **Complex Choice Fields**: Multi-option dropdowns with defaults and validation
- **Person/User Fields**: Proper user picker field implementation
- **Currency & Number Fields**: Financial and numerical data with proper formatting
- **Rich Text Support**: Note fields for complex content

## 🚀 Installation

### Prerequisites
- Python 3.7+
- OpenAI API Key (for AI features)
- SharePoint Online environment (for testing)

### Quick Setup
1. **Install Dependencies**:
```bash
pip install -r requirements.txt
```

2. **Configure AI Integration**:
   
   **Option A: .env File (Recommended)**
   ```bash
   # Create .env file in project root
   OPENAI_API_KEY=sk-proj-your-actual-api-key-here
   ```
   
   **Option B: Environment Variables**
   ```bash
   # Windows
   set OPENAI_API_KEY=your-openai-api-key-here

   # Linux/Mac
   export OPENAI_API_KEY=your-openai-api-key-here
   ```

3. **Get OpenAI API Key**: https://platform.openai.com/api-keys

## 💻 Usage

### Command Line Interface
```bash
# Simple example
python generate_template.py "Create a project management site with document tracking"

# Complex enterprise example
python generate_template.py "Create an enterprise legal contracts library with approval workflows, confidentiality levels, and vendor management"
```

### Interactive Mode
```bash
python generate_template.py
# Follow the prompts to describe your requirements
```

## 🎯 Real-World Examples

### **Simple Project Management**
```bash
python generate_template.py "Create a project management site with document library, task tracking, and team collaboration features"
```

### **Enterprise Knowledge Management**
```bash
python generate_template.py "Create an enterprise knowledge management communication site with: policy documents library, training materials library, company updates list, FAQ management list. Include navigation and category/priority fields."
```

### **Advanced Legal & Vendor Portal**
```bash
python generate_template.py "Create an ultra-complex enterprise solution with Legal Contracts library (ContractType, ContractValue, SigningDate, LegalReviewer, ApprovalLevel, ComplianceStatus) and Vendor Management list (VendorID, VendorType, ContactPerson, PaymentTerms, CreditRating, SecurityClearance)"
```

### **HR Management System**
```bash
python generate_template.py "Create an HR management site with employee documents library, policy tracking, training schedules, performance reviews, and department management"
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
- **Content Types**: Custom content types with field associations
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

## 🧪 Testing & Deployment

### **PowerShell Deployment Script**
The included `invoke-template.ps1` script provides:
- **Certificate-based Authentication**
- **Automated Site Creation**
- **Template Deployment**
- **Comprehensive Logging**
- **Error Handling**

### **Usage**
```powershell
.\script\invoke-template.ps1 -TemplateFilePath ".\generated-templates\your-template.xml"
```

## 🎯 Proven Results

### **Complexity Testing**
- ✅ **Simple Templates**: 1-4 fields, single library
- ✅ **Enterprise Templates**: 20+ fields, multiple lists
- ✅ **Ultra-Complex**: 28+ fields, advanced field types
- ✅ **Real Deployment**: Tested with SharePoint Online

### **Field Creation Success**
- ✅ **Resolved Legacy Issues**: Fixed "Column does not exist" errors
- ✅ **Microsoft Pattern Compliance**: Uses proven field creation approach
- ✅ **Advanced Types**: Currency, Person, Note fields working
- ✅ **Enterprise Scale**: 500+ document capacity tested

## 🔧 Technical Architecture

### **AI Processing Pipeline**
1. **Natural Language Input** → GPT-4 Analysis
2. **Structure Generation** → JSON Schema Creation  
3. **XML Generation** → PnP Template Building
4. **Validation** → XSD Schema Validation
5. **Deployment** → SharePoint Provisioning

### **Field Creation Innovation**
- **Microsoft Pattern Implementation**: Uses list-specific field definitions
- **Automatic Field Distribution**: Smart assignment of fields to appropriate lists
- **Type-Specific Attributes**: Proper attributes for each field type
- **Enterprise Standards**: Follows SharePoint best practices

## 📋 Requirements

### **Core Dependencies**
```
openai>=1.0.0          # AI integration
lxml>=4.9.0            # XML validation
python-dotenv>=1.0.0   # Configuration management
```

### **Optional Components**
- **PnP PowerShell Module**: For template deployment
- **SharePoint Online**: For testing and validation
- **VS Code**: For development and debugging

## 🚀 Enterprise Features

### **Scalability**
- **Multi-List Templates**: Handle complex organizational structures
- **Bulk Field Creation**: 25+ fields per template
- **Large Document Libraries**: 500+ document capacity
- **Performance Optimized**: Efficient XML generation

### **Security & Compliance**
- **Confidentiality Levels**: Public, Internal, Confidential, Restricted, Top Secret
- **Approval Workflows**: Department, VP, CEO, Board level approvals
- **User Permissions**: Proper person field implementation
- **Audit Trail**: Comprehensive logging and reporting

### **Business Process Support**
- **Legal Contract Management**: Full contract lifecycle
- **Vendor Management**: Complete vendor relationship tracking
- **Project Management**: Advanced project tracking capabilities
- **Knowledge Management**: Enterprise knowledge sharing

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Test with real SharePoint environments  
4. Submit pull request with validation reports

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- **Microsoft PnP Community**: For the provisioning schema and patterns
- **OpenAI**: For GPT-4 integration capabilities
- **SharePoint Community**: For field creation best practices

---

**Built with ❤️ for Enterprise SharePoint Solutions**