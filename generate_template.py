#!/usr/bin/env python3
"""
SharePoint PnP Provisioning Template Generator

This tool generates SharePoint PnP Provisioning XML templates with AI assistance.
It creates comprehensive site structures including lists, libraries, content types,
site columns, navigation, and more.

Usage:
    python generate_template.py "Create a communication site with a document library..."
    python generate_template.py  # Interactive mode
"""

import sys
import re
import uuid
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any
import xml.etree.ElementTree as ET
from xml.dom import minidom

try:
    from lxml import etree
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False
    print("Warning: lxml package not found. XSD validation disabled. Install with: pip install lxml")

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    print("Warning: openai package not found. Install with: pip install openai")

try:
    from dotenv import load_dotenv
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False
    print("Warning: python-dotenv package not found. Install with: pip install python-dotenv")

def validate_xml_against_xsd(xml_content: str, xsd_path: str = "ProvisioningSchema-2022-09.xsd") -> tuple[bool, list]:
    """
    Validate XML against PnP Provisioning Schema XSD
    Returns (is_valid, errors_list)
    """
    if not LXML_AVAILABLE:
        return False, ["lxml package not available for XSD validation"]
    
    try:
        # Check if XSD file exists
        if not os.path.exists(xsd_path):
            return False, [f"XSD schema file not found: {xsd_path}"]
        
        # Read XSD schema
        with open(xsd_path, 'r', encoding='utf-8') as f:
            xsd_content = f.read()
        
        # Parse XSD
        xsd_doc = etree.fromstring(xsd_content.encode('utf-8'))
        xsd_schema = etree.XMLSchema(xsd_doc)
        
        # Parse XML to validate
        xml_doc = etree.fromstring(xml_content.encode('utf-8'))
        
        # Validate
        is_valid = xsd_schema.validate(xml_doc)
        errors = []
        
        if not is_valid:
            for error in xsd_schema.error_log:
                errors.append(f"Line {error.line}, Column {error.column}: {error.message}")
        
        return is_valid, errors
        
    except Exception as e:
        return False, [f"Validation error: {str(e)}"]

def save_llm_output(json_content: str, timestamp: str) -> str:
    """Save LLM JSON output to file"""
    os.makedirs("llm-outputs", exist_ok=True)
    filename = f"llm-outputs/llm_response_{timestamp}.json"
    
    try:
        # Pretty format JSON
        parsed_json = json.loads(json_content)
        formatted_json = json.dumps(parsed_json, indent=2, ensure_ascii=False)
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(formatted_json)
        
        return filename
    except Exception as e:
        # Save as plain text if JSON parsing fails
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(json_content)
        return filename

def save_comprehensive_report(json_content: str, xml_filename: str, is_valid: bool, validation_errors: list, timestamp: str) -> str:
    """Save comprehensive report with JSON structure and XSD validation results"""
    os.makedirs("validation-reports", exist_ok=True)
    filename = f"validation-reports/comprehensive_report_{timestamp}.txt"
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write("SharePoint PnP Provisioning Template - Comprehensive Report\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"Timestamp: {timestamp}\n")
        f.write(f"XSD Schema: ProvisioningSchema-2022-09.xsd\n")
        f.write(f"Generated XML: {xml_filename}\n\n")
        
        # XSD Validation Results
        f.write("XSD VALIDATION RESULTS\n")
        f.write("-" * 25 + "\n")
        if is_valid:
            f.write("✅ VALIDATION STATUS: PASSED\n")
            f.write("The XML template is valid according to the PnP Provisioning Schema.\n")
            f.write("✓ Ready for SharePoint deployment\n")
        else:
            f.write("❌ VALIDATION STATUS: FAILED\n")
            f.write(f"Found {len(validation_errors)} validation error(s):\n\n")
            for i, error in enumerate(validation_errors, 1):
                f.write(f"  {i}. {error}\n")
            f.write(f"\n⚠️  Template may not deploy properly to SharePoint.\n")
        
        # JSON Structure
        f.write(f"\n" + "=" * 60 + "\n")
        f.write("LLM GENERATED JSON STRUCTURE\n")
        f.write("=" * 60 + "\n")
        
        try:
            # Pretty format JSON for readability
            parsed_json = json.loads(json_content)
            formatted_json = json.dumps(parsed_json, indent=2, ensure_ascii=False)
            f.write(formatted_json)
        except:
            # Fallback to raw content if JSON parsing fails
            f.write(json_content)
        
        f.write(f"\n\n" + "=" * 60 + "\n")
        f.write("REPORT END\n")
        f.write("=" * 60 + "\n")
    
    return filename

def llm_generate_structure(description: str) -> Dict[str, Any]:
    """
    Use OpenAI ChatGPT to generate a proper JSON structure for SharePoint site requirements.
    Falls back to basic parsing if OpenAI is not available.
    """
    if not OPENAI_AVAILABLE:
        print("⚠️  OpenAI not available, using basic parsing...")
        return fallback_generate_structure(description)
    
    # Load environment variables from .env file
    if DOTENV_AVAILABLE:
        load_dotenv()
    
    # Get API key from .env file or environment variable
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        print("⚠️  OPENAI_API_KEY not found in .env file or environment variables, using basic parsing...")
        print("💡 Create a .env file with: OPENAI_API_KEY=your-api-key-here")
        return fallback_generate_structure(description)
    
    try:
        # Initialize OpenAI client
        client = openai.OpenAI(api_key=api_key)
        
        # Create the enhanced prompt for GPT-4
        prompt = f"""
You are a SharePoint expert with deep knowledge of PnP Provisioning Templates. Based on the following description, generate a comprehensive JSON structure for a SharePoint PnP provisioning template.

Description: "{description}"

Analyze the description carefully and extract:
1. Site type and purpose
2. Exact names from quotes or "called" phrases
3. Document libraries and their purposes
4. Lists and their types
5. Relevant site columns/fields
6. Content types that group related fields
7. Navigation structure

Return ONLY a valid JSON object with this exact structure:

{{
    "site_type": "TeamSite" or "CommunicationSite",
    "base_template": "GROUP#0" for modern TeamSite or "SITEPAGEPUBLISHING#0" for CommunicationSite,
    "site_title": "Meaningful title extracted from description - be specific",
    "description": "Brief description of the site purpose",
    "site_fields": [
        {{
            "name": "FieldName",
            "displayName": "Display Name",
            "type": "Text|Choice|DateTime|User|Number|Boolean|Lookup",
            "group": "Custom Columns",
            "choices": ["Option1", "Option2"] // only for Choice fields,
            "required": false
        }}
    ],
    "content_types": [
        {{
            "name": "Content Type Name",
            "description": "Content type description",
            "group": "Custom Content Types",
            "parent": "Document", // or "Item" for list items
            "fields": ["FieldName1", "FieldName2"] // reference site field names
        }}
    ],
    "lists": [
        {{
            "title": "Exact List/Library Title",
            "template_type": 100, // 100=Custom List, 101=Document Library, 104=Announcements, 105=Contacts, 106=Events, 107=Tasks
            "url": "Lists/ListName or LibraryName",
            "description": "Purpose of this list/library",
            "enable_versioning": true, // only for document libraries (101)
            "enable_content_types": true,
            "content_types": ["Content Type Name"], // reference content type names
            "on_quick_launch": true,
            "fields": [] // leave empty - fields come through content types
        }}
    ],
    "navigation": [
        {{
            "title": "Navigation Item",
            "url": "{{site}}/path or external URL",
            "description": "Purpose of this navigation item"
        }}
    ],
    "features": [
        {{
            "id": "feature-guid",
            "name": "Feature Name",
            "scope": "Site" // or "Web"
        }}
    ]
}}

SharePoint Best Practice Rules:
1. Design SiteFields → ContentTypes → Lists architecture
2. Create reusable site fields that multiple content types can use
3. Group related fields into logical content types
4. Associate content types with appropriate lists/libraries
5. Extract EXACT names from quotes or "called" phrases - don't modify them
6. Be specific with site titles - avoid generic names like "Team Site"
7. For document libraries, always use template_type 101
8. Choose appropriate template types: 106=Events, 107=Tasks, 105=Contacts, 104=Announcements
9. Include relevant site columns based on the context
10. Add meaningful navigation items
11. Return ONLY the JSON, no explanation text
12. Use proper SharePoint naming conventions
13. Make URLs SharePoint-friendly (no spaces, proper casing)

Content Type Design Guidelines:
- For documents: inherit from "Document" (0x0101)
- For list items: inherit from "Item" (0x01)
- Group related fields together logically
- Name content types descriptively (e.g., "Policy Document", "Employee Record", "Training Material")
- Reference existing site fields in the "fields" array

Examples of good parsing:
- "document library called 'Project Files'" -> title: "Project Files", url: "ProjectFiles"
- "HR policies site" -> site_title: "HR Policies Portal"
- "team site for marketing" -> site_title: "Marketing Team Site"
"""

        # Call GPT-4 (or fallback to GPT-3.5-turbo if GPT-4 is not available)
        models_to_try = ["gpt-4-turbo-preview", "gpt-4", "gpt-3.5-turbo"]
        
        for model in models_to_try:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": "You are a SharePoint expert who returns only valid JSON structures for PnP Provisioning Templates. Follow the exact schema provided and extract precise information from the user's description."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.1,
                    max_tokens=3000
                )
                
                print(f"✅ Using {model} for structure generation")
                break
                
            except Exception as model_error:
                print(f"⚠️  {model} not available: {model_error}")
                if model == models_to_try[-1]:  # Last model in list
                    raise model_error
                continue
        
        # Extract and parse the JSON response
        response_text = response.choices[0].message.content.strip()
        
        # Remove any markdown code blocks if present
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0].strip()
        
        # Display the cleaned JSON response from LLM
        print("📋 LLM JSON Response:")
        print("=" * 50)
        print(response_text)
        print("=" * 50)
        
        # Note: LLM output will be saved in comprehensive report
        
        # Parse JSON
        structure = json.loads(response_text)
        
        # Add UUIDs for fields that don't have them
        add_field_ids(structure)
        
        print(f"🤖 Generated structure using ChatGPT")
        return structure, response_text  # Return both structure and original JSON
        
    except json.JSONDecodeError as e:
        print(f"⚠️  Failed to parse ChatGPT response as JSON: {e}")
        print("Falling back to basic parsing...")
        return fallback_generate_structure(description)
    except Exception as e:
        print(f"⚠️  ChatGPT API error: {e}")
        print("Falling back to basic parsing...")
        return fallback_generate_structure(description)

def add_field_ids(structure: Dict[str, Any]) -> None:
    """Add UUIDs to fields and ensure proper field synchronization."""
    # Common SharePoint built-in field names to avoid
    builtin_fields = {
        'location', 'title', 'description', 'author', 'editor', 'created', 'modified',
        'id', 'version', 'name', 'url', 'path', 'type', 'size', 'status', 'category',
        'comments', 'tags', 'keywords', 'subject', 'company', 'manager', 'department',
        'priority', 'assignedto', 'duedate', 'startdate', 'percentcomplete', 'outcome',
        'contenttype', 'attachments', 'linkfilename', 'docicon', 'edit', 'folder',
        'order', 'guid', 'fileleafref', 'fileref', 'filepath', 'filesizebytes',
        'checkedoutto', 'owner', 'workflow', 'importance', 'sensitivity'
    }
    
    # Create a centralized field registry to ensure consistency
    field_registry = {}
    
    # First, collect all unique fields from all lists and site fields
    all_fields = []
    
    # Add site fields to registry
    for field in structure.get("site_fields", []):
        field_key = field["name"].lower()
        if field_key not in field_registry:
            # Sanitize field name if needed
            if field_key in builtin_fields:
                field["name"] = f"Custom{field['name']}"
                field["displayName"] = f"Custom {field['displayName']}"
            
            if "id" not in field:
                field["id"] = f"{{{str(uuid.uuid4()).upper()}}}"
            
            field_registry[field["name"]] = field
            all_fields.append(field)
    
    # Add list fields to registry and update list references
    for list_def in structure.get("lists", []):
        for field in list_def.get("fields", []):
            field_key = field["name"].lower()
            sanitized_name = field["name"]
            
            # Sanitize field name if needed
            if field_key in builtin_fields:
                sanitized_name = f"Custom{field['name']}"
                field["displayName"] = f"Custom {field['displayName']}"
            
            # Check if this field already exists in registry
            if sanitized_name in field_registry:
                # Update the list field to reference the existing site field
                field["name"] = sanitized_name
                field["id"] = field_registry[sanitized_name]["id"]
            else:
                # Create new field in registry
                field["name"] = sanitized_name
                if "id" not in field:
                    field["id"] = f"{{{str(uuid.uuid4()).upper()}}}"
                
                # Create corresponding site field
                site_field = {
                    "name": sanitized_name,
                    "displayName": field["displayName"],
                    "type": field["type"],
                    "id": field["id"]
                }
                if field.get("choices"):
                    site_field["choices"] = field["choices"]
                
                field_registry[sanitized_name] = site_field
                all_fields.append(site_field)
    
    # Update the structure with the consolidated field list
    structure["site_fields"] = all_fields
    
    # Add IDs to content types following SharePoint conventions
    for ct in structure.get("content_types", []):
        if "id" not in ct:
            # Generate proper SharePoint content type ID
            # 0x0101 = Document base, 0x01 = Item base
            base_id = "0x0101" if ct.get("parent") == "Document" else "0x01"
            # Add two more hex digits to ensure uniqueness
            unique_suffix = str(uuid.uuid4()).replace('-', '').upper()[:30]
            ct["id"] = base_id + "00" + unique_suffix

def fallback_generate_structure(description: str) -> Dict[str, Any]:
    """
    Fallback function with basic parsing when ChatGPT is not available.
    This is the original logic but simplified.
    """
    desc_lower = description.lower()
    
    # Determine site type
    site_type = "CommunicationSite"
    base_template = "SITEPAGEPUBLISHING#0"
    
    if any(keyword in desc_lower for keyword in ["team", "collaboration", "project"]):
        site_type = "TeamSite"
        # Use GROUP#0 for modern Team Sites (Microsoft 365 Groups)
        # Use STS#0 for classic team sites without O365 Group
        base_template = "GROUP#0"
    
    # Extract site title - look for quoted strings first
    site_title = "Corporate Portal"
    quoted_match = re.search(r"['\"]([^'\"]+)['\"]", description)
    if quoted_match:
        site_title = quoted_match.group(1).title()
    elif "team" in desc_lower:
        site_title = "Team Collaboration Site"
    elif "project" in desc_lower:
        site_title = "Project Management Site"
    
    # Extract document library name
    library_name = "Documents"
    called_match = re.search(r"called\s+['\"]([^'\"]+)['\"]", description)
    if called_match:
        library_name = called_match.group(1).title()
    elif "library" in desc_lower:
        lib_match = re.search(r"library\s+['\"]?([^'\".\s]+)['\"]?", desc_lower)
        if lib_match:
            library_name = lib_match.group(1).title()
    
    structure = {
        "site_type": site_type,
        "base_template": base_template,
        "site_title": site_title,
        "description": description[:100] + "..." if len(description) > 100 else description,
        "site_fields": [
            {
                "id": f"{{{str(uuid.uuid4()).upper()}}}",
                "name": "DocumentStatus",
                "displayName": "Document Status",
                "type": "Choice",
                "choices": ["Draft", "Review", "Approved"],
                "group": "Custom Columns"
            }
        ],
        "lists": [
            {
                "title": library_name,
                "template_type": 101,
                "url": library_name.replace(' ', ''),
                "enable_versioning": True,
                "enable_content_types": True,
                "fields": ["DocumentStatus"]
            }
        ],
        "navigation": [
            {"title": "Home", "url": "{site}"},
            {"title": "Documents", "url": "{site}/Shared Documents/Forms/AllItems.aspx"}
        ],
        "content_types": [],
        "features": []
    }
    
    json_content = json.dumps(structure, indent=2)
    return structure, json_content

def create_field_xml(field: Dict[str, Any]) -> str:
    """
    Create proper SharePoint field XML definition matching Microsoft samples.
    """
    field_type = field.get("type", "Text")
    
    if field_type == "Choice" and "choices" in field:
        choices_xml = ""
        default_choice = field.get("choices", [])[0] if field.get("choices") else "Option1"
        for choice in field.get("choices", ["Option1", "Option2"]):
            choices_xml += f"<CHOICE>{choice}</CHOICE>"
        
        return f'''<pnp:Field ID="{field['id']}" 
                   Type="Choice" 
                   Name="{field['name']}" 
                   StaticName="{field['name']}" 
                   DisplayName="{field['displayName']}" 
                   Group="{field.get('group', 'Custom Columns')}" 
                   Required="FALSE" 
                   ShowInDisplayForm="TRUE" 
                   ShowInEditForm="TRUE" 
                   ShowInNewForm="TRUE"
                   FillInChoice="FALSE"
                   Format="Dropdown">
  <CHOICES>
    {choices_xml}
  </CHOICES>
  <pnp:DefaultValue>{default_choice}</pnp:DefaultValue>
</pnp:Field>'''
    else:
        return f'''<pnp:Field ID="{field['id']}" 
                   Type="{field_type}" 
                   Name="{field['name']}" 
                   StaticName="{field['name']}" 
                   DisplayName="{field['displayName']}" 
                   Group="{field.get('group', 'Custom Columns')}" 
                   Required="FALSE" 
                   ShowInDisplayForm="TRUE" 
                   ShowInEditForm="TRUE" 
                   ShowInNewForm="TRUE" />'''

def structure_to_pnp_xml(structure: Dict[str, Any]) -> str:
    """
    Convert structure dictionary to PnP Provisioning XML with proper namespace handling.
    Based on official SharePoint PnP template patterns.
    """
    # Use the correct PnP namespace from official templates (updated to 2022-09)
    namespace = "http://schemas.dev.office.com/PnP/2022/09/ProvisioningSchema"
    
    # Register namespace with pnp prefix
    ET.register_namespace('pnp', namespace)
    
    # Create root elements with Microsoft-style structure
    root = ET.Element("pnp:Provisioning")
    root.set("xmlns:pnp", namespace)
    root.set("Author", "PnP Template Generator")
    root.set("Generator", "SharePoint PnP CLI Tool")
    root.set("Version", "1.0")
    root.set("Description", structure.get("description", "SharePoint provisioning template"))
    root.set("DisplayName", structure.get("site_title", "SharePoint Site Template"))
    
    # Add Templates container with meaningful ID
    templates = ET.SubElement(root, "pnp:Templates")
    templates.set("ID", "MAIN-TEMPLATES")
    
    # Create ProvisioningTemplate with Microsoft patterns
    template = ET.SubElement(templates, "pnp:ProvisioningTemplate")
    template.set("ID", "SITE-TEMPLATE")
    template.set("Version", "1")
    # Use appropriate base template based on site type
    site_type = structure.get("site_type", "team")
    base_templates = {
        "team": "GROUP#0",
        "communication": "SITEPAGEPUBLISHING#0", 
        "classic": "STS#3",
        "document": "STS#0"
    }
    template.set("BaseSiteTemplate", base_templates.get(site_type, "GROUP#0"))
    template.set("Scope", "RootSite")
    template.set("DisplayName", structure.get("site_title", "SharePoint Site Template"))
    template.set("Description", structure.get("description", "Microsoft-style SharePoint provisioning template"))
    
    # Add WebSettings with comprehensive Microsoft-style configuration
    web_settings = ET.SubElement(template, "pnp:WebSettings")
    web_settings.set("RequestAccessEmail", "")
    web_settings.set("NoCrawl", "false")
    web_settings.set("WelcomePage", "SitePages/Home.aspx")
    web_settings.set("Title", structure.get("site_title", "SharePoint Site"))
    web_settings.set("Description", structure.get("description", ""))
    web_settings.set("AlternateCSS", "")
    web_settings.set("CommentsOnSitePagesDisabled", "false")
    web_settings.set("QuickLaunchEnabled", "true")
    web_settings.set("MembersCanShare", "true")
    web_settings.set("ExcludeFromOfflineClient", "false")
    web_settings.set("DisableFlows", "false")
    web_settings.set("DisableAppViews", "false")
    
    # Add Features if any
    if structure.get("features"):
        features = ET.SubElement(template, "pnp:Features")
        site_features = ET.SubElement(features, "pnp:SiteFeatures")
        for feature_id in structure["features"]:
            feature = ET.SubElement(site_features, "pnp:Feature")
            feature.set("ID", feature_id)
    
    # Skip ContentTypes - using Microsoft pattern with list-specific fields instead
    # Microsoft's working templates show fields defined within lists without content types
    
    # Add Lists (following Microsoft SharePoint best practices - using list-specific fields)
    if structure.get("lists"):
        lists = ET.SubElement(template, "pnp:Lists")
        for list_def in structure["lists"]:
            list_instance = ET.SubElement(lists, "pnp:ListInstance")
            list_instance.set("Title", list_def["title"])
            list_instance.set("Description", list_def.get("description", f"List for managing {list_def['title'].lower()}"))
            list_instance.set("TemplateType", str(list_def["template_type"]))
            list_instance.set("Url", list_def["url"])
            
            # Add comprehensive Microsoft-style attributes
            if list_def["template_type"] == 101:  # Document Library
                list_instance.set("EnableVersioning", "true")
                list_instance.set("EnableMinorVersions", "true")
                list_instance.set("EnableModeration", "false")
                list_instance.set("MinorVersionLimit", "10")
                list_instance.set("MaxVersionLimit", "50")
                list_instance.set("DraftVersionVisibility", "1")  # Author
                list_instance.set("EnableAttachments", "false")
                list_instance.set("EnableFolderCreation", "true")
            elif list_def["template_type"] == 106:  # Events/Calendar
                list_instance.set("EnableAttachments", "false")
                list_instance.set("EnableFolderCreation", "false")
            elif list_def["template_type"] == 100:  # Custom List
                list_instance.set("EnableAttachments", "true")
                list_instance.set("EnableFolderCreation", "false")
            elif list_def["template_type"] == 105:  # Contacts
                list_instance.set("EnableAttachments", "true")
                list_instance.set("EnableFolderCreation", "false")
            elif list_def["template_type"] == 109:  # Picture Library
                list_instance.set("EnableFolderCreation", "true")
                list_instance.set("EnableAttachments", "false")
            
            # Common attributes for all lists
            list_instance.set("ContentTypesEnabled", "false")  # Set to false to use list fields directly
            list_instance.set("OnQuickLaunch", "true")
            list_instance.set("Hidden", "false")
            list_instance.set("NoCrawl", "false")
            list_instance.set("RemoveExistingContentTypes", "false")
            
            # Add list-specific fields using Microsoft pattern (inside the list, not as site fields)
            # If list has no fields specified, add all site fields to the list
            fields_to_add = list_def.get("fields", [])
            if not fields_to_add and structure.get("site_fields"):
                # Auto-add all site fields to lists that have no fields specified
                fields_to_add = [f["name"] for f in structure.get("site_fields", [])]
            
            if fields_to_add:
                fields_elem = ET.SubElement(list_instance, "pnp:Fields")
                for field_name in fields_to_add:
                    # Find the field definition in site_fields
                    field_def = next((f for f in structure.get("site_fields", []) if f["name"] == field_name), None)
                    if field_def:
                        # Create Field element (no namespace - Microsoft pattern)
                        field_elem = ET.SubElement(fields_elem, "Field")
                        
                        # Use Microsoft field pattern attributes
                        field_elem.set("Type", field_def["type"])
                        field_elem.set("DisplayName", field_def["displayName"])
                        field_elem.set("Required", "TRUE" if field_def.get("required", False) else "FALSE")
                        field_elem.set("EnforceUniqueValues", "FALSE")
                        field_elem.set("Indexed", "FALSE")
                        
                        # Ensure field ID has braces format
                        field_id = field_def["id"]
                        if not field_id.startswith("{"):
                            field_id = "{" + field_id + "}"
                        field_elem.set("ID", field_id)
                        field_elem.set("StaticName", field_def["name"])
                        field_elem.set("Name", field_def["name"])
                        
                        # Add type-specific attributes following Microsoft pattern
                        if field_def["type"] == "Text":
                            max_length = field_def.get("maxLength", 50)
                            field_elem.set("MaxLength", str(max_length))
                            if field_def.get("default"):
                                default_elem = ET.SubElement(field_elem, "Default")
                                default_elem.text = field_def["default"]
                                
                        elif field_def["type"] == "Choice" and field_def.get("choices"):
                            field_elem.set("Format", "Dropdown")
                            field_elem.set("FillInChoice", "FALSE")
                            
                            # Add default value first (Microsoft pattern)
                            default_value = field_def.get("default", field_def["choices"][0] if field_def["choices"] else "Option1")
                            default_elem = ET.SubElement(field_elem, "Default")
                            default_elem.text = default_value
                            
                            # Create CHOICES element (no namespace)
                            choices_elem = ET.SubElement(field_elem, "CHOICES")
                            for choice in field_def["choices"]:
                                choice_elem = ET.SubElement(choices_elem, "CHOICE")
                                choice_elem.text = choice
    
    # Add Navigation (following official template pattern)
    if structure.get("navigation"):
        navigation = ET.SubElement(template, "pnp:Navigation")
        navigation.set("AddNewPagesToNavigation", "true")
        navigation.set("CreateFriendlyUrlsForNewPages", "true")
        
        # Global Navigation
        global_nav = ET.SubElement(navigation, "pnp:GlobalNavigation")
        global_nav.set("NavigationType", "Structural")
        global_struct_nav = ET.SubElement(global_nav, "pnp:StructuralNavigation")
        global_struct_nav.set("RemoveExistingNodes", "true")
        
        # Current Navigation
        current_nav = ET.SubElement(navigation, "pnp:CurrentNavigation")
        current_nav.set("NavigationType", "StructuralLocal")
        current_struct_nav = ET.SubElement(current_nav, "pnp:StructuralNavigation")
        current_struct_nav.set("RemoveExistingNodes", "true")
        
        # Add navigation nodes
        for nav_item in structure["navigation"]:
            nav_node = ET.SubElement(current_struct_nav, "pnp:NavigationNode")
            nav_node.set("Title", nav_item["title"])
            
            # Sanitize URL - replace problematic values and escape XML characters
            url = nav_item["url"]
            if url == "external URL" or " " in url:
                # Replace generic placeholders with valid URLs
                url = "https://example.com"
            
            # Escape XML special characters in URL
            url = url.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")
            nav_node.set("Url", url)
    
    # Convert to string with proper formatting
    rough_string = ET.tostring(root, encoding='unicode')
    
    # Parse with minidom for pretty formatting
    reparsed = minidom.parseString(rough_string)
    
    # Get pretty XML
    xml_str = reparsed.toprettyxml(indent="  ")
    
    # Clean up the XML (remove empty lines and fix formatting)
    lines = [line for line in xml_str.split('\n') if line.strip()]
    
    # Remove the default XML declaration and add our own
    if lines[0].startswith('<?xml'):
        lines = lines[1:]
    
    # Insert proper XML declaration
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n' + '\n'.join(lines)
    
    return xml_str

def validate_xml(xml_content: str) -> bool:
    """
    Validate XML against PnP schema if available.
    Returns True if valid or no schema available.
    """
    xsd_path = Path("pnp-provisioning.xsd")
    if not xsd_path.exists():
        print("Warning: PnP Provisioning XSD not found. Skipping validation.")
        return True
    
    try:
        from lxml import etree
        
        # Parse XSD
        with open(xsd_path, 'r', encoding='utf-8') as xsd_file:
            schema_doc = etree.parse(xsd_file)
            schema = etree.XMLSchema(schema_doc)
        
        # Parse XML
        xml_doc = etree.fromstring(xml_content.encode('utf-8'))
        
        # Validate
        if schema.validate(xml_doc):
            print("✓ XML validation passed")
            return True
        else:
            print("✗ XML validation failed:")
            for error in schema.error_log:
                print(f"  Line {error.line}: {error.message}")
            return False
            
    except ImportError:
        print("Warning: lxml not available. Skipping XML validation.")
        return True
    except Exception as e:
        print(f"Warning: XML validation failed: {e}")
        return True

def save_xml(xml_content: str) -> Path:
    """
    Save XML to generated-templates directory with timestamp.
    """
    output_dir = Path("generated-templates")
    output_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"enhanced_mock_{timestamp}.xml"
    filepath = output_dir / filename
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(xml_content)
    
    return filepath

def main():
    """
    Main CLI entry point.
    """
    print("SharePoint PnP Provisioning XML Generator")
    print("=" * 45)
    
    # Load .env file if available
    if DOTENV_AVAILABLE:
        load_dotenv()
    
    # Check for OpenAI API key
    api_key_available = OPENAI_AVAILABLE and os.getenv('OPENAI_API_KEY')
    if api_key_available:
        print("🤖 ChatGPT integration: ✅ Available (using .env or environment)")
    else:
        print("🤖 ChatGPT integration: ❌ Not available (using basic parsing)")
        if not OPENAI_AVAILABLE:
            print("   Install openai package: pip install openai")
        else:
            env_file = Path(".env")
            if env_file.exists():
                print("   Add OPENAI_API_KEY to your .env file")
            else:
                print("   Create .env file with: OPENAI_API_KEY=your-api-key-here")
    
    # Get description from command line or interactively
    if len(sys.argv) > 1:
        description = " ".join(sys.argv[1:])
    else:
        print("\nExamples:")
        print("  'Create a communication site for HR policies with document library'")
        print("  'Team site for project management with a document library called \"Project Files\"'")
        print("  'Educational hub with class materials and events calendar'")
        print()
        description = input("Enter site description: ").strip()
        if not description:
            print("No description provided. Exiting.")
            sys.exit(1)
    
    print(f"\n📝 Processing: {description}")
    
    # Generate structure using ChatGPT or fallback parser
    json_content = ""
    try:
        result = llm_generate_structure(description)
        if isinstance(result, tuple):
            structure, json_content = result
        else:
            structure = result
            json_content = json.dumps(structure, indent=2)
        
        print(f"✅ Generated structure for {structure['site_type']}: '{structure['site_title']}'")
        
        # Show what was detected
        if structure.get("lists"):
            print(f"📋 Lists/Libraries: {', '.join([l['title'] for l in structure['lists']])}")
        if structure.get("site_fields"):
            print(f"📊 Site Columns: {', '.join([f['displayName'] for f in structure['site_fields']])}")
        if structure.get("navigation"):
            print(f"🧭 Navigation: {', '.join([n['title'] for n in structure['navigation']])}")
    
    except Exception as e:
        print(f"❌ Error generating structure: {e}")
        sys.exit(1)
    
    # Convert to XML
    try:
        print("\n🔨 Converting to PnP XML...")
        xml_content = structure_to_pnp_xml(structure)
    except Exception as e:
        print(f"❌ Error converting to XML: {e}")
        sys.exit(1)
    
    # XSD Validation using official PnP schema
    print("🔍 Validating XML against official PnP Schema...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    is_valid, validation_errors = validate_xml_against_xsd(xml_content)
    
    if is_valid:
        print("✅ XML is valid according to PnP Provisioning Schema 2022-09!")
    else:
        print(f"❌ XML validation failed with {len(validation_errors)} error(s):")
        for i, error in enumerate(validation_errors[:5], 1):  # Show first 5 errors
            print(f"   {i}. {error}")
        if len(validation_errors) > 5:
            print(f"   ... and {len(validation_errors) - 5} more errors")
        print("\n⚠️  Template will still be saved, but may not deploy properly to SharePoint.")
    
    # Save XML file
    try:
        filepath = save_xml(xml_content)
        print(f"💾 Saved: {filepath.absolute()}")
        print(f"📏 File size: {filepath.stat().st_size:,} bytes")
        
        # Save comprehensive report (JSON + XSD validation)
        report_file = save_comprehensive_report(json_content, str(filepath.name), is_valid, validation_errors, timestamp)
        print(f"📋 Comprehensive report saved: {report_file}")
        
        # Show success message with next steps
        print("\n🎉 Success! Next steps:")
        print("   1. Review the generated XML file")
        print("   2. Upload to SharePoint using PnP PowerShell or CLI")
        print("   3. Test the provisioning template")
        
        # Show ChatGPT setup hint if not available
        if not api_key_available:
            print("\n💡 Tip: For better parsing, create a .env file:")
            print("   OPENAI_API_KEY=your-api-key-here")
        
    except Exception as e:
        print(f"❌ Error saving file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
