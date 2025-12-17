"""
Direct .docx manipulation for updating DOCVARIABLE fields without opening Word.
"""

import zipfile
import shutil
import tempfile
import os
import re
from lxml import etree
from typing import Optional


# XML namespaces used in Word documents
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}


def update_docx_variables(docx_path: str, variables: dict[str, str], backup: bool = True) -> bool:
    """
    Update DOCVARIABLE fields in a closed .docx file.

    Args:
        docx_path: Path to the .docx file
        variables: Dict of variable_name -> new_value
        backup: If True, create a .bak backup before modifying

    Returns:
        True if successful, False otherwise
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"File not found: {docx_path}")

    if not docx_path.lower().endswith('.docx'):
        raise ValueError("File must be a .docx file")

    # Create backup if requested
    if backup:
        backup_path = docx_path + '.bak'
        shutil.copy2(docx_path, backup_path)

    # Create a temporary directory for extraction
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract the docx
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Update settings.xml (document variables)
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        if os.path.exists(settings_path):
            _update_settings_xml(settings_path, variables)

        # Update document.xml (field display values)
        document_path = os.path.join(temp_dir, 'word', 'document.xml')
        if os.path.exists(document_path):
            _update_document_xml(document_path, variables)

        # Repack the docx
        _repack_docx(temp_dir, docx_path)

    return True


def _update_settings_xml(settings_path: str, variables: dict[str, str]):
    """Update document variables in settings.xml."""
    tree = etree.parse(settings_path)
    root = tree.getroot()

    # Find all docVar elements
    for doc_var in root.findall('.//w:docVar', NAMESPACES):
        var_name = doc_var.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
        if var_name in variables:
            doc_var.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', variables[var_name])

    # Write back
    tree.write(settings_path, xml_declaration=True, encoding='UTF-8', standalone='yes')


def _update_document_xml(document_path: str, variables: dict[str, str]):
    """Update field display values in document.xml."""
    tree = etree.parse(document_path)
    root = tree.getroot()

    # Find all fldSimple elements (simple field codes)
    for fld_simple in root.findall('.//w:fldSimple', NAMESPACES):
        instr = fld_simple.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr', '')

        # Check if this is a DOCVARIABLE field
        match = re.search(r'DOCVARIABLE\s+(\S+)', instr)
        if match:
            var_name = match.group(1).strip('"')
            if var_name in variables:
                # Find the text element and update it
                for text_elem in fld_simple.findall('.//w:t', NAMESPACES):
                    text_elem.text = variables[var_name]

    # Also handle complex field codes (w:fldChar based)
    # These are more complex and look like:
    # <w:fldChar w:fldCharType="begin"/>
    # <w:instrText> DOCVARIABLE varname </w:instrText>
    # <w:fldChar w:fldCharType="separate"/>
    # <w:t>value</w:t>
    # <w:fldChar w:fldCharType="end"/>

    current_var_name = None
    in_field = False
    after_separate = False

    for elem in root.iter():
        tag = etree.QName(elem.tag).localname if elem.tag else ''

        if tag == 'fldChar':
            fld_type = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', '')
            if fld_type == 'begin':
                in_field = True
                current_var_name = None
                after_separate = False
            elif fld_type == 'separate':
                after_separate = True
            elif fld_type == 'end':
                in_field = False
                current_var_name = None
                after_separate = False

        elif tag == 'instrText' and in_field:
            instr = elem.text or ''
            match = re.search(r'DOCVARIABLE\s+(\S+)', instr)
            if match:
                current_var_name = match.group(1).strip('"')

        elif tag == 't' and in_field and after_separate and current_var_name:
            if current_var_name in variables:
                elem.text = variables[current_var_name]

    # Write back
    tree.write(document_path, xml_declaration=True, encoding='UTF-8', standalone='yes')


def _repack_docx(temp_dir: str, output_path: str):
    """Repack the extracted files into a .docx file."""
    # Remove the old file
    if os.path.exists(output_path):
        os.remove(output_path)

    # Create new zip with proper compression
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zipf.write(file_path, arcname)


def get_docx_variables(docx_path: str) -> dict[str, str]:
    """
    Read all document variables from a .docx file without opening Word.

    Args:
        docx_path: Path to the .docx file

    Returns:
        Dict of variable_name -> value
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"File not found: {docx_path}")

    variables = {}

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        try:
            with zip_ref.open('word/settings.xml') as settings_file:
                tree = etree.parse(settings_file)
                root = tree.getroot()

                for doc_var in root.findall('.//w:docVar', NAMESPACES):
                    name = doc_var.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
                    val = doc_var.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    if name:
                        variables[name] = val or ''
        except KeyError:
            pass  # No settings.xml

    return variables


def get_docx_field_names(docx_path: str) -> list[str]:
    """
    Get list of DOCVARIABLE field names used in the document.

    Args:
        docx_path: Path to the .docx file

    Returns:
        List of variable names found in DOCVARIABLE fields
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"File not found: {docx_path}")

    field_names = []

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        try:
            with zip_ref.open('word/document.xml') as doc_file:
                content = doc_file.read().decode('utf-8')

                # Find DOCVARIABLE references in field codes
                matches = re.findall(r'DOCVARIABLE\s+(\S+)', content)
                for match in matches:
                    var_name = match.strip('"')
                    if var_name and var_name not in field_names:
                        field_names.append(var_name)
        except KeyError:
            pass  # No document.xml

    return field_names


# Test function
if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Usage: python docx_updater.py <docx_file>")
        sys.exit(1)

    docx_path = sys.argv[1]

    print(f"Reading: {docx_path}")
    print("\nDocument Variables:")
    variables = get_docx_variables(docx_path)
    for name, val in variables.items():
        print(f"  {name}: {val}")

    print("\nDOCVARIABLE Fields Used:")
    fields = get_docx_field_names(docx_path)
    for name in fields:
        print(f"  {name}")
