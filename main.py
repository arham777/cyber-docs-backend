import os
import shutil
from docx import Document
from docx.shared import Inches  # Added import for Inches
import fitz  # PyMuPDF
from zipfile import ZipFile
import tempfile
import xml.etree.ElementTree as ET
import re
from pdf2docx import Converter  # Add this import

# Configuration
TEMPLATE_DOCX = "cybergen-template.docx"  # Your branded template
INPUT_DIR = "input"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Register XML namespaces to prevent prefix generation
namespaces = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships"
}

# Register namespaces for proper XML generation
for prefix, uri in namespaces.items():
    ET.register_namespace(prefix, uri)

def update_section_margins(sect_pr_element):
    """Update or add margin settings to a section properties element."""
    # Define standard margins (in twentieths of a point)
    margins = {
        "top": "1440",      # 1 inch = 1440 twentieths of a point
        "right": "1440",    # 1 inch right margin
        "bottom": "1440",   # 1 inch bottom margin
        "left": "1440",     # 1 inch left margin
        "header": "720",    # 0.5 inch for header
        "footer": "720"     # 0.5 inch for footer
    }
    
    # Remove existing pgMar element if present
    existing_pg_mar = sect_pr_element.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar")
    if existing_pg_mar is not None:
        sect_pr_element.remove(existing_pg_mar)
    
    # Create new pgMar element at the correct position (after any existing pgSz)
    pg_mar = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar")
    
    # Set all margin values
    for margin_type, value in margins.items():
        pg_mar.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + margin_type, value)
    
    # Insert pgMar at the correct position
    pg_sz = sect_pr_element.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz")
    if pg_sz is not None:
        pg_sz_index = list(sect_pr_element).index(pg_sz)
        sect_pr_element.insert(pg_sz_index + 1, pg_mar)
    else:
        # If no pgSz, add pgSz first with A4 dimensions
        pg_sz = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz")
        pg_sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w", "11906")  # A4 width (210mm)
        pg_sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}h", "16838")  # A4 height (297mm)
        pg_sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}orient", "portrait")
        sect_pr_element.insert(0, pg_sz)
        sect_pr_element.insert(1, pg_mar)

    # Ensure proper section type
    type_element = sect_pr_element.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type")
    if type_element is not None:
        sect_pr_element.remove(type_element)
    type_element = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type")
    type_element.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "nextPage")
    sect_pr_element.insert(0, type_element)

def apply_branding_to_docx(input_path, output_path):
    """Apply template's headers/footers to a DOCX file using low-level zip manipulation."""
    # Create temporary directories
    with tempfile.TemporaryDirectory() as temp_dir:
        template_dir = os.path.join(temp_dir, "template")
        target_dir = os.path.join(temp_dir, "target")
        
        os.makedirs(template_dir, exist_ok=True)
        os.makedirs(target_dir, exist_ok=True)
        
        # Extract both docx files (they're just zip files)
        with ZipFile(TEMPLATE_DOCX, 'r') as template_zip:
            template_zip.extractall(template_dir)
            
        with ZipFile(input_path, 'r') as target_zip:
            target_zip.extractall(target_dir)
        
        # First update the document.xml with proper margins
        target_doc_file = os.path.join(target_dir, "word", "document.xml")
        if os.path.exists(target_doc_file):
            try:
                tree = ET.parse(target_doc_file)
                root = tree.getroot()
                
                # Remove any existing sectPr elements in the document body
                for body in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body"):
                    for sect_pr in body.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr"):
                        body.remove(sect_pr)
                
                # Create new sectPr element
                body = root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
                if body is not None:
                    sect_pr = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr")
                    update_section_margins(sect_pr)
                    body.append(sect_pr)
                
                # Save the changes
                tree.write(target_doc_file, encoding="utf-8", xml_declaration=True)
                print("Updated document margins and section properties")
            except Exception as e:
                print(f"Error updating margins: {e}")
        
        # Copy critical style and theme files
        style_files = [
            "styles.xml",                # Document styles
            "theme/theme1.xml",          # Document theme
            "settings.xml",              # Document settings
            "fontTable.xml",             # Font references
            "webSettings.xml",           # Web settings
            "numbering.xml"              # Numbering definitions
        ]
        
        for style_file in style_files:
            template_file = os.path.join(template_dir, "word", style_file)
            target_file = os.path.join(target_dir, "word", style_file)
            
            if os.path.exists(template_file):
                # Ensure target directory exists
                os.makedirs(os.path.dirname(target_file), exist_ok=True)
                
                # Copy file if it doesn't exist in target
                if not os.path.exists(target_file):
                    shutil.copy2(template_file, target_file)
                    print(f"Copied style file: {style_file}")
                else:
                    # For styles.xml, merge styles instead of replacing
                    if style_file == "styles.xml":
                        try:
                            # Try to merge styles (advanced approach)
                            # This is a simplified approach - full style merging would be more complex
                            template_tree = ET.parse(template_file)
                            target_tree = ET.parse(target_file)
                            
                            template_root = template_tree.getroot()
                            target_root = target_tree.getroot()
                            
                            # Find docDefaults in template and copy it to target if missing
                            template_defaults = template_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults")
                            if template_defaults is not None:
                                target_defaults = target_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults")
                                if target_defaults is None:
                                    # Insert at the beginning of the document
                                    if len(target_root) > 0:
                                        target_root.insert(0, template_defaults)
                                        print("Added document defaults from template")
                            
                            # Look for header/footer styles in template
                            header_footer_styles = []
                            for style in template_root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style"):
                                style_id = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", "")
                                if "Header" in style_id or "Footer" in style_id:
                                    header_footer_styles.append(style)
                            
                            # Check if these styles exist in target, if not add them
                            for style in header_footer_styles:
                                style_id = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", "")
                                
                                # Check if style exists in target
                                exists = False
                                for target_style in target_root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style"):
                                    if target_style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", "") == style_id:
                                        exists = True
                                        break
                                
                                if not exists:
                                    # Find the styles element to append to
                                    styles_element = target_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styles")
                                    if styles_element is not None:
                                        styles_element.append(style)
                                        print(f"Added style: {style_id}")
                            
                            # Save the modified file
                            target_tree.write(target_file, encoding="utf-8", xml_declaration=True)
                            print("Merged style information")
                            
                        except Exception as e:
                            print(f"Error merging styles: {e}")
                            # Fall back to direct copy if merging fails
                            shutil.copy2(template_file, target_file)
                            print(f"Copied style file (fallback): {style_file}")
        
        # Copy all media files from template to target
        template_media_dir = os.path.join(template_dir, "word", "media")
        target_media_dir = os.path.join(target_dir, "word", "media")
        
        # Create target media directory if it doesn't exist
        os.makedirs(target_media_dir, exist_ok=True)
        
        # Copy all media files if template media directory exists
        media_files = []
        if os.path.exists(template_media_dir):
            for filename in os.listdir(template_media_dir):
                source_file = os.path.join(template_media_dir, filename)
                target_file = os.path.join(target_media_dir, filename)
                
                # Only copy if it's a file (not a directory)
                if os.path.isfile(source_file):
                    shutil.copy2(source_file, target_file)
                    media_files.append(filename)
                    print(f"Copied media file: {filename}")
        
        # Copy header and footer files from template to target
        header_footer_files = []
        for file_type in ["header1.xml", "header2.xml", "header3.xml", 
                          "footer1.xml", "footer2.xml", "footer3.xml"]:
            template_file = os.path.join(template_dir, "word", file_type)
            target_file = os.path.join(target_dir, "word", file_type)
            
            if os.path.exists(template_file):
                # Ensure target directory exists
                os.makedirs(os.path.dirname(target_file), exist_ok=True)
                shutil.copy2(template_file, target_file)
                header_footer_files.append(file_type)
                print(f"Copied {file_type}")
        
        # Copy all relationship files (.rels) for headers and footers
        template_rels_dir = os.path.join(template_dir, "word", "_rels")
        target_rels_dir = os.path.join(target_dir, "word", "_rels")
        
        # Create target rels directory if it doesn't exist
        os.makedirs(target_rels_dir, exist_ok=True)
        
        # Keep track of copied relationship files
        rel_files = []
        
        # Look for header and footer relationship files and copy them
        if os.path.exists(template_rels_dir):
            for filename in os.listdir(template_rels_dir):
                if "header" in filename or "footer" in filename:
                    source_file = os.path.join(template_rels_dir, filename)
                    target_file = os.path.join(target_rels_dir, filename)
                    shutil.copy2(source_file, target_file)
                    rel_files.append(filename)
                    print(f"Copied relationship file: {filename}")
        
        # Update document.xml.rels to include references to headers and footers
        template_doc_rels_file = os.path.join(template_dir, "word", "_rels", "document.xml.rels")
        target_doc_rels_file = os.path.join(target_dir, "word", "_rels", "document.xml.rels")
        
        # Create relationship mapping from template IDs to new IDs
        template_to_new_rids = {}
        
        if os.path.exists(template_doc_rels_file) and os.path.exists(target_doc_rels_file):
            try:
                # Parse the relationship files
                template_tree = ET.parse(template_doc_rels_file)
                target_tree = ET.parse(target_doc_rels_file)
                
                template_root = template_tree.getroot()
                target_root = target_tree.getroot()
                
                # Find max ID in target document
                max_rid = 0
                for rel in target_root.findall(".//{*}Relationship"):
                    rid = rel.get("Id", "")
                    if rid.startswith("rId"):
                        try:
                            num = int(rid[3:])
                            max_rid = max(max_rid, num)
                        except ValueError:
                            pass
                
                # Extract header and footer relationships from template
                header_footer_rels = []
                for rel in template_root.findall(".//{*}Relationship"):
                    target = rel.get("Target", "")
                    rel_type = rel.get("Type", "")
                    
                    # Check if this is a header/footer relationship or a media relationship used by headers/footers
                    if ("header" in target.lower() or "footer" in target.lower() or 
                        any(media_file in target for media_file in media_files)):
                        original_id = rel.get("Id", "")
                        new_id = f"rId{max_rid + 1}"
                        max_rid += 1
                        
                        # Store mapping from template ID to new ID
                        template_to_new_rids[original_id] = new_id
                        
                        # Add to relationships to be added
                        header_footer_rels.append({
                            "Id": new_id,
                            "Type": rel_type,
                            "Target": target,
                            "OriginalId": original_id
                        })
                
                # Add these relationships to the target document
                for rel_data in header_footer_rels:
                    # Create new relationship element
                    new_rel = ET.SubElement(target_root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
                    new_rel.set("Id", rel_data["Id"])
                    new_rel.set("Type", rel_data["Type"])
                    new_rel.set("Target", rel_data["Target"])
                    print(f"Added relationship: {rel_data['Id']} -> {rel_data['Target']}")
                
                # Save the modified file
                target_tree.write(target_doc_rels_file, encoding="utf-8", xml_declaration=True)
                print("Updated document relationships")
                
            except Exception as e:
                print(f"Error updating document relationships: {e}")
                # In case of error, continue with a more direct approach
                
        # Now process the header and footer relationship files
        for rel_file in rel_files:
            template_rel_path = os.path.join(template_rels_dir, rel_file)
            target_rel_path = os.path.join(target_rels_dir, rel_file)
            
            try:
                # Parse files
                template_tree = ET.parse(template_rel_path)
                if os.path.exists(target_rel_path):
                    target_tree = ET.parse(target_rel_path)
                    target_root = target_tree.getroot()
                else:
                    # Create new relationship file if it doesn't exist
                    target_root = ET.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
                    target_tree = ET.ElementTree(target_root)
                
                # Copy media references
                for rel in template_tree.findall(".//{*}Relationship"):
                    target = rel.get("Target", "")
                    rel_type = rel.get("Type", "")
                    
                    # If it's a media reference, copy it
                    if "media/" in target:
                        # Create new relationship element
                        new_rel = ET.SubElement(target_root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
                        new_rel.set("Id", rel.get("Id", ""))
                        new_rel.set("Type", rel_type)
                        new_rel.set("Target", target)
                
                # Save the modified file
                target_tree.write(target_rel_path, encoding="utf-8", xml_declaration=True)
                print(f"Updated relationship file: {rel_file}")
                
            except Exception as e:
                print(f"Error updating relationship file {rel_file}: {e}")
                # In case of error, copy the entire file
                if os.path.exists(template_rel_path):
                    shutil.copy2(template_rel_path, target_rel_path)
                    print(f"Copied relationship file as fallback: {rel_file}")
        
        # Now update the document.xml to reference headers and footers
        template_doc_file = os.path.join(template_dir, "word", "document.xml")
        target_doc_file = os.path.join(target_dir, "word", "document.xml")
        
        if os.path.exists(template_doc_file) and os.path.exists(target_doc_file):
            try:
                # Read the files
                with open(template_doc_file, 'r', encoding='utf-8') as f:
                    template_content = f.read()
                
                with open(target_doc_file, 'r', encoding='utf-8') as f:
                    target_content = f.read()
                
                # Extract header and footer references
                ref_pattern = r'<w:(header|footer)Reference[^>]*?r:id="([^"]+)"[^>]*?/>'
                header_footer_refs = []
                
                for match in re.finditer(ref_pattern, template_content):
                    ref_type = match.group(1)  # header or footer
                    old_rid = match.group(2)   # rId value
                    
                    if old_rid in template_to_new_rids:
                        new_rid = template_to_new_rids[old_rid]
                        ref_element = match.group(0).replace(f'r:id="{old_rid}"', f'r:id="{new_rid}"')
                        header_footer_refs.append(ref_element)
                        print(f"Mapped header/footer reference: {old_rid} -> {new_rid}")
                    else:
                        # If we don't have a mapping, keep the original ID
                        header_footer_refs.append(match.group(0))
                
                # Find the section properties in the target
                sect_pr_pattern = r'<w:sectPr[^>]*>.*?</w:sectPr>'
                target_sect_pr_match = re.search(sect_pr_pattern, target_content, re.DOTALL)
                
                if target_sect_pr_match:
                    # If target has sectPr, we'll modify it
                    sect_pr_start = target_sect_pr_match.start()
                    sect_pr_end = target_sect_pr_match.end()
                    
                    # Extract the section properties element
                    current_sect_pr = target_sect_pr_match.group(0)
                    
                    # Remove any existing header/footer references
                    cleaned_sect_pr = re.sub(r'<w:(header|footer)Reference[^>]*?/>', '', current_sect_pr)
                    
                    # Add our header/footer references just before the closing tag
                    insert_point = cleaned_sect_pr.rfind('</w:sectPr>')
                    new_sect_pr = cleaned_sect_pr[:insert_point] + ''.join(header_footer_refs) + cleaned_sect_pr[insert_point:]
                    
                    # Replace in the full document
                    new_target_content = target_content[:sect_pr_start] + new_sect_pr + target_content[sect_pr_end:]
                    print("Updated existing section properties with header/footer references")
                else:
                    # If no sectPr in target, extract from template
                    template_sect_pr_match = re.search(sect_pr_pattern, template_content, re.DOTALL)
                    
                    if template_sect_pr_match:
                        # Extract the complete sectPr element from template
                        template_sect_pr = template_sect_pr_match.group(0)
                        
                        # Replace any template rIds with new ones
                        for old_rid, new_rid in template_to_new_rids.items():
                            template_sect_pr = template_sect_pr.replace(f'r:id="{old_rid}"', f'r:id="{new_rid}"')
                        
                        # Add the template sectPr before </w:body>
                        body_end = target_content.rfind('</w:body>')
                        if body_end != -1:
                            new_target_content = target_content[:body_end] + template_sect_pr + target_content[body_end:]
                            print("Added template section properties with header/footer references")
                        else:
                            # If no body end tag found, this is likely not a valid Word document
                            new_target_content = target_content
                            print("Warning: Could not find </w:body> tag in target document")
                    else:
                        # If no sectPr in template either, create a simple sectPr with header/footer references
                        new_sect_pr = f'<w:sectPr>{"".join(header_footer_refs)}</w:sectPr>'
                        
                        # Add before </w:body>
                        body_end = target_content.rfind('</w:body>')
                        if body_end != -1:
                            new_target_content = target_content[:body_end] + new_sect_pr + target_content[body_end:]
                            print("Added new section properties with header/footer references")
                        else:
                            new_target_content = target_content
                            print("Warning: Could not find </w:body> tag in target document")
                
                # Write the modified document
                with open(target_doc_file, 'w', encoding='utf-8') as f:
                    f.write(new_target_content)
                
                print("Updated document with header/footer references")
                
            except Exception as e:
                print(f"Error updating document structure: {e}")
                # Fall back to a more direct approach
                try:
                    # Extract the entire sectPr element from template
                    sect_pr_match = re.search(r'<w:sectPr[^>]*>.*?</w:sectPr>', template_content, re.DOTALL)
                    
                    if sect_pr_match:
                        template_sect_pr = sect_pr_match.group(0)
                        
                        # Update the rIds
                        for old_rid, new_rid in template_to_new_rids.items():
                            template_sect_pr = template_sect_pr.replace(f'r:id="{old_rid}"', f'r:id="{new_rid}"')
                        
                        # Replace or add to the target document
                        target_sect_pr_match = re.search(r'<w:sectPr[^>]*>.*?</w:sectPr>', target_content, re.DOTALL)
                        
                        if target_sect_pr_match:
                            # Replace existing sectPr
                            new_target_content = target_content.replace(target_sect_pr_match.group(0), template_sect_pr)
                        else:
                            # Add before </w:body>
                            body_end = target_content.rfind('</w:body>')
                            if body_end != -1:
                                new_target_content = target_content[:body_end] + template_sect_pr + target_content[body_end:]
                            else:
                                new_target_content = target_content
                        
                        with open(target_doc_file, 'w', encoding='utf-8') as f:
                            f.write(new_target_content)
                        
                        print("Updated document structure using fallback method")
                    
                except Exception as e:
                    print(f"Error in fallback method: {e}")
        
        # Fix content types for headers and footers
        content_types_file = os.path.join(target_dir, "[Content_Types].xml")
        template_types_file = os.path.join(template_dir, "[Content_Types].xml")
        
        try:
            if os.path.exists(content_types_file) and os.path.exists(template_types_file):
                template_types_tree = ET.parse(template_types_file)
                target_types_tree = ET.parse(content_types_file)
                
                template_types_root = template_types_tree.getroot()
                target_types_root = target_types_tree.getroot()
                
                # Find all header/footer content types in template
                for override in template_types_root.findall(".//{*}Override"):
                    part_name = override.get("PartName", "")
                    
                    if "header" in part_name.lower() or "footer" in part_name.lower() or "media" in part_name.lower():
                        content_type = override.get("ContentType", "")
                        
                        # Check if this part already exists in target
                        exists = False
                        for target_override in target_types_root.findall(".//{*}Override"):
                            if target_override.get("PartName", "") == part_name:
                                exists = True
                                break
                        
                        if not exists:
                            # Add the content type override
                            new_override = ET.SubElement(target_types_root, "{http://schemas.openxmlformats.org/package/2006/content-types}Override")
                            new_override.set("PartName", part_name)
                            new_override.set("ContentType", content_type)
                            print(f"Added content type override for: {part_name}")
                
                # Make sure image relationships are defined
                image_types = {
                    "png": "image/png",
                    "jpg": "image/jpeg",
                    "jpeg": "image/jpeg",
                    "gif": "image/gif",
                    "bmp": "image/bmp",
                    "tif": "image/tiff",
                    "tiff": "image/tiff",
                    "wmf": "image/x-wmf"
                }
                
                # Check if default content types for images exist
                for ext, content_type in image_types.items():
                    exists = False
                    for default in target_types_root.findall(".//{*}Default"):
                        if default.get("Extension", "") == ext:
                            exists = True
                            break
                    
                    if not exists:
                        # Check if it exists in template
                        for default in template_types_root.findall(".//{*}Default"):
                            if default.get("Extension", "") == ext:
                                # Copy from template
                                new_default = ET.SubElement(target_types_root, "{http://schemas.openxmlformats.org/package/2006/content-types}Default")
                                new_default.set("Extension", ext)
                                new_default.set("ContentType", content_type)
                                print(f"Added default content type for: .{ext}")
                                break
                
                target_types_tree.write(content_types_file, encoding="utf-8", xml_declaration=True)
                print("Updated content types")
                
        except Exception as e:
            print(f"Error updating content types: {e}")
            # Fall back to copying the entire content types file
            if os.path.exists(template_types_file):
                shutil.copy2(template_types_file, content_types_file)
                print("Copied content types file as fallback")
        
        # Create new docx file (zip) from the modified directory
        try:
            shutil.make_archive(os.path.splitext(output_path)[0], 'zip', target_dir)
            
            # Rename zip to docx
            zip_path = f"{os.path.splitext(output_path)[0]}.zip"
            if os.path.exists(zip_path):
                # Make sure the output directory exists
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                # Remove existing output file if it exists
                if os.path.exists(output_path):
                    os.remove(output_path)
                os.rename(zip_path, output_path)
                print(f"Created branded document: {output_path}")
        except Exception as e:
            print(f"Error creating final document: {e}")
            # Try with an alternative approach
            try:
                # Create zip file with a different name to avoid conflicts
                alt_zip_path = f"{os.path.splitext(output_path)[0]}_alt.zip"
                with ZipFile(alt_zip_path, 'w') as zipf:
                    for root, dirs, files in os.walk(target_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, target_dir)
                            zipf.write(file_path, arcname)
                
                # Rename to docx
                if os.path.exists(alt_zip_path):
                    alt_output_path = f"{os.path.splitext(output_path)[0]}_alt.docx"
                    os.rename(alt_zip_path, alt_output_path)
                    print(f"Created alternative branded document: {alt_output_path}")
            except Exception as alt_e:
                print(f"Error with alternative approach: {alt_e}")

def convert_pdf_to_docx(pdf_path, docx_path):
    """Convert PDF file to DOCX format."""
    try:
        # Create a PDF converter object
        cv = Converter(pdf_path)
        # Convert PDF to DOCX with A4 settings
        cv.convert(docx_path, start=0, end=None, pages=None, 
                  layout_kwargs={
                      'margin_top': 72.0,      # 1 inch
                      'margin_bottom': 72.0,   # 1 inch
                      'margin_left': 72.0,     # 1 inch
                      'margin_right': 72.0,    # 1 inch
                      'header_height': 36.0,   # 0.5 inch for header
                      'footer_height': 36.0    # 0.5 inch for footer
                  })
        cv.close()
        
        try:
            doc = Document(docx_path)
            for section in doc.sections:
                section.different_first_page_header_footer = False
                section.header_distance = Inches(0.5)
                section.footer_distance = Inches(0.5)
            doc.save(docx_path)
        except Exception as e:
            print(f"Warning: Could not optimize converted document: {e}")
        
        return True
    except Exception as e:
        print(f"Error converting PDF to DOCX: {e}")
        return False

def batch_process():
    """Process all files in input directory."""
    for filename in os.listdir(INPUT_DIR):
        input_path = os.path.join(INPUT_DIR, filename)
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        # Check if output file exists - create a unique name if needed
        temp_output_path = output_path
        if os.path.exists(output_path):
            base, ext = os.path.splitext(output_path)
            temp_output_path = f"{base}_temp{ext}"
        
        try:
            if filename.endswith(".docx"):
                print(f"Processing DOCX: {filename}")
                try:
                    apply_branding_to_docx(input_path, temp_output_path)
                    
                    # Try to rename the temp file to the final output if needed
                    if temp_output_path != output_path:
                        try:
                            # First try to remove the existing file
                            if os.path.exists(output_path):
                                os.remove(output_path)
                            os.rename(temp_output_path, output_path)
                        except (PermissionError, OSError) as e:
                            print(f"Could not rename to final output: {e}")
                            print(f"Output saved as: {temp_output_path}")
                            
                    print(f"Completed: {filename}")
                except Exception as e:
                    print(f"Error processing DOCX {filename}: {e}")
                    # If branding fails, copy the original file to output
                    print(f"Copying original file to output as fallback")
                    try:
                        temp_fallback = f"{os.path.splitext(temp_output_path)[0]}_fallback{os.path.splitext(temp_output_path)[1]}"
                        shutil.copy2(input_path, temp_fallback)
                        print(f"Fallback saved as: {temp_fallback}")
                    except Exception as copy_error:
                        print(f"Error creating fallback: {copy_error}")
            elif filename.endswith(".pdf"):
                print(f"Processing PDF: {filename}")
                try:
                    # First convert PDF to DOCX
                    temp_docx = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_converted.docx")
                    if convert_pdf_to_docx(input_path, temp_docx):
                        # Now apply branding to the converted DOCX
                        output_docx = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}.docx")
                        apply_branding_to_docx(temp_docx, output_docx)
                        
                        # Clean up temporary DOCX file
                        if os.path.exists(temp_docx):
                            os.remove(temp_docx)
                            
                        print(f"Completed: {filename} (converted to DOCX)")
                    else:
                        print(f"Failed to convert PDF: {filename}")
                        # Create fallback copy
                        temp_fallback = f"{os.path.splitext(temp_output_path)[0]}_fallback.pdf"
                        shutil.copy2(input_path, temp_fallback)
                        print(f"Fallback saved as: {temp_fallback}")
                except Exception as e:
                    print(f"Error processing PDF {filename}: {e}")
                    try:
                        temp_fallback = f"{os.path.splitext(temp_output_path)[0]}_fallback.pdf"
                        shutil.copy2(input_path, temp_fallback)
                        print(f"Fallback saved as: {temp_fallback}")
                    except Exception as copy_error:
                        print(f"Error creating fallback: {copy_error}")
            else:
                print(f"Skipping unsupported file: {filename}")
        except Exception as e:
            print(f"Error processing {filename}: {e}")

if __name__ == "__main__":
    batch_process()
    print(f"Processing complete. Files saved to: {OUTPUT_DIR}")