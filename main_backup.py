import os
import shutil
from docx import Document
from docx.shared import Inches
import fitz  # PyMuPDF
from zipfile import ZipFile
import tempfile
import xml.etree.ElementTree as ET
import re
from pdf2docx import Converter

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

def apply_branding_to_docx(input_path, output_path):
    """Apply template's headers/footers to a DOCX file using low-level zip manipulation."""
    try:
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
            
            # Copy header and footer files from template to target
            template_word_dir = os.path.join(template_dir, "word")
            target_word_dir = os.path.join(target_dir, "word")
            
            # Create target directories if they don't exist
            os.makedirs(os.path.join(target_word_dir, "_rels"), exist_ok=True)
            
            # Copy header and footer files and their relationships
            header_footer_files = []
            for file_type in ["header1.xml", "header2.xml", "header3.xml", 
                            "footer1.xml", "footer2.xml", "footer3.xml"]:
                template_file = os.path.join(template_word_dir, file_type)
                target_file = os.path.join(target_word_dir, file_type)
                
                if os.path.exists(template_file):
                    shutil.copy2(template_file, target_file)
                    header_footer_files.append(file_type)
                    print(f"Copied {file_type}")
                    
                    # Also copy relationship files if they exist
                    rel_file = f"{file_type}.rels"
                    template_rel = os.path.join(template_word_dir, "_rels", rel_file)
                    target_rel = os.path.join(target_word_dir, "_rels", rel_file)
                    
                    if os.path.exists(template_rel):
                        shutil.copy2(template_rel, target_rel)
                        print(f"Copied {rel_file}")
            
            # Copy media files from template to target
            template_media = os.path.join(template_word_dir, "media")
            target_media = os.path.join(target_word_dir, "media")
            
            os.makedirs(target_media, exist_ok=True)
            
            if os.path.exists(template_media):
                for media_file in os.listdir(template_media):
                    if os.path.isfile(os.path.join(template_media, media_file)):
                        shutil.copy2(
                            os.path.join(template_media, media_file),
                            os.path.join(target_media, media_file)
                        )
                        print(f"Copied media file: {media_file}")
            
            # Update document.xml.rels to include references to headers and footers
            template_doc_rels = os.path.join(template_word_dir, "_rels", "document.xml.rels")
            target_doc_rels = os.path.join(target_word_dir, "_rels", "document.xml.rels")
            
            if os.path.exists(template_doc_rels) and os.path.exists(target_doc_rels):
                # Load relationship files
                template_rels_tree = ET.parse(template_doc_rels)
                target_rels_tree = ET.parse(target_doc_rels)
                
                template_rels = template_rels_tree.getroot()
                target_rels = target_rels_tree.getroot()
                
                # Find existing relationship IDs in target
                existing_ids = []
                for rel in target_rels.findall(".//{*}Relationship"):
                    existing_ids.append(rel.get("Id", ""))
                
                # Find max ID in target
                max_id = 0
                for rel_id in existing_ids:
                    if rel_id.startswith("rId"):
                        try:
                            id_num = int(rel_id[3:])
                            max_id = max(max_id, id_num)
                        except ValueError:
                            pass
                
                # Track mapping of old template IDs to new IDs
                id_mapping = {}
                
                # Copy header and footer relationships from template
                for rel in template_rels.findall(".//{*}Relationship"):
                    target = rel.get("Target", "")
                    rel_type = rel.get("Type", "")
                    
                    if any(hf in target for hf in ["header", "footer"]):
                        original_id = rel.get("Id", "")
                        
                        # Generate a new unique ID
                        max_id += 1
                        new_id = f"rId{max_id}"
                        
                        # Store mapping
                        id_mapping[original_id] = new_id
                        
                        # Create new relationship in target
                        new_rel = ET.SubElement(target_rels, 
                                            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
                        new_rel.set("Id", new_id)
                        new_rel.set("Type", rel_type)
                        new_rel.set("Target", target)
                        
                        print(f"Added relationship: {new_id} -> {target}")
                
                # Save updated relationships
                target_rels_tree.write(target_doc_rels, encoding="utf-8", xml_declaration=True)
                print("Updated document relationships")
                
                # Now update document.xml to reference headers and footers
                template_doc = os.path.join(template_word_dir, "document.xml")
                target_doc = os.path.join(target_word_dir, "document.xml")
                
                if os.path.exists(template_doc) and os.path.exists(target_doc):
                    # Extract section properties from template
                    template_tree = ET.parse(template_doc)
                    template_root = template_tree.getroot()
                    
                    # Extract section properties from target
                    target_tree = ET.parse(target_doc)
                    target_root = target_tree.getroot()
                    
                    # Get template sectPr
                    template_sectPr = template_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr")
                    
                    if template_sectPr is not None:
                        # Find or create sectPr in target
                        target_body = target_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
                        if target_body is not None:
                            # Remove any existing sectPr
                            for existing_sectPr in target_body.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr"):
                                target_body.remove(existing_sectPr)
                            
                            # Create new sectPr
                            new_sectPr = ET.SubElement(target_body, 
                                                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr")
                            
                            # Copy page size and margins
                            for element_type in ["pgSz", "pgMar", "cols", "docGrid"]:
                                element = template_sectPr.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + element_type)
                                if element is not None:
                                    # Create a deep copy
                                    new_element = ET.Element(element.tag)
                                    for key, value in element.attrib.items():
                                        new_element.set(key, value)
                                    
                                    # Special handling for pgMar (ensure footer is correctly positioned)
                                    if element_type == "pgMar":
                                        # Set standard margins
                                        margins = {
                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top": "1440",
                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right": "1440",
                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom": "1440",
                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left": "1440",
                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}header": "720",
                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footer": "720"
                                        }
                                        
                                        for key, value in margins.items():
                                            new_element.set(key, value)
                                    
                                    new_sectPr.append(new_element)
                            
                            # Copy header and footer references with updated IDs
                            for ref_type in ["headerReference", "footerReference"]:
                                for ref in template_sectPr.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + ref_type):
                                    old_id = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                                    
                                    if old_id in id_mapping:
                                        new_ref = ET.SubElement(new_sectPr, 
                                                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}" + ref_type)
                                        
                                        # Copy all attributes
                                        for key, value in ref.attrib.items():
                                            if key == "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id":
                                                new_ref.set(key, id_mapping[old_id])
                                            else:
                                                new_ref.set(key, value)
                                        
                                        print(f"Added {ref_type} with ID: {id_mapping[old_id]}")
                            
                            # Save target document
                            target_tree.write(target_doc, encoding="utf-8", xml_declaration=True)
                            print("Updated document structure with header/footer references")
            
            # Update [Content_Types].xml
            template_types = os.path.join(template_dir, "[Content_Types].xml")
            target_types = os.path.join(target_dir, "[Content_Types].xml")
            
            if os.path.exists(template_types) and os.path.exists(target_types):
                template_types_tree = ET.parse(template_types)
                target_types_tree = ET.parse(target_types)
                
                template_types_root = template_types_tree.getroot()
                target_types_root = target_types_tree.getroot()
                
                # Find content types for headers and footers in template
                content_types = {}
                for override in template_types_root.findall(".//{*}Override"):
                    part_name = override.get("PartName", "")
                    
                    if any(hf in part_name for hf in ["header", "footer"]):
                        content_type = override.get("ContentType", "")
                        content_types[part_name] = content_type
                
                # Add these content types to target if they don't exist
                for part_name, content_type in content_types.items():
                    # Check if it already exists
                    exists = False
                    for override in target_types_root.findall(".//{*}Override"):
                        if override.get("PartName", "") == part_name:
                            exists = True
                            break
                    
                    if not exists:
                        new_override = ET.SubElement(target_types_root, 
                                                "{http://schemas.openxmlformats.org/package/2006/content-types}Override")
                        new_override.set("PartName", part_name)
                        new_override.set("ContentType", content_type)
                        print(f"Added content type for: {part_name}")
                
                # Add default content types for image formats if needed
                image_types = {
                    "png": "image/png",
                    "jpg": "image/jpeg",
                    "jpeg": "image/jpeg",
                    "gif": "image/gif",
                    "bmp": "image/bmp"
                }
                
                for ext, content_type in image_types.items():
                    # Check if it exists
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
                                new_default = ET.SubElement(target_types_root, 
                                                        "{http://schemas.openxmlformats.org/package/2006/content-types}Default")
                                new_default.set("Extension", ext)
                                new_default.set("ContentType", content_type)
                                print(f"Added default content type for: .{ext}")
                                break
                
                # Save updated content types
                target_types_tree.write(target_types, encoding="utf-8", xml_declaration=True)
                print("Updated content types")
            
            # Create output docx file
            shutil.make_archive(os.path.splitext(output_path)[0], 'zip', target_dir)
            
            # Rename zip to docx
            zip_path = f"{os.path.splitext(output_path)[0]}.zip"
            if os.path.exists(zip_path):
                # Make sure output directory exists
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                # Remove existing output file if it exists
                if os.path.exists(output_path):
                    os.remove(output_path)
                
                os.rename(zip_path, output_path)
                print(f"Created branded document: {output_path}")
        
        return True
    except Exception as e:
        print(f"Error applying branding: {e}")
        return False

def convert_pdf_to_docx(pdf_path, docx_path):
    """Convert PDF file to DOCX format."""
    try:
        # Create a PDF converter object
        cv = Converter(pdf_path)
        # Convert PDF to DOCX with default settings (more reliable)
        cv.convert(docx_path)
        cv.close()
        
        # Try to optimize using python-docx
        try:
            doc = Document(docx_path)
            for section in doc.sections:
                # Use conservative margin settings
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
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
        
        try:
            if filename.endswith(".docx"):
                print(f"Processing DOCX: {filename}")
                output_path = os.path.join(OUTPUT_DIR, filename)
                if apply_branding_to_docx(input_path, output_path):
                    print(f"Completed: {filename}")
                else:
                    # If branding fails, create a fallback copy
                    fallback = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_fallback.docx")
                    shutil.copy2(input_path, fallback)
                    print(f"Created fallback copy: {os.path.basename(fallback)}")
            
            elif filename.endswith(".pdf"):
                print(f"Processing PDF: {filename}")
                # Create temporary file for conversion
                with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
                    temp_docx = tmp.name
                
                try:
                    # Convert PDF to DOCX
                    if convert_pdf_to_docx(input_path, temp_docx):
                        # Apply branding
                        output_docx = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}.docx")
                        if apply_branding_to_docx(temp_docx, output_docx):
                            print(f"Completed: {filename} (converted to DOCX)")
                        else:
                            print(f"Branding failed for converted PDF: {filename}")
                            # Create fallback docx
                            fallback = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_converted_only.docx")
                            shutil.copy2(temp_docx, fallback)
                            print(f"Created fallback converted file: {os.path.basename(fallback)}")
                    else:
                        print(f"PDF conversion failed: {filename}")
                        # Create fallback PDF copy
                        fallback = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_fallback.pdf")
                        shutil.copy2(input_path, fallback)
                        print(f"Created fallback PDF copy: {os.path.basename(fallback)}")
                finally:
                    # Clean up temp file
                    if os.path.exists(temp_docx):
                        try:
                            os.remove(temp_docx)
                        except:
                            pass
            else:
                print(f"Skipping unsupported file: {filename}")
        
        except Exception as e:
            print(f"Error processing {filename}: {e}")
            try:
                # Create fallback copy
                fallback = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_fallback{os.path.splitext(filename)[1]}")
                shutil.copy2(input_path, fallback)
                print(f"Created fallback copy: {os.path.basename(fallback)}")
            except:
                pass

if __name__ == "__main__":
    batch_process()
    print(f"Processing complete. Files saved to: {OUTPUT_DIR}")
