import zipfile
import re
import os
import xml.etree.ElementTree as ET

docx_path = 'Investors.docx'

def get_image_mapping():
    rels = {}
    images_in_order = []
    
    with zipfile.ZipFile(docx_path, 'r') as docx:
        # 1. Parse _rels to map rId -> filename
        with docx.open('word/_rels/document.xml.rels') as rels_file:
            tree = ET.parse(rels_file)
            root = tree.getroot()
            for child in root:
                if 'image' in child.attrib.get('Type', ''):
                    rels[child.attrib.get('Id')] = child.attrib.get('Target').split('/')[-1]

        # 2. Parse document.xml for text and images in order
        with docx.open('word/document.xml') as doc_file:
            xml_content = doc_file.read().decode('utf-8')
            
            # Remove namespaces for easier parsing (hacky but works for simple extraction)
            xml_content = re.sub(r' xmlns:[^=]*="[^"]*"', '', xml_content)
            
            # Find all text and blips
            # We'll just tokenize by <w:t>text</w:t> and <a:blip r:embed="rIdX"/>
            tokens = re.split(r'(<w:t>.*?</w:t>|<a:blip r:embed=".*?/>)', xml_content)
            
            current_text = []
            results = []
            
            for token in tokens:
                if token.startswith('<w:t>'):
                    text = re.sub(r'<[^>]+>', '', token).strip()
                    if text:
                        current_text.append(text)
                        if len(current_text) > 5: # Keep context small
                            current_text.pop(0)
                elif token.startswith('<a:blip'):
                    match = re.search(r'r:embed="([^"]+)"', token)
                    if match:
                        rid = match.group(1)
                        if rid in rels:
                            img_filename = rels[rid]
                            # Associate with recent text
                            context = " ".join(current_text)
                            results.append((context, img_filename))
                            # Clear text after finding an image to avoid overlap
                            current_text = [] 
                            
            return results

mapping = get_image_mapping()

# Generate JS file content
js_content = "const INVESTOR_IMAGES = {\n"
for ctx, img in mapping:
    # Clean context to be a safe key (it's the Rep Name usually)
    # The context from the XML parsing might be "Rep Name" or "Rep Name Description"
    # We will trust the extraction for now, but strip whitespace.
    clean_name = ctx.strip()
    js_content += f'  "{clean_name}": "assets/investors/{img}",\n'
js_content += "};\n"

output_path = 'assets/investor_images.js'
with open(output_path, 'w') as f:
    f.write(js_content)

print(f"Generated {output_path} with {len(mapping)} images.")
