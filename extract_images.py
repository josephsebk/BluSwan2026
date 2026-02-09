import zipfile
import re
import os
from PIL import Image
import io
import xml.etree.ElementTree as ET

docx_path = 'Investors.docx'
output_dir = 'assets/investors'
os.makedirs(output_dir, exist_ok=True)

def process_image(image_data, filename):
    try:
        img = Image.open(io.BytesIO(image_data))
        # Resize to max 200x200
        img.thumbnail((200, 200))
        # Save as WebP
        webp_filename = os.path.splitext(filename)[0] + '.webp'
        save_path = os.path.join(output_dir, webp_filename)
        img.save(save_path, 'WEBP', quality=80)
        return webp_filename
    except Exception as e:
        print(f"Error processing {filename}: {e}")
        return None

# Dictionary to map rId to image filename
rels = {}
with zipfile.ZipFile(docx_path, 'r') as docx:
    # 1. Parse relationships to map rId to filename
    try:
        with docx.open('word/_rels/document.xml.rels') as rels_file:
            tree = ET.parse(rels_file)
            root = tree.getroot()
            ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            for child in root:
                if 'image' in child.attrib.get('Type', ''):
                    rId = child.attrib.get('Id')
                    target = child.attrib.get('Target')
                    # Target is relative to word/ directory, usually "media/image1.png"
                    rels[rId] = target.split('/')[-1] # Just filename
    except Exception as e:
        print(f"Error parsing rels: {e}")

    # 2. Extract and process all images
    full_image_list = []
    for file in docx.namelist():
        if file.startswith('word/media/'):
            with docx.open(file) as f:
                img_data = f.read()
                filename = os.path.basename(file)
                webp_name = process_image(img_data, filename)
                if webp_name:
                    full_image_list.append(webp_name)

    print(f"Extracted {len(full_image_list)} images to {output_dir}")
    print("Files:", full_image_list)

    # 3. Attempt to map names (Very basic heuristic)
    # Read document.xml and look for text followed by blip/image
    # This is hard because XML structure varies. 
    # For now, let's just output the list and I will manually map them in index.html based on order/viewing.
