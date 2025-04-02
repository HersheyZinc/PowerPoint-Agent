import os, re, pymupdf, sys
from pptx.util import Mm, Length
import json, shutil
from PIL import Image

OS_NAME = sys.platform
if OS_NAME == "win32":
    from pptxtopdf import convert as convert_pptx_to_pdf
    import pythoncom

def fromEmus(emus):
    # Convert emus to mm
    try: return round(Length(emus).mm, 2)
    except: return 0
    
def toEmus(length):
    # Convert mm to emus
    try: return Mm(length)
    except: return 0

def fromPts(pts):
    try: return Length(pts).pt
    except: return 0


def ppt_to_pdf(ppt_path, dst_dir):

    pdf_path = os.path.join(dst_dir, os.path.basename(ppt_path).replace(".pptx", ".pdf"))

    if OS_NAME == "win32":
        pythoncom.CoInitialize()
        convert_pptx_to_pdf(ppt_path, dst_dir) # Windows
        pythoncom.CoUninitialize()
    else:
        os.system(f'libreoffice --headless --convert-to pdf "{ppt_path}" --outdir {dst_dir}') # Linux

    return pdf_path


def pdf_to_img(pdf_path):
    images = []
    with pymupdf.open(pdf_path) as doc:
        for idx, page in enumerate(doc):
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)

    return images



def validate_hex(hex_string):
    if re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', hex_string):
        return True
    elif re.search(r'(?:[0-9a-fA-F]{3}){1,2}$', hex_string):
        return True
    else:
        return False


def empty_directory(dir):
    for filename in os.listdir(dir):
        file_path = os.path.join(dir, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')