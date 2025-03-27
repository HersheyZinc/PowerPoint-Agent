import  os, re, pymupdf
from pptx.util import Mm, Length
import json, shutil
from pptxtopdf import convert as convert_pptx_to_pdf

def fromEmus(emus):
    try: return round(Length(emus).mm, 2)
    except: return 0
    
def toEmus(length):
    try: return Mm(length)
    except: return 0

def fromPts(pts):
    try: return Length(pts).pt
    except: return 0


def render_slides(ppt_path="test.pptx", dst_dir="slide_previews"):
    os.makedirs(dst_dir, exist_ok=True)
    empty_directory(dst_dir)
    pdf_path = ppt_path.replace(".pptx", ".pdf")
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    convert_pptx_to_pdf(ppt_path, "")

    with pymupdf.open(pdf_path) as doc:
        for idx, page in enumerate(doc):
            pix = page.get_pixmap()
            pix.save(f"{dst_dir}/{idx}.png")



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