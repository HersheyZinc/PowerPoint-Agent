import win32com.client, os, pythoncom, re
from pptx.util import Mm, Length
import json, shutil

def fromEmus(emus):
    try: return round(Length(emus).mm, 2)
    except: return 0
    
def toEmus(length):
    try: return Mm(length)
    except: return 0

def fromPts(pts):
    try: return Length(pts).pt
    except: return 0


def render_slides(ppt_path="test.pptx", dst_dir="slide_previews", slide_indexes=[0]):
    os.makedirs(dst_dir, exist_ok=True)
    pythoncom.CoInitialize()
    Application = win32com.client.Dispatch("PowerPoint.Application")
    full_path = os.getcwd()
    Presentation = Application.Presentations.Open(os.path.join(full_path, ppt_path), WithWindow=False)
    for slide_idx in slide_indexes:
        Presentation.Slides[slide_idx].Export(os.path.join(full_path, f"{dst_dir}/{slide_idx}.jpg"), "JPG")
    Application.Quit()


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