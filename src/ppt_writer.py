from pptx.dml.color import RGBColor
from .utils import validate_hex, toEmus
from pptx.util import Pt
from .openai import generate_image, query
from .prompts import select_shape_prompt


def set_shape_properties(shape, parameters):
    if "top" in parameters:
        shape.top = toEmus(parameters["top"])
    if "left" in parameters:
        shape.left = toEmus(parameters["left"])
    if "height" in parameters:
        shape.height = toEmus(parameters["height"])
    if "width" in parameters:
        shape.width = toEmus(parameters["width"])
    if "fill_color" in parameters:
        set_shape_fill(shape, parameters["fill_color"])
    if "has_border" in parameters:
        if parameters["has_border"]:
            shape.line.fill.solid()
        else:
            shape.line.fill.background()
    if "border_width" in parameters:
        shape.line.width = Pt(parameters["border_width"])
    if "border_color" in parameters and validate_hex(parameters["border_color"]):
        shape.line.color.rgb = RGBColor.from_string(parameters["border_color"].replace("#",""))
    if "text" in parameters:
        set_shape_text(shape, parameters)


def set_shape_text(shape, parameters):
    if "text" in parameters:
        shape.text = parameters["text"]
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if "fill_color" in parameters and validate_hex(parameters["font_color"]):
                run.font.color.rgb = RGBColor.from_string(parameters["font_color"].replace("#",""))
            if "font_size" in parameters:
                run.font.size = Pt(parameters["font_size"])
            if "bold" in parameters:
                run.font.bold = parameters["bold"]


def set_shape_fill(shape, fill_color):
    if fill_color =="transparent":
        shape.fill.background()
    elif validate_hex(fill_color):
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill_color.replace("#",""))


def set_table_properties(table, parameters):
    table_data = parameters["table_data"]
    data_index = 0
    for row in table.rows:
        for col in row.cells:
            if data_index < len(table_data):
                col.text = table_data[data_index]
                data_index += 1
            else:
                col.text = ""
    return




def modify_shape(slide, parameters):
    shape_idx = parameters["shape_index"]
    shape = slide.shapes[shape_idx]
    set_shape_properties(shape, parameters)


def modify_text(slide, parameters):
    shape_idx = parameters["shape_index"]
    shape = slide.shapes[shape_idx]
    set_shape_text(shape, parameters)


def modify_picture(slide, parameters):
    shape_idx = parameters["shape_index"]
    shape = slide.shapes[shape_idx]
    img = generate_image(parameters["description"])
    shape.image = img


def modify_background(slide, parameters):
    if "fill_color" in parameters:
        fill_color = parameters["fill_color"]

        if fill_color =="transparent":
            slide.background.fill.background()
        elif validate_hex(fill_color):
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor.from_string(fill_color.replace("#",""))


def modify_table(slide, parameters):
    shape_idx = parameters["shape_index"]
    table = slide.shapes[shape_idx]
    set_table_properties(table, parameters)

    return


def insert_shape(slide, parameters):
    messages = [{"role":"system", "content":select_shape_prompt}, {"role":"user", "content":parameters["shape_type"]}]
    r = query(messages, json_mode=True, max_tokens=10)
    shape_id = int(r["id"])
    shape = slide.shapes.add_shape(shape_id, 10, 10, 10, 10)
    shape.fill.solid()
    set_shape_properties(shape, parameters)
    return "Shape successfully inserted"


def insert_picture(slide, parameters):
    img = generate_image(parameters["description"])
    shape = slide.shapes.add_picture(img, 10, 10)
    set_shape_properties(shape, parameters)
    return


def insert_table(slide, parameters):
    rows, cols = parameters["rows"], parameters["columns"]
    table = slide.shapes.add_table(rows, cols, 10, 10, 10, 10).table
    set_shape_properties(table, parameters)
    set_table_properties(table, parameters)
    return


def delete_shapes(slide, parameters):
    shape_indexes = parameters["shape_indexes"]
    shapes = slide.shapes
    shapes_to_remove = [shapes[idx].element for idx in shape_indexes]
    remove_count = 0
    for shape in shapes_to_remove:
        try: 
            shapes.element.remove(shape)
            remove_count += 1
        except:
            continue
    return f"{remove_count} shapes successfully deleted"



def delete_all_shapes(slide):
    shape_indexes = [i for i, _ in enumerate(slide.shapes)]
    r = delete_shapes(slide, {"shape_indexes": shape_indexes})
    return r






