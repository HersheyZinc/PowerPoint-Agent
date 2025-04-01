from .ppt_writer import *


class API():
    def __init__(self, name:str, description:str, parameters:dict, required:list, function=None):
        self.name = name
        self.description = description
        self.parameters = parameters
        self.required = required
        self.function = function

    def get_openai_args(self):
        args = {
            'type': 'function',
            'function': {
                'name': self.name,
                'description': self.description,
                'parameters': {
                    'type': 'object',
                    'properties': self.parameters
                },
                'required': self.required
            }
        }
        return args
    
    def run(self, selected_slide, function_args):
        response = self.function(selected_slide, function_args)
        return response

# PLANS = [
#     API(name="clear_slide", description="Delete all objects from a specified slide.", required=['slide_index'],
#         parameters={
#             "slide_index": {"description": "Index of the slide to clear.", "type": "integer"},
#         }),
#     API(name="modify_slide_basic", description="Update text and image content on a slide. Use for simple edits.", required=['slide_index', 'instructions'],
#         parameters={
#             "slide_index": {"description": "Index of the slide to update.", "type": "integer"},
#             "instructions": {"description": "Detailed instructions for modifying text or images.", "type": "string"},
#         }),
#     API(name="modify_slide_advanced", description="Change any attributes of objects on a slide, such as size, position, text, or color.", required=['slide_index', 'instructions'],
#         parameters={
#             "slide_index": {"description": "Index of the slide to modify.", "type": "integer"},
#             "instructions": {"description": "Detailed instructions for modifying objects.", "type": "string"},
#         }),
#     API(name="insert_slide", description="Add a new empty slide to the end of the presentation.", required=['slide_template', 'instructions'],
#         parameters={
#             "slide_template": {"description": "Description of the slide layout, e.g., 'Title with two side-by-side textboxes'.", "type": "string"},
#             "instructions": {"description": "Detailed instructions for text content.", "type": "string"},
#         }),
# ]

# DESIGNS = [

#     API(name="modify_slide_basic", description="Modifies only the text and image content of existing shapes. Use by default.", required=['instructions'],
#         parameters={
#             "instructions": {"description": "A detailed list of instructions describing what needs to be modified on the slide. Texts should be in point form.", "type": "string"},
#         }),

#     API(name="modify_slide_advanced", description="Function to insert, delete, and manipulate slide elements such as size, color and position. Use when the instructions extend beyond text/image modification.", required=['instructions'],
#         parameters={
#             "instructions": {"description": "A detailed list of instructions describing what needs to be modified on the slide", "type": "string"},
#         }),
        
# ]


# ACTIONS_BASIC = [
#     API(name="modify_shape_text", description="Modifies the text of an existing shape", required=["text","shape_index"], function=modify_text,
#         parameters={
#             'text': {'description': 'Text content to be overwritten to the shape', 'type': 'string'},
#             'shape_index': {'description': 'Index of shape to be modified', 'type':'integer'}
#         }),

#     API(name="modify_picture", description="Modifies the image of an existing picture", required=["description", "shape_index"], function=modify_picture,
#         parameters={
#             'description': {'description': 'Text description of the content of the image to be generated', 'type':'string'},
#             'shape_index': {'description': 'Index of picture to be modified', 'type':'integer'}
#         })
# ]


# ACTIONS_ADVANCED = [
#     API(name="insert_shape", description="Inserts a shape into the slide", required=['shape_type', 'height', 'width', 'top', 'left'], function=insert_shape,
#         parameters={
#             'shape_type': {'description': 'Type of shape to be inserted. Example: Rectangle, Oval', 'type': 'string'},
#             'height': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
#             'width': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
#             'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
#             'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
#             'fill_color': {'description': 'Shape fill color, in hex color code. If no fill, return "transparent"', 'type': 'string'},
#             'has_border': {'description': 'Whether the shape has borders (1) or no borders (0).', 'type': 'integer'},
#             'border_width': {'description': 'Width of border.', 'type': 'integer'},
#             'border_color': {'description': 'Border color, in hex color code.', 'type':'string'},
#             'text': {'description': 'Text to be displayed in the shape.', 'type': 'string'},
#             'font_color': {'description': 'Text font color, in hex color code.', 'type': 'string'},
#             'font_size': {'description': 'Text font size', 'type': 'integer'},
#             }),

#     API(name="insert_table", description="Insert a table into the slide", required=["rows", "columns", "table_data", "top", "left", "height", "width"], function=insert_table,
#         parameters={
#             'rows': {'description': 'Number of rows in the table', 'type': 'integer'},
#             'columns': {'description': 'Number of columns in the table', 'type': 'integer'},
#             'table_data': {'description': 'A list of string values to be inserted into the table. Length should correspond to rows * columns', 'type':'array', 'items': {'type': 'string'}},
#             'height': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
#             'width': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
#             'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
#             'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
#         }),

#     API(name="remove_shape", description="Removes one or more shapes by index", required=["shape_indexes"], function=delete_shapes,
#         parameters={
#             'shape_indexes': {"description": 'An array of integer indexes corresponding to the shapes to be deleted.', 'type': 'array', 'items':{'type': 'integer'}}
#         }),

#     API(name="modify_shape_properties", description="Modifies an existing shape", required=["shape_index"], function=modify_shape,
#         parameters={
#             'shape_index': {'description': 'Index of shape to be modified', 'type':'integer'},
#             'height': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
#             'width': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
#             'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
#             'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
#             'fill_color': {'description': 'Shape fill color, in hex color code. If no fill, return "transparent"', 'type': 'string'},
#             'has_border': {'description': 'Whether the shape has borders (1) or no borders (0).', 'type': 'integer'},
#             'border_width': {'description': 'Width of border.', 'type': 'integer'},
#             'border_color': {'description': 'Border color, in hex color code.', 'type':'string'},
#             }),

#     API(name="modify_shape_text", description="Modifies the text of an existing shape", required=["shape_index"], function=modify_text,
#         parameters={
#             'text': {'description': 'New text', 'type': 'string'},
#             'shape_index': {'description': 'Index of shape to be modified', 'type':'integer'},
#             'font_color': {'description': 'Text font color, in hex color code.', 'type': 'string'},
#             'font_size': {'description': 'Text font size', 'type': 'integer'},
#             'bold': {'description': 'Whether the text is bolded', 'type':'boolean'},
#         }),

#     API(name="insert_picture", description="Insert a picture into the slide", required=["description", "top", "left", "height", "width"], function=insert_picture,
#         parameters={
#             'description': {'description': 'Text description of the image to be generated', 'type':'string'},
#             'height': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
#             'width': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
#             'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
#             'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
#         }),

#     API(name="modify_picture", description="Modifies the image of an existing picture", required=["description", "shape_index"], function=modify_picture,
#         parameters={
#             'description': {'description': 'Text description of the content of the image to be generated', 'type':'string'},
#             'shape_index': {'description': 'Index of picture to be modified', 'type':'integer'}
#         }),


    
#     API(name="modify_table", description="Modifies the cell values of an existing table", required=["table_data", "shape_index"], function=modify_table,
#         parameters={
#             'table_data': {'description': 'A list of string values to be inserted into the table', 'type':'array', 'items': {'type': 'string'}},
#             'shape_index': {'description': 'Index of table to be modified', 'type':'integer'}
#         }),

#     API(name="modify_background", description="Modify the background color of the slide", required=["fill_color"], function=modify_background,
#         parameters={
#             'fill_color': {'description': 'Shape fill color, in hex color code. If no fill, return "transparent"', 'type': 'string'},
#         })
    
# ]


ACTIONS = [
    API(name="modify_shape", description="Modifies visual appearance of slide shape, such as text, size, position, color.", required=["instructions", "shape_index"], function=modify_shape,
        parameters={
            'instructions': {'description': 'Detailed instructions on how the shape should be modified.', 'type':'string'},
            'shape_index': {'description': 'Index of shape to be modified.', 'type':'integer'}
        }),

    API(name="modify_background", description="Modifies the background color of the slide.", required=["fill_color"], function=modify_background,
        parameters={
            'fill_color': {'description': 'Background fill color, in hex color code. If no fill, return "transparent".', 'type': 'string'},
        }),

    API(name="remove_shapes", description="Removes one or more shapes by index.", required=["shape_indexes"], function=delete_shapes,
        parameters={
            'shape_indexes': {"description": 'An array of integer indexes corresponding to the shapes to be deleted.', 'type': 'array', 'items':{'type': 'integer'}}
        }),

    API(name="insert_shape", description="Inserts a shape into the slide.", required=["shape_type", "instructions"], function=insert_shape,
        parameters={
            'shape_type': {"description": "The type of shape to be inserted. Must be from ['PICTURE', 'CHART', 'TABLE', 'TEXT_BOX', 'AUTO_SHAPE'].", 'type': 'string'},
            'instructions': {'description': 'Detailed instructions on how the shape should look. Include important information such as size and position.', 'type':'string'}
        }),
    

]