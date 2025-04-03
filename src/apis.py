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
    
    def run(self, ppt, slide_idx, function_args, model):
        if self.function:
            response = self.function(ppt, slide_idx, function_args, model)
            return response
        
OUTLINES = [
    API(name="insert_slide", description="Inserts a new slide with the content.", required=['title', 'text'], function=insert_slide,
    parameters={
        "title": {"description": "Title of slide", "type": "string"},
        "text": {"description": "Exact text to be displayed on slide. Bullet points are indicated by newlines.", "type": "string"},
    }),
]


PLANS = [
    API(name="modify_slide", description="Modifies a specified slide according to the instructions given.", required=['slide_index', 'instructions'],
        parameters={
            "slide_index": {"description": "Index of the slide to update.", "type": "integer"},
            "instructions": {"description": "Detailed instructions for modifying a pre-existing slide.", "type": "string"},
        }),
    API(name="insert_slide", description="Inserts a new slide and modifies it according to the instructions given.", required=['slide_template'], function=insert_slide,
        parameters={
            "slide_template": {"description": "Description of the slide layout, e.g., 'Title with two side-by-side textboxes'.", "type": "string"},
            "instructions": {"description": "Detailed instructions for modifying a slide template. Skip this parameter to leave slide blank.", "type": "string"},
        }),
    # API(name="redo_slide", description="Empty the content of a specified slide and modifies it according to the instructions given.", required=['slide_index'], function=delete_all_shapes,
    #     parameters={
    #         "slide_index": {"description": "Index of the slide to clear.", "type": "integer"},
    #         "instructions": {"description": "Detailed instructions for modifying an empty slide. Skip this parameter to leave slide blank.", "type": "string"},
    #     }),
]


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