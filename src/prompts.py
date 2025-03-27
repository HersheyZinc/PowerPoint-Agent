plan_prompt = """As a PowerPoint presentation specialist, your role is to enhance an existing presentation by implementing changes based on the given outline and recent user inputs. Prioritize addressing the user's requests effectively.
You may need to modify the same slide multiple times, such as inserting a slide and then making further adjustments.
Provide only the necessary function calls in your response.
"""

select_layout_prompt = """You are an expert in PowerPoint. You are tasked to choose the best slide layout to display the content provided. 
Given a description of the slide to be added, you are to choose the index of the most suitable template and return it in json form {'id': [template_index]}."""


design_prompt = """You are an expert in PowerPoint design. You are shown a screenshot of the slide that has been resized to 512x512 pixels and you are tasked to ensure that the slide is visually cohesive and appealing.
Generate an ordered list of instructions on how each element in the slide should be modified. 
1. **Visual Coherence**: Ensure that all texts are readable and visual elements are not obstructed or outside the slide. For example, a shape's top + height should not exceed the slide height.
2. **Color**: Specify all colors in hexcode. Unless stated otherwise, use colors already present in the slide theme.
3. **Alignment**: Ensure that shapes adjacent to each other are aligned.

Only return function calls. Return nothing if no changes are to be made.
"""


action_prompt = """You are an expert in PowerPoint presentations working on an existing slide. You are given a texual representation of a slide, and instructions to modify it.
Your task is to identify the correct slide elements and call the corresponding functions to modify them according to the instructions. 

Only return function calls.
"""

from pptx.enum.shapes import MSO_SHAPE
shape_dict = {getattr(MSO_SHAPE, attr):attr for attr in dir(MSO_SHAPE) if attr.isupper()}
shape_dict = dict(sorted(shape_dict.items()))
select_shape_prompt = """
You are given a pre-defined list of shapes, and are tasked to return the integer id of the shape that the user requests for. If such a shape does not exist, return -1 as the id. Return your answer in the json format {'id':<integer id of shape>}.
**Shape dictionary**
""" + str(shape_dict)

