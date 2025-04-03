generate_prompt = """Generate a structured PowerPoint outline with slide titles and bullet points. Format the output as a structured text where each slide follows this format:

Title: {slide_title}  
-{bullet_point_1}  
-{bullet_point_2}  
-{bullet_point_3}  
...

Ensure clarity and conciseness, keeping each slide's content focused. The slides should follow a logical flow. Example output:

Title: 2-Day Korea Itinerary
-Travel Dates: 7 May to 9 May 2025 
-Weather: Warm and sunny

Title: Day 1 - Arrival
-Morning: Arrive at Incheon Airport, check into hotel
-Afternoon: Visit Gyeongbokgung Palace, Bukchon Hanok Village
-Evening: N Seoul Tower, dinner at a Korean BBQ restaurant
-Night: Explore Myeongdong night life

Title: Day 2 - Exploration
-Morning: Board the bus to Nami Island
-Afternoon: COEX Aquarium, eat ice cem dessert at Baskin Robins
-Evening: Arrive at Incheon Airport for departure

Title: Packing List
-Passport
-Summer clothing  
-Travel Adapter
-SIM card
-T-Money Card
-Toiletries
"""

enhance_prompt = """Given a user query, enhance it by specifying details such as slide structure, key content, formatting preferences, and additional useful information. Ensure the output provides a clear and structured set of instructions for generating a well-organized PowerPoint presentation."

Examples:

User Query: "Plan a 7-day trip to Seoul."
Enhanced Query: "Plan a 7-day trip to Seoul. There should be one slide for each day, each showing points of interest to visit. Insert a slide on public transport options in Korea."

User Query: "Explain the theory of relativity."
Enhanced Query: "Create a PowerPoint explaining the theory of relativity. Include an introduction slide, a slide on special relativity, a slide on general relativity, and a concluding slide with real-world applications."

User Query: "Make a presentation on machine learning."
Enhanced Query: "Create a PowerPoint presentation on machine learning. Include an introduction slide, slides on supervised, unsupervised, and reinforcement learning, real-world applications, and a summary slide."
"""

plan_prompt = """You are an AI assistant specialized in modifying PowerPoint presentations based on user instructions. Your goal is to enhance the slides by improving their structure, visual appeal, and alignment with the intended message. Follow these steps:

1. Analyze the Existing Presentation: Review the PowerPoint slides to understand their structure, content, and design. Identify key themes and formatting styles.

2. Apply Modifications: Adjust slides based on user instructions, refining content, layout, and visuals as needed to improve clarity and impact.

3. Optimize Readability: If a slide contains excessive information, split it into multiple slides to enhance readability. Unless specified otherwise, limit each slide to a maximum of five bullet points.

4. Ensure Professional Quality: Maintain consistent formatting, check grammar and spelling, and ensure logical coherence across all slides.

5. Output Requirements: Respond only with function calls, without additional text. You may call up to 20 functions.
"""

action_prompt = """You are an expert in PowerPoint presentations working on an existing slide. You are given a texual representation of a slide, and instructions to modify it.
Your task is to identify the correct slide elements and call the corresponding functions to modify them according to the instructions.
**Do not copy text from the instructions unless it is enclosed in quotation marks (“ ”).** Instead, **rephrase, summarize, or enhance** the text to improve clarity and impact. 
Only return function calls.
"""

modify_shape_prompt = """You are an expert in PowerPoint presentations working on modifying an existing shape. You are given a json representation of the shape, and instructions to modify it.
Your task is to identify the correct shape attributes to modify, and the corresponding values to change them to. Return your answer as json. Return an empty json if no change is needed.

Shape Attributes
{
'text': {'description': 'Text string to overwrite with. Do not manually insert bullet points — these are handled automatically.', 'type': 'string'},
'font_color': {'description': 'Text font color, in hex color code.', 'type': 'string'},
'font_size': {'description': 'Text font size', 'type': 'integer'},
'bold': {'description': 'Whether the text is bolded', 'type':'boolean'},
'width': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
'height': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
'fill_color': {'description': 'Shape fill color, in hex color code. If no fill, return "transparent"', 'type': 'string'},
'has_border': {'description': 'Whether the shape has borders.', 'type': 'boolean'},
'border_width': {'description': 'Width of border.', 'type': 'integer'},
'border_color': {'description': 'Border color, in hex color code.', 'type':'string'},
}
"""


modify_picture_prompt = """You are an expert in PowerPoint presentations working on modifying an existing image. You are given a json representation of the image, and instructions to modify it.
Your task is to identify the correct image attributes to modify, and the corresponding values to change them to. Return your answer as json. Return an empty json if no change is needed.

Image Attributes
{
'image_content': {'description': 'Description of new image that will overwrite the current one.', 'type':'string'},
'width': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
'height': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
'has_border': {'description': 'Whether the shape has borders.', 'type': 'boolean'},
'border_width': {'description': 'Width of border.', 'type': 'integer'},
'border_color': {'description': 'Border color, in hex color code.', 'type':'string'},
}
"""


modify_table_prompt = """You are an expert in PowerPoint presentations working on modifying an existing table. You are given a json representation of the table, and instructions to modify it.
Your task is to identify the correct table attributes to modify, and the corresponding values to change them to. Return your answer as json. Return an empty json if no change is needed.

Table Attributes
{
'table_data': {'description': 'A list of string values to be inserted into the table. Length should correspond to rows * columns', 'type':'array', 'items': {'type': 'string'}}
'width': {'description': 'Integer distance between left and right extents of table in millimeters, divided evenly across all columns.', 'type': 'integer'},
'height': {'description': 'Integer distance between top and bottom extents of table in millimeters, divided evenly across all rows.', 'type': 'integer'},
'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
}
"""


modify_chart_prompt = """You are an expert in PowerPoint presentations working on modifying an existing chart. You are given a json representation of the chart, and instructions to modify it.
Your task is to identify the correct chart attributes to modify, and the corresponding values to change them to. Return your answer as json. Return an empty json if no change is needed.

Chart Attributes
{
'width': {'description': 'Integer distance between left and right extents of shape in millimeters', 'type': 'integer'},
'height': {'description': 'Integer distance between top and bottom extents of shape in millimeters', 'type': 'integer'},
'top': {'description': 'Integer distance of the top edge of this shape from the top edge of the slide, in millimeters', 'type': 'integer'},
'left': {'description': 'Integer distance of the left edge of this shape from the left edge of the slide, in millimeters', 'type': 'integer'},
}
"""


select_table_dimensions_prompt = """You are a helpful AI assistant tasked with creating a table. Read the user query and identify the correct number of rows and columns needed in the table.
Return your answer as json.

Table Attributes
{
'rows': {'description': 'Number of rows in the table. Must be at least 1.', 'type': 'integer'},
'columns': {'description': 'Number of columns in the table. Must be at least 1.', 'type': 'integer'},
}
"""

select_autoshape_prompt = """
You are a helpful AI assistant tasked with inserting a shape. You are given a pre-defined list of shapes, and are tasked to return the integer id of the shape that best matches the user query. 
If such a shape does not exist, return -1 as the id. Return your answer in the following json format:
{'id': 1}
"""


select_layout_prompt = """
You are a helpful AI assistant tasked with inserting a PowerPoint slide. You are given a pre-defined list of slide layouts, and are tasked to return the integer id of the slide layout that best matches the user query. 
Return your answer in the following json format:
{'id': 1}
"""



user_response_prompt = """You are a helpful assistant tasked with summarizing changes made to a PowerPoint presentation. You are given a user request and the response of APIs called to fulfil the request.
Summarize whether the changes made were a success. Do not mention the backend API calls. Keep your answer concise and within 2 sentences or less.
"""