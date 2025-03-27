from pptx import Presentation
from .utils import render_slides, fromEmus, empty_directory
from .ppt_reader import get_slide_content, get_ppt_outline
from .ppt_writer import delete_all_shapes
from .openai import query, query_tools
from . import apis, prompts
import json, os, base64

# TODO
# Ingest existing presentations -> Iterate through each slide and generate a summary


class AgentPPT():
    def __init__(self, model="gpt-4o-mini", dst_path="test.pptx"):
        self.debug = True
        self.ppt = None
        self.ppt_path = dst_path
        self.slide_preview_dir = "slide_previews"
        os.makedirs(self.slide_preview_dir, exist_ok=True)

        # LLM args
        self.model = model
        self.model_temp = 0

        self.chat_history = []
        self.log = []

        self.new_ppt()
        self.save_ppt()
        self.clear_chat_history()


    def new_ppt(self, file_path=""):
        empty_directory(self.slide_preview_dir)
        if file_path:
            self.ppt = Presentation(file_path)
        else:
            self.ppt = Presentation()
        # self.clear_chat_history()
        # self.insert_slide()
        self.log.append("New presentation created")
        

    def save_ppt(self):
        # Write content to pptx file
        self.ppt.save(self.ppt_path)
        self.log.append(f"Presentation saved to {self.ppt_path}")


    def clear_chat_history(self):
        system_prompt = "You are an expert Powerpoint slide designer. Use the tools provided to fulfil the user's request."
        self.chat_history = [{"role":"system", "content":system_prompt}]
        self.log = []


    def print_chat_history(self):
        for msg in self.chat_history:
            print(f"{msg["role"]}:\n{msg["content"]}\n{'--------------'*10}")


    def print_log(self):
        for msg in self.log:
            print(msg)
            print('--------------'*10)


    def print_ppt(self):
        for i, slide in enumerate(self.ppt.slides):
            print(get_slide_content(self.ppt, i))
    

    def render_slide(self, slide_index):
        render_slides(self.ppt_path, self.slide_preview_dir, [slide_index])


    def render_all_slides(self):
        render_slides(self.ppt_path, self.slide_preview_dir, slide_indexes=list(range(len(self.ppt.slides))))
        
    
    def plan_module(self, user_prompt):
        self.chat_history.append({"role":"user", "content": user_prompt})
        ppt_outline = get_ppt_outline(self.ppt)

        messages = [{"role":"system", "content":prompts.plan_prompt}] + self.chat_history[-5:] + [{"role":"system", "content":ppt_outline}]
        toolkit = [a.get_openai_args() for a in apis.PLANS]
        _, tool_calls = query_tools(messages, toolkit)

        output_str = ""
        for i, tool_call in enumerate(tool_calls):
            
            fn_name = tool_call.function.name
            fn_args = json.loads(tool_call.function.arguments)
            print(fn_name, fn_args)

            if fn_name == "insert_slide":
                slide_template = fn_args["slide_template"]
                instructions = fn_args["instructions"]
                layout_index = self.select_slide_layout(slide_template)
                r = self.insert_slide(layout_index)
                output_str += r + "\n"

                slide_index = len(self.ppt.slides) - 1
                r = self.action_module(slide_index, instructions, apis.ACTIONS_BASIC)
                output_str += r + "\n"

            elif fn_name == "clear_slide":
                slide_index = fn_args["slide_index"]
                slide = self.ppt.slides[slide_index]
                delete_all_shapes(slide)
                output_str += f"Slide {slide_index+1} cleared.\n"

            elif fn_name == "modify_slide_basic":
                slide_index = fn_args["slide_index"]
                instructions = fn_args["instructions"]
                r = self.action_module(slide_index, instructions, apis.ACTIONS_BASIC)
                output_str += r + "\n"

            elif fn_name == "modify_slide_advanced":
                slide_index = fn_args["slide_index"]
                instructions = fn_args["instructions"]
                r = self.action_module(slide_index, instructions, apis.ACTIONS_ADVANCED)
                output_str += r + "\n"


        self.chat_history.append({"role":"assistant", "content":output_str})
        print(output_str)

    

    def action_module(self, slide_index, instructions, apis_selected=apis.ACTIONS_BASIC):
        # Selects and calls the PowerPoint APIs corresponding to the instructions given
        slide = self.ppt.slides[slide_index]
        slide_content = get_slide_content(self.ppt, slide_index)
        toolkit = [a.get_openai_args() for a in apis_selected]

        slide_height, slide_width = fromEmus(self.ppt.slide_width), fromEmus(self.ppt.slide_width)
        slide_content += f"The actual dimensions of the slide is {slide_width}mm width and {slide_height}mm height."

        prompt = [{"role":"system", "content": prompts.action_prompt + slide_content}, {"role": "user", "content": instructions}]

        _, tool_calls = query_tools(prompt, toolkit)
        print(tool_calls)

        output_str = f"**Modified slide {slide_index+1}**\n"
        for tool_call in tool_calls:
            fn_name = tool_call.function.name
            fn_args = json.loads(tool_call.function.arguments)
            for api in apis_selected:
                if api.name == fn_name:
                    try:
                        r = api.run(slide, fn_args)
                        output_str += f"API - {fn_name} | parameters - {fn_args} | Status - SUCCESS\n"
                    except Exception as e:
                        output_str += f"API - {fn_name} | parameters - {fn_args} | Status - FAILED | {e}\n"
        
        return output_str


    def insert_slide(self, layout_idx=1, name="", summary="", script=""):
        slide_layout = self.ppt.slide_layouts[layout_idx]
        slide = self.ppt.slides.add_slide(slide_layout)

        slide.name = name
        slide.description = f"An empty slide with layout - {slide_layout.name}"
        slide.summary = summary
        slide.script = script

        return "Slide inserted."


    
    def select_slide_layout(self, slide_description):
        layout_s = "Slide templates:\n" + "\n".join([f"{i} - {layout.name}" for i, layout in enumerate(self.ppt.slide_layouts)]) # Get all layouts given by slide master
        
        prompt = [{"role": "system", "content":prompts.select_layout_prompt + layout_s}, {"role":"user", "content":slide_description}]

        r = query(prompt, json_mode=True, temperature=0.1, max_tokens=10)

        slide_layout_idx = int(r["id"]) # prompt engineering hardcodes json output with 'id' key
        return slide_layout_idx
        
