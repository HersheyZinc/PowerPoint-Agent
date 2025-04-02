from pptx import Presentation
from .utils import fromEmus, ppt_to_pdf, pdf_to_img
from .ppt_reader import get_slide_content, get_ppt_content
from .ppt_writer import insert_slide, delete_all_shapes
from .openai import query, query_tools
from . import apis, prompts
import json, os, tempfile, shutil


# TODO
# Ingest existing presentations -> Iterate through each slide and generate a summary


class AgentPPT():
    def __init__(self, model="gpt-4o", src_path="", dst_path="working.pptx"):
        self.ppt = None
        self.ppt_path = dst_path
        self.slide_idx = 0

        # LLM args
        self.model = model
        self.model_temp = 0

        self.chat_history = []
        self.log = []

        self.new_ppt(src_path)


    def new_ppt(self, file_path=""):
        if file_path:
            self.ppt = Presentation(file_path)
        else:
            self.ppt = Presentation()
            self.insert_slide()

        self.slide_idx = 0
        self.clear_chat_history()
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
    
        
    def render(self):
        temp_dir = tempfile.mkdtemp()

        temp_ppt_path = os.path.join(temp_dir, "temp_presentation.pptx")
        self.ppt.save(temp_ppt_path)
        pdf_path = ppt_to_pdf(temp_ppt_path, temp_dir)
        slide_images = pdf_to_img(pdf_path)

        shutil.rmtree(temp_dir)

        return slide_images


    def plan_module(self, prompt):
        self.chat_history.append({"role":"user", "content": prompt})
        ppt_content = get_ppt_content(self.ppt)
        messages = [{"role":"system", "content":prompts.plan_prompt}, {"role":"system", "content":ppt_content}] + self.chat_history[-5:]
        toolkit = [a.get_openai_args() for a in apis.PLANS]
        _, tool_calls = query_tools(messages, toolkit, model=self.model)
        
        for tool_call in tool_calls:
            fn_name = tool_call.function.name
            fn_args = json.loads(tool_call.function.arguments)

            output_str = ""
            if fn_name == 'insert_slide':
                template_request = fn_args["slide_template"]
                output_str += insert_slide(self.ppt, template_request, self.model) + "\n"
                slide_idx = len(self.ppt.slides)-1
                

            elif fn_name == "redo_slide":
                slide_idx = fn_args['slide_index']
                output_str += delete_all_shapes(self.ppt, slide_idx) + "\n"

            elif fn_name == "modify_slide":
                slide_idx = fn_args['slide_index']

            else:
                continue

            if "instructions" in fn_args:
                output_str += f"Modified slide {slide_idx} with instructions: {fn_args["instructions"]}\n"
                output_str += self.action_module(fn_args["instructions"], slide_idx)


            yield output_str


    def action_module(self, prompt, slide_idx):
        slide_content = get_slide_content(self.ppt, slide_idx)
        slide = self.ppt.slides[slide_idx]
        messages = [{"role":"system", "content":prompts.action_prompt}, {"role":"system", "content":slide_content}, {"role":"user", "content":prompt}]
        toolkit = [a.get_openai_args() for a in apis.ACTIONS]
        _, tool_calls = query_tools(messages, toolkit, model=self.model)

        output_str = ""
        for tool_call in tool_calls:
            fn_name = tool_call.function.name
            fn_args = json.loads(tool_call.function.arguments)
            for api in apis.ACTIONS:
                if api.name == fn_name:
                    try:
                        r = api.run(self.ppt, slide_idx, fn_args, self.model)
                        output_str += f"API - {fn_name} | Status - SUCCESS | {r} | Arguments - {fn_args}\n"
                    except Exception as e:
                        output_str += f"API - {fn_name} | Status - FAILED | parameters - {fn_args} | {e}\n"

        return output_str
    



    
    # def plan_module(self, user_prompt):
    #     self.chat_history.append({"role":"user", "content": user_prompt})
    #     ppt_outline = get_ppt_content(self.ppt)

    #     messages = [{"role":"system", "content":prompts.plan_prompt}] + self.chat_history[-5:] + [{"role":"system", "content":ppt_outline}]
    #     toolkit = [a.get_openai_args() for a in apis.PLANS]
    #     _, tool_calls = query_tools(messages, toolkit, model=self.model)

    #     output_str = ""
    #     for i, tool_call in enumerate(tool_calls):
            
    #         fn_name = tool_call.function.name
    #         fn_args = json.loads(tool_call.function.arguments)
    #         print(fn_name, fn_args)

    #         if fn_name == "insert_slide":
    #             slide_template = fn_args["slide_template"]
    #             instructions = fn_args["instructions"]
    #             layout_index = self.select_slide_layout(slide_template)
    #             r = self.insert_slide(layout_index)
    #             output_str += r + "\n"

    #             slide_index = len(self.ppt.slides) - 1
    #             r = self.action_module(slide_index, instructions, apis.ACTIONS_BASIC)
    #             output_str += r + "\n"

    #         elif fn_name == "clear_slide":
    #             slide_index = fn_args["slide_index"]
    #             slide = self.ppt.slides[slide_index]
    #             delete_all_shapes(slide)
    #             output_str += f"Slide {slide_index+1} cleared.\n"

    #         elif fn_name == "modify_slide_basic":
    #             slide_index = fn_args["slide_index"]
    #             instructions = fn_args["instructions"]
    #             r = self.action_module(slide_index, instructions, apis.ACTIONS_BASIC)
    #             output_str += r + "\n"

    #         elif fn_name == "modify_slide_advanced":
    #             slide_index = fn_args["slide_index"]
    #             instructions = fn_args["instructions"]
    #             r = self.action_module(slide_index, instructions, apis.ACTIONS_ADVANCED)
    #             output_str += r + "\n"


    #     self.log.append(output_str)
    #     system_prompt = f"{user_response_prompt}\n\nUser Request:\n{user_prompt}\n\nAPI reponses:{output_str}"
    #     api_summary = query([{"role":"system", "content":system_prompt}], model=self.model)
    #     self.chat_history.append({"role":"assistant", "content":api_summary})

    #     return api_summary


    def insert_slide(self, layout_idx=1, name="", summary="", script=""):
        slide_layout = self.ppt.slide_layouts[layout_idx]
        slide = self.ppt.slides.add_slide(slide_layout)

        slide.name = name
        slide.description = f"An empty slide with layout - {slide_layout.name}"
        slide.summary = summary
        slide.script = script

        return "Slide inserted."
        
