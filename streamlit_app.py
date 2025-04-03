import streamlit as st
from src.agent import AgentPPT
import os, tempfile, shutil
from io import StringIO
st.set_page_config(layout="wide")


if "agent" not in st.session_state:
    agent = AgentPPT()
    agent.new_ppt()
    st.session_state["slide_imgs"] = agent.render()
    st.session_state["agent"] = agent
    st.session_state["slide_idx"] = 0


col_left, col_right = st.columns([0.6, 0.4])
agent = st.session_state["agent"]
slide_idx = st.session_state["slide_idx"]


with col_right:
    tab_chat, tab_config, tab_log = st.tabs(["Chat", "Settings", "Logs"])

    with tab_config:
        config = st.container(height=600)
        with config:
            if st.button("Reset", type="primary"):
                agent.clear_chat_history()
                agent.new_ppt()
                st.session_state["slide_imgs"] = agent.render()
                slide_idx = 0

            if st.button("Export Presentation"):
                if len(agent.ppt.slides) > 0:
                    agent.save_ppt()
                    with open(agent.ppt_path, "rb") as f:
                        st.download_button(label="Download", data=f, file_name="presentation.pptx")

    with tab_chat:
        chat = st.container(height=600)
        # 
        with chat:
            with st.chat_message("assistant"):
                st.markdown("Welcome to AgentPPT! Enter an idea for a PowerPoint presentation to get started!")
            for message in st.session_state["agent"].chat_history:
                if message["role"] == "system":
                    continue
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])


        if prompt := st.chat_input("Type instructions to modify the current slide"):
            with chat:
                agent.slide_idx = st.session_state["slide_idx"]
                st.chat_message("user").markdown(prompt)

                if len(agent.chat_history) == 0:
                    response = agent.generate_module(prompt)
                else:
                    response = agent.plan_module(prompt)
                
                st.chat_message("assistant").markdown(response)
                st.session_state["slide_imgs"] = agent.render()


    with tab_log:
        log = st.container(height=600)
        with log:
            for msg in agent.logger:
                st.write(msg)


with col_left:
    
    st.title("PowerPoint Agent")
    slide_preview_container = st.container()
    slide_selection_container = st.container()
    
    
    with slide_selection_container:
        col1, col2, col3, col4, col5, = st.columns([4,1.3,2,1.3,4])
        
        with col2:
            if st.button("prev", use_container_width=True):
                slide_idx = max(0, slide_idx-1)
        with col4:
            if st.button("next", use_container_width=True):
                slide_idx = min(len(agent.ppt.slides)-1, slide_idx+1)
        with col3:
            st.button(f"Slide {slide_idx+1} of {len(agent.ppt.slides)}", use_container_width=True, disabled=True)
    

    with slide_preview_container:
        st.session_state["slide_idx"] = slide_idx
        with st.container(border=True):
            slide_preview = st.image(st.session_state["slide_imgs"][slide_idx], use_column_width=True)
