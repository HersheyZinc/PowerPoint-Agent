import streamlit as st
from src.agent import AgentPPT
import os

st.set_page_config(layout="wide")


if "agent" not in st.session_state:
    agent = AgentPPT()
    agent.new_ppt()
    st.session_state["slide_imgs"] = agent.render()
    st.session_state["agent"] = agent
    st.session_state["model"] = "gpt-4o"
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


    with tab_chat:
        chat = st.container(height=600)
        # 
        with chat:
            for message in st.session_state["agent"].chat_history:
                if message["role"] == "system":
                    continue
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])


        if prompt := st.chat_input("Type instructions to modify the current slide"):
            with chat:
                agent.slide_idx = st.session_state["slide_idx"]
                st.chat_message("user").markdown(prompt)

                response = agent.plan_module(prompt)
                
                st.chat_message("assistant").markdown(response)
                st.session_state["slide_imgs"] = agent.render()


    with tab_log:
        log = st.container(height=600)
        with log:
            for msg in agent.log:
                for m in msg.split("\n"):
                    st.write(m)


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
        slide_preview = st.image(st.session_state["slide_imgs"][slide_idx])
