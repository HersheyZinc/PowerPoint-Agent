# 🧠 AgentPPT: LLM-Powered PowerPoint Editing Agent

**AgentPPT** is a cloud-deployed LLM agent that allows users to interact with and edit PowerPoint presentations using natural language. It combines GPT-4o’s reasoning capabilities with fine-grained control over presentation content, delivered through a chat-based web interface.

## 🚀 Overview

AgentPPT is built as a modular AI system that goes beyond simple prompt-to-slide generation. It supports **continuous, iterative editing** of presentations—users can request slide modifications, insert new content, or ask for design changes via natural language, and the agent responds by executing structured API calls.

The application is deployed on **Streamlit Community Cloud** and follows a **Software-as-a-Service (SaaS)** model.

---

## 🧱 Architecture

AgentPPT is composed of three core modules:

### 1. **Generate Module**
- Converts a user prompt into a structured slide outline.
- Applies rule-based logic to convert outlines into templated PowerPoint slides.
- Supports customization such as tone, slide count, and layout preferences.

### 2. **Plan Module**
- Parses user queries and presentation JSON.
- Decomposes instructions into slide-level tasks for targeted editing.
- Supports multi-slide and context-aware commands.

### 3. **Action Module**
- Executes slide-level tasks using function calling via GPT-4o.
- Interfaces with a suite of PowerPoint APIs to modify slide elements directly.

---

## 🧩 PowerPoint API

AgentPPT interfaces with PowerPoint through the [`python-pptx`](https://python-pptx.readthedocs.io/) library.

### Read APIs:
- Extract slide content and return structured JSON.
- Includes shape dimensions, positions, colors, and text.

### Write APIs:
- Insert/delete slides or shapes.
- Resize, reposition, and recolor objects.
- Modify text, images, tables, and charts.

These APIs enable **fine-grained control** beyond typical LLM assistants.

---

## 🤖 LLM Integration

AgentPPT uses the **OpenAI GPT-4o** model with the following features:

### 🔧 Function Calling
- Structured outputs map to predefined API calls (e.g. `resize_shape`, `change_background`).
- Ensures reliable and executable modifications.

### ✍️ Prompt Engineering
- Custom prompt templates guide the LLM’s behavior across all modules.
- Prompts include few-shot examples, role instructions, and constraints.

---

## 💻 User Interface

The frontend is built using **Streamlit**, featuring:

- **Slide View** – Displays rendered images of current slides.
- **Chat Interface** – Enables conversation with AgentPPT for querying and feedback.
- **Slide Navigation** – Allows users to browse and select slides.

---

## ☁️ Cloud Deployment

AgentPPT is deployed on **Streamlit Community Cloud (PaaS)**:
- Code hosted on GitHub
- Automatic environment setup via `requirements.txt`
- Secure key management via Streamlit Secrets
- Public URL access with no local setup required

---

## 📌 Example Use Cases

- “Change the title font on slide 3 to Arial Bold.”
- “Insert an image of a raccoon on slide 1.”
- “Add three more slides explaining our business strategy.”
- “Align the icons on slide 4 with those on slide 2.”

Here’s the updated **GitHub README** with a citation added for [PPTC](https://github.com/gydpku/PPTC), which you can include under the **References** section:

---

## 📚 References

- [PPTC: Programmatically Controllable LLM Agent for PowerPoint Tasks](https://github.com/gydpku/PPTC) – A research-driven system for programmatically controlling PowerPoint through LLM agents. Inspired our use of PowerPoint reader and writer modules.
