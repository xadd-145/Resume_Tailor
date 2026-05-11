import streamlit as st
import os
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(page_title="Resume Tailor", layout="wide")
st.title("Resume Tailoring Engine")
st.write("Pipeline is ready. Phases coming next.")