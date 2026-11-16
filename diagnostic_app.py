"""Minimal Streamlit test to diagnose Community Cloud startup independently of the main app."""
import sys
import streamlit as st

st.set_page_config(page_title="Matching Mapper - Cloud diagnostics")
st.title("Streamlit Cloud health check")
st.success("The frontend, Python server, and Streamlit script are responding.")
st.write("Python version:", sys.version.split()[0])
st.write("Streamlit version:", st.__version__)
st.caption("This page deliberately does not load the matching or Excel processing modules.")
