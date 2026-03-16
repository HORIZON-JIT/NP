@echo off
cd /d C:\NP
call venv\Scripts\activate
streamlit run npa\app.py --server.headless true
