#!/bin/bash
cd /home/user/NP
nohup streamlit run npa/app.py --server.headless true --server.port 8501 > /home/user/NP/streamlit.log 2>&1 &
