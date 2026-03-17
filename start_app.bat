@echo off
cd /d C:\NP
echo [%date% %time%] Starting NPA app... >> C:\NP\startup.log
call venv\Scripts\activate
if errorlevel 1 (
    echo [%date% %time%] ERROR: venv activate failed >> C:\NP\startup.log
    exit /b 1
)
echo [%date% %time%] venv activated, launching streamlit... >> C:\NP\startup.log
streamlit run npa\app.py --server.headless true --server.port 8501 >> C:\NP\startup.log 2>&1
