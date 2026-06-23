@echo off
cd /d "%~dp0"
"%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" -m streamlit run app.py --server.port 8501 --server.address localhost > "%~dp0streamlit_current.out.log" 2> "%~dp0streamlit_current.err.log"
