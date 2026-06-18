@echo off
REM Launch the Home Leads Funnel dashboard.
cd /d "%~dp0"
python -m streamlit run app.py
pause
