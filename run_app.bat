@echo off
echo Starting All In One Validation Formatter...
cd /d "%~dp0"
python -m streamlit run app.py
pause
