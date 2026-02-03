@echo off
setlocal
set PYTHONUTF8=1
set PLAYWRIGHT_HEADLESS=0
if not exist ".venv" (
  python -m venv .venv
)
call .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --noconfirm --onefile --name "ROBOT_AUDIT_SRI" ^
  --collect-all streamlit ^
  --collect-all streamlit.runtime ^
  --collect-all streamlit.elements ^
  --collect-all streamlit.components.v1 ^
  desktop_launcher.py
endlocal
