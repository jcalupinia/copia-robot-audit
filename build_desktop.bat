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
python scripts\generate_icon.py
python scripts\generate_version.py
set ICON_ARG=
if exist "LogoAUDIT.ico" (
  set ICON_ARG=--icon "LogoAUDIT.ico"
)
pyinstaller --noconfirm --onefile --name "ROBOT_AUDIT_SRI" ^
  --noupx ^
  %ICON_ARG% ^
  --collect-all streamlit ^
  --collect-all streamlit.runtime ^
  --collect-all streamlit.elements ^
  --collect-all streamlit.components.v1 ^
  --collect-all playwright ^
  --add-data "aplicacion.py;." ^
  --add-data "AUDIT_IA_sin_fondo_transparente_FINAL.png;." ^
  --add-data "licensing_client.py;." ^
  --add-data "version.txt;." ^
  --add-data "robot;robot" ^
  --add-data "licensing_api;licensing_api" ^
  --add-data ".streamlit;.streamlit" ^
  desktop_launcher.py
if /I "%PUBLISH_UPDATE%"=="1" (
  python scripts\publish_update.py
)
endlocal
