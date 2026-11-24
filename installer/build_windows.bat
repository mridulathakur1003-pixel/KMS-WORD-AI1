@echo off
setlocal
cd /d "%~dp0\.."
py -m pip install --upgrade pip
py -m pip install -r installer\requirements-build.txt
py -m PyInstaller --noconfirm --clean --windowed --onedir --name "KMS WORD AI" ^
  --add-data "kms_config.json;." KMS_WORD_AI_V20_PRO.py
if errorlevel 1 (
  echo Build failed. Review the errors above.
  pause
  exit /b 1
)
echo.
echo Build complete: dist\KMS WORD AI\
echo Next: install Inno Setup 6 and compile installer\KMS_Word_AI_Setup.iss
pause
