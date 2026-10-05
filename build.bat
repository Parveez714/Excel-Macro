@echo off
python -m PyInstaller --noconfirm --onefile --windowed --name ExcelMacroAssistant --hidden-import macro_assistant.providers.gemini_provider --hidden-import macro_assistant.providers.openai_provider main.py
if errorlevel 1 exit /b %errorlevel%
echo Build complete: dist\ExcelMacroAssistant.exe
