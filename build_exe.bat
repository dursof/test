@echo off
REM ============================================================
REM  Build script per extract_ddt.exe (Windows)
REM  Eseguire da prompt dei comandi nella cartella del progetto
REM ============================================================

echo === Installazione dipendenze ===
pip install -r requirements.txt
pip install pyinstaller

echo.
echo === Build eseguibile ===
pyinstaller --onefile --console --name extract_ddt extract_ddt.py

echo.
echo === Fine build ===
echo L'eseguibile si trova in: dist\extract_ddt.exe
echo.
echo IMPORTANTE: Su Windows e' necessario installare Tesseract OCR
echo   Scarica da: https://github.com/UB-Mannheim/tesseract/wiki
echo   Durante l'installazione seleziona anche: "Italian" (lingua)
echo.
pause
