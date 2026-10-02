@echo off
REM Genera dist\ValidadorPagares.exe en un entorno limpio (opcional: logo.png junto a este archivo)
python -m venv venv_pagares
venv_pagares\Scripts\python -m pip install -r requirements.txt pyinstaller
set ADD=
if exist logo.png set ADD=--add-data "logo.png;."
venv_pagares\Scripts\python -m PyInstaller --onefile --windowed --name ValidadorPagares %ADD% validador_pagares.py
pause
