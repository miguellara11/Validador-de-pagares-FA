# Validador de pagarés

Aplicación de escritorio en Python que compara un reporte de préstamos en Excel contra pagarés en PDF y marca cada celda por color. Reduce la revisión quincenal de 200-300 pagarés de medio día a minutos.

- **Verde / rojo / amarillo:** el dato coincide / no coincide / no aparece en el PDF (ID, nombre, capital, monto quincenal, interés y quincena).
- **Avales:** detecta si hay firma sobre la línea FIRMA (verde con firma, rojo sin firma).

**Stack:** Python, PyMuPDF (lectura de PDF y detección visual de firma), openpyxl (Excel), Tkinter (interfaz), PyInstaller (.exe).

## Uso
```
cd desktop
pip install -r requirements.txt
python validador_pagares.py
```
Para generar el `.exe` en Windows: ejecuta `desktop/build.bat`.

## Privacidad y limitaciones
Todo se procesa localmente. El repositorio no incluye datos reales (PDF y Excel están en `.gitignore`). La detección de firma es una heurística visual: revisa a mano los casos dudosos.
