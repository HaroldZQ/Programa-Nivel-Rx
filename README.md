# APU SPACE // Mission Control

Aplicación de escritorio para analizar telemetría de vuelos OpenRocket y calcular el link budget de un sistema LoRa.

## Funciones

- Carga de archivos CSV de OpenRocket.
- Cálculo de distancia al punto de estación terrena, FSPL, potencia recibida y margen de enlace.
- Exportación de resultados a Excel.
- Dashboard PNG y dashboard HTML interactivo con zoom, hover y descarga.
- Interfaz visual inspirada en una consola de control aeroespacial.

## Ejecución

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app_link_budget_gui.py
```

También puede ejecutarse el archivo generado `dist/app_link_budget_gui.exe` en Windows.

## Salidas

Los archivos generados se guardan junto al CSV seleccionado:

- `*_LinkBudget_Calculado.xlsx`
- `*_Dashboard_Graficas.png`
- `*_Dashboard_Graficas_Interactivo.html`
