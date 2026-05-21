# Pitlane FastF1 API

API Flask para servir datos livianos de FastF1 al Race Center.

## Instalar

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Correr

```powershell
python app.py
```

URL local:

http://localhost:5000

## Probar

http://localhost:5000/

http://localhost:5000/session-view/2026/Miami/R

http://localhost:5000/fastest-laps/2026/Miami/R?drivers=ANT,LEC

http://localhost:5000/telemetry/2026/Miami/R/ANT/34

## Endpoints

- `GET /`
- `GET /race-result/<year>/<race>`
- `GET /qualifying-result/<year>/<race>`
- `GET /fastest-laps/<year>/<race>`
- `GET /fastest-laps/<year>/<race>/<session_code>`
- `GET /session-view/<year>/<race>/<session_code>`
- `GET /telemetry/<year>/<race>/<session_code>/<driver>/<lap_number>`

La cache de FastF1 se guarda en `cache/fastf1/` y no se versiona.
