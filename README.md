# Pitlane FastF1 API

API Flask para servir JSON livianos de FastF1 al Race Center.

Los endpoints de session view, fastest laps y telemetria no procesan FastF1 por request. Primero leen archivos JSON preprocesados en `processed/`. Si la sesion todavia no fue procesada, devuelven `status: not_processed`.

## Flujo local

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python scripts/preprocess_session.py --year 2026 --gp Miami --session R
python scripts/upload_to_supabase.py --year 2026 --gp Miami --session R
```

## Correr API local

```powershell
python app.py
```

URL local:

http://localhost:5000

## GitHub Actions

El workflow manual esta en `.github/workflows/preprocess-fastf1.yml`.

Para ejecutarlo:

GitHub -> Actions -> Preprocess FastF1 -> Run workflow

Inputs:

- `year`
- `gp`
- `session`

Ejemplo:

- `year`: `2026`
- `gp`: `Miami`
- `session`: `R`

## Secrets necesarios

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_FASTF1_BUCKET`

## Bucket necesario

Crear un bucket de Supabase Storage llamado:

```text
fastf1-cache
```

Los archivos se suben con esta estructura:

```text
fastf1-cache/
  2026/
    Miami/
      R/
        session-view.json
        fastest-laps.json
        telemetry/
          ANT_L34.json
```

## Endpoints

- `GET /`
- `GET /race-result/<year>/<race>`
- `GET /qualifying-result/<year>/<race>`
- `GET /fastest-laps/<year>/<race>`
- `GET /fastest-laps/<year>/<race>/<session_code>`
- `GET /session-view/<year>/<race>/<session_code>`
- `GET /telemetry/<year>/<race>/<session_code>/<driver>/<lap_number>`

La cache de FastF1 se guarda en `cache/fastf1/` y no se versiona.
