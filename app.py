import os

import fastf1
import pandas as pd
from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

FASTF1_CACHE_DIR = os.path.join("cache", "fastf1")
os.makedirs(FASTF1_CACHE_DIR, exist_ok=True)
fastf1.Cache.enable_cache(FASTF1_CACHE_DIR)


def safe_value(value):
    if pd.isna(value):
        return None

    if hasattr(value, "item"):
        value = value.item()

    if isinstance(value, float) and value.is_integer():
        return int(value)

    return value


def timedelta_to_seconds(value):
    if pd.isna(value):
        return None

    if hasattr(value, "total_seconds"):
        return value.total_seconds()

    return safe_value(value)


def format_lap_time(value):
    if pd.isna(value):
        return None

    total_seconds = timedelta_to_seconds(value)
    if total_seconds is None:
        return None

    minutes = int(total_seconds // 60)
    seconds = total_seconds - (minutes * 60)
    return f"{minutes}:{seconds:06.3f}"


def serialize_lap(lap):
    return {
        "lapNumber": safe_value(lap.get("LapNumber")),
        "lapTime": format_lap_time(lap.get("LapTime")),
        "lapDuration": timedelta_to_seconds(lap.get("LapTime")),
        "sector1": timedelta_to_seconds(lap.get("Sector1Time")),
        "sector2": timedelta_to_seconds(lap.get("Sector2Time")),
        "sector3": timedelta_to_seconds(lap.get("Sector3Time")),
        "compound": safe_value(lap.get("Compound")),
        "stint": safe_value(lap.get("Stint")),
        "tyreLife": safe_value(lap.get("TyreLife")),
        "isPersonalBest": safe_value(lap.get("IsPersonalBest")),
    }


def load_fastf1_session(year, race, session_code):
    session = fastf1.get_session(year, race, session_code)
    session.load()
    return session


def session_load_error(error):
    return jsonify({
        "status": "error",
        "message": "No se pudo cargar la sesión",
        "detail": str(error)
    }), 500

@app.route("/")
def home():
    return jsonify({
        "status": "ok",
        "message": "Pitlane FastF1 API funcionando"
    })

@app.route("/race-result/<int:year>/<race>")
def race_result(year, race):
    try:
        session = load_fastf1_session(year, race, "R")

        results = session.results

        data = []

        for _, row in results.iterrows():
            data.append({
                "position": str(row.get("Position", "")),
                "driver": str(row.get("FullName", "")),
                "abbreviation": str(row.get("Abbreviation", "")),
                "team": str(row.get("TeamName", "")),
                "points": str(row.get("Points", ""))
            })

        return jsonify({
            "year": year,
            "race": race,
            "session": "Race",
            "results": data
        })

    except Exception as e:
        return session_load_error(e)


@app.route("/qualifying-result/<int:year>/<race>")
def qualifying_result(year, race):
    try:
        session = load_fastf1_session(year, race, "Q")
        results = session.results

        data = []
        for _, row in results.iterrows():
            data.append({
                "position": str(row.get("Position", "")),
                "driver": str(row.get("FullName", "")),
                "abbreviation": str(row.get("Abbreviation", "")),
                "team": str(row.get("TeamName", "")),
            })

        return jsonify({
            "year": year,
            "race": race,
            "session": "Qualifying",
            "results": data
        })

    except Exception as e:
        return session_load_error(e)


@app.route("/fastest-laps/<int:year>/<race>")
def fastest_laps_race(year, race):
    return fastest_laps(year, race, "R")


@app.route("/fastest-laps/<int:year>/<race>/<session_code>")
def fastest_laps(year, race, session_code):
    try:
        session = load_fastf1_session(year, race, session_code)
        laps = session.laps.pick_quicklaps()
        requested_drivers = request.args.get("drivers")

        if requested_drivers:
            drivers = [driver.strip().upper() for driver in requested_drivers.split(",") if driver.strip()]
        else:
            drivers = sorted(laps["Driver"].dropna().unique())

        data = []
        for driver in drivers:
            driver_laps = laps.pick_driver(driver)
            if driver_laps.empty:
                continue

            fastest_lap = driver_laps.pick_fastest()
            if fastest_lap is None:
                continue

            data.append({
                "driver": driver,
                "lapNumber": safe_value(fastest_lap.get("LapNumber")),
                "lapTime": format_lap_time(fastest_lap.get("LapTime")),
                "lapDuration": timedelta_to_seconds(fastest_lap.get("LapTime")),
                "sector1": timedelta_to_seconds(fastest_lap.get("Sector1Time")),
                "sector2": timedelta_to_seconds(fastest_lap.get("Sector2Time")),
                "sector3": timedelta_to_seconds(fastest_lap.get("Sector3Time")),
                "compound": safe_value(fastest_lap.get("Compound")),
                "tyreLife": safe_value(fastest_lap.get("TyreLife")),
            })

        return jsonify({
            "year": year,
            "race": race,
            "session": session_code,
            "laps": data
        })

    except Exception as e:
        return session_load_error(e)


@app.route("/session-view/<int:year>/<race>/<session_code>")
def session_view(year, race, session_code):
    try:
        session = load_fastf1_session(year, race, session_code)
        laps = session.laps
        drivers = []

        for driver in sorted(laps["Driver"].dropna().unique()):
            driver_laps = laps.pick_driver(driver).sort_values("LapNumber")
            drivers.append({
                "code": driver,
                "laps": [serialize_lap(lap) for _, lap in driver_laps.iterrows()]
            })

        return jsonify({
            "year": year,
            "race": race,
            "session": session_code,
            "drivers": drivers
        })

    except Exception as e:
        return session_load_error(e)


@app.route("/telemetry/<int:year>/<race>/<session_code>/<driver>/<int:lap_number>")
def telemetry(year, race, session_code, driver, lap_number):
    try:
        session = load_fastf1_session(year, race, session_code)
    except Exception as e:
        return session_load_error(e)

    driver_laps = session.laps.pick_driver(driver.upper())
    lap_matches = driver_laps[driver_laps["LapNumber"] == lap_number]

    if lap_matches.empty:
        return jsonify({
            "status": "error",
            "message": "Vuelta no encontrada"
        }), 404

    lap = lap_matches.iloc[0]

    try:
        car_data = lap.get_car_data().add_distance()
    except Exception as e:
        return session_load_error(e)

    points = []
    for _, row in car_data.iterrows():
        points.append({
            "distance": safe_value(row.get("Distance")),
            "speed": safe_value(row.get("Speed")),
            "throttle": safe_value(row.get("Throttle")),
            "brake": safe_value(row.get("Brake")),
            "rpm": safe_value(row.get("RPM")),
            "gear": safe_value(row.get("nGear")),
            "timeSeconds": timedelta_to_seconds(row.get("Time")),
        })

    return jsonify({
        "year": year,
        "race": race,
        "session": session_code,
        "driver": driver.upper(),
        "lapNumber": lap_number,
        "points": points
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
