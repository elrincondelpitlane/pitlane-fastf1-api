import argparse
import json
import os
from datetime import datetime, timezone
from pathlib import Path

import fastf1
import pandas as pd


FASTF1_CACHE_DIR = Path("cache") / "fastf1"
PROCESSED_DIR = Path("processed")


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


def serialize_fastest_lap(driver, lap):
    return {
        "driver": driver,
        "lapNumber": safe_value(lap.get("LapNumber")),
        "lapTime": format_lap_time(lap.get("LapTime")),
        "lapDuration": timedelta_to_seconds(lap.get("LapTime")),
        "sector1": timedelta_to_seconds(lap.get("Sector1Time")),
        "sector2": timedelta_to_seconds(lap.get("Sector2Time")),
        "sector3": timedelta_to_seconds(lap.get("Sector3Time")),
        "compound": safe_value(lap.get("Compound")),
        "tyreLife": safe_value(lap.get("TyreLife")),
    }


def serialize_telemetry_point(row):
    return {
        "distance": safe_value(row.get("Distance")),
        "speed": safe_value(row.get("Speed")),
        "throttle": safe_value(row.get("Throttle")),
        "brake": safe_value(row.get("Brake")),
        "rpm": safe_value(row.get("RPM")),
        "gear": safe_value(row.get("nGear")),
        "timeSeconds": timedelta_to_seconds(row.get("Time")),
    }


def write_json(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
        file.write("\n")


def load_session(year, gp, session_code):
    FASTF1_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    fastf1.Cache.enable_cache(str(FASTF1_CACHE_DIR))

    session = fastf1.get_session(year, gp, session_code)
    try:
        session.load(
            laps=True,
            telemetry=True,
            weather=False,
            messages=False,
        )
    except TypeError:
        session.load()

    if session.laps is None or session.laps.empty:
        raise RuntimeError("No se cargaron vueltas para esta sesión")

    return session


def generated_at():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def build_session_view(year, gp, session_code, laps, generated_timestamp):
    drivers = []
    for driver in sorted(laps["Driver"].dropna().unique()):
        driver_laps = laps.pick_driver(driver).sort_values("LapNumber")
        drivers.append({
            "code": driver,
            "laps": [serialize_lap(lap) for _, lap in driver_laps.iterrows()],
        })

    return {
        "year": year,
        "gp": gp,
        "session": session_code,
        "generated_at": generated_timestamp,
        "drivers": drivers,
    }


def build_fastest_laps(year, gp, session_code, laps, generated_timestamp):
    fastest_laps = []
    quick_laps = laps.pick_quicklaps()

    for driver in sorted(quick_laps["Driver"].dropna().unique()):
        driver_laps = quick_laps.pick_driver(driver)
        if driver_laps.empty:
            continue

        fastest_lap = driver_laps.pick_fastest()
        if fastest_lap is None:
            continue

        fastest_laps.append(serialize_fastest_lap(driver, fastest_lap))

    return {
        "year": year,
        "gp": gp,
        "session": session_code,
        "generated_at": generated_timestamp,
        "laps": fastest_laps,
    }


def build_telemetry_payload(year, gp, session_code, driver, lap):
    lap_number = safe_value(lap.get("LapNumber"))
    car_data = lap.get_car_data().add_distance()

    return {
        "year": year,
        "gp": gp,
        "session": session_code,
        "driver": driver,
        "lapNumber": lap_number,
        "lapDuration": timedelta_to_seconds(lap.get("LapTime")),
        "points": [serialize_telemetry_point(row) for _, row in car_data.iterrows()],
    }


def preprocess_session(year, gp, session_code):
    session = load_session(year, gp, session_code)
    session_dir = PROCESSED_DIR / "sessions" / str(year) / gp / session_code
    telemetry_dir = PROCESSED_DIR / "telemetry" / str(year) / gp / session_code
    session_dir.mkdir(parents=True, exist_ok=True)
    telemetry_dir.mkdir(parents=True, exist_ok=True)
    generated_timestamp = generated_at()

    session_view = build_session_view(year, gp, session_code, session.laps, generated_timestamp)
    write_json(session_dir / "session-view.json", session_view)

    fastest_laps = build_fastest_laps(year, gp, session_code, session.laps, generated_timestamp)
    write_json(session_dir / "fastest-laps.json", fastest_laps)

    quick_laps = session.laps.pick_quicklaps()
    for fastest_lap in fastest_laps["laps"]:
        driver = fastest_lap["driver"]
        lap_number = fastest_lap["lapNumber"]
        if lap_number is None:
            continue

        driver_laps = quick_laps.pick_driver(driver)
        lap_matches = driver_laps[driver_laps["LapNumber"] == lap_number]
        if lap_matches.empty:
            continue

        telemetry = build_telemetry_payload(year, gp, session_code, driver, lap_matches.iloc[0])
        write_json(telemetry_dir / f"{driver}_L{lap_number}.json", telemetry)

    return session_dir, telemetry_dir


def parse_args():
    parser = argparse.ArgumentParser(description="Preprocess FastF1 session JSON files.")
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--gp", required=True)
    parser.add_argument("--session", required=True)
    return parser.parse_args()


def main():
    args = parse_args()
    session_dir, telemetry_dir = preprocess_session(args.year, args.gp, args.session)
    print(f"Session JSON written to {os.fspath(session_dir)}")
    print(f"Telemetry JSON written to {os.fspath(telemetry_dir)}")


if __name__ == "__main__":
    main()
