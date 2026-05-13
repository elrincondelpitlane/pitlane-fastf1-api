from flask import Flask, jsonify
import fastf1

app = Flask(__name__)

@app.route("/")
def home():
    return jsonify({
        "status": "ok",
        "message": "Pitlane FastF1 API funcionando"
    })

@app.route("/race-result/<int:year>/<race>")
def race_result(year, race):
    session = fastf1.get_session(year, race, "R")
    session.load()

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

@app.route("/qualifying-result/<int:year>/<race>")
def qualifying_result(year, race):
    session = fastf1.get_session(year, race, "Q")
    session.load()

    results = session.results

    data = []
    for _, row in results.iterrows():
        data.append({
            "position": str(row.get("Position", "")),
            "driver": str(row.get("FullName", "")),
            "abbreviation": str(row.get("Abbreviation", "")),
            "team": str(row.get("TeamName", "")),
            "q1": str(row.get("Q1", "")),
            "q2": str(row.get("Q2", "")),
            "q3": str(row.get("Q3", ""))
        })

    return jsonify({
        "year": year,
        "race": race,
        "session": "Qualifying",
        "results": data
    })

@app.route("/fastest-laps/<int:year>/<race>")
def fastest_laps(year, race):
    session = fastf1.get_session(year, race, "R")
    session.load()

    laps = session.laps.pick_quicklaps()
    fastest = laps.sort_values("LapTime").head(10)

    data = []
    for _, lap in fastest.iterrows():
        data.append({
            "driver": str(lap.get("Driver", "")),
            "lap_number": int(lap.get("LapNumber", 0)),
            "lap_time": str(lap.get("LapTime", "")),
            "compound": str(lap.get("Compound", "")),
            "tyre_life": str(lap.get("TyreLife", ""))
        })

    return jsonify({
        "year": year,
        "race": race,
        "fastest_laps": data
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)