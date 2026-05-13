from flask import Flask, jsonify
import fastf1

app = Flask(__name__)

@app.route("/")
def home():
    return jsonify({
        "status": "ok",
        "message": "Pitlane FastF1 API funcionando"
    })

@app.route("/race-result/<int:year>/<int:round_number>")
def race_result(year, round_number):
    try:
        session = fastf1.get_session(year, round_number, "R")
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
            "round": round_number,
            "session": "Race",
            "results": data
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
