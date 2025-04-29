
from flask import Flask, render_template, request, send_file, redirect, url_for, send_from_directory
import pandas as pd
import os
import json
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
STATS_FILE = "data.json"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def load_data():
    if not os.path.exists(STATS_FILE):
        return {"total": 0, "dates": {}}
    with open(STATS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    with open(STATS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

data_store = load_data()

@app.route("/", methods=["GET"])
def index():
    today = datetime.now().strftime("%Y-%m-%d")
    total = data_store.get("total", 0)
    today_count = data_store.get("dates", {}).get(today, 0)

    latest_files = sorted(os.listdir(RESULT_FOLDER), key=lambda x: os.path.getmtime(os.path.join(RESULT_FOLDER, x)), reverse=True)[:5]
    return render_template("index.html", stats={"total": total, "today": today_count}, files=latest_files)

@app.route("/process", methods=["POST"])
def process():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".xlsx"):
        return redirect(url_for("index"))

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(RESULT_FOLDER, f"✅Taxrirlangan fayl_{file.filename}")
    file.save(filepath)

    try:
        df = pd.read_excel(filepath, header=None)
        df = df.iloc[9:].reset_index(drop=True)
        df = df.drop(index=1).reset_index(drop=True)
        selected_columns = df.iloc[:, [1, 5, 7]].copy()
        selected_columns.iloc[1:, 2] = "Y"

        last_row_index = selected_columns.iloc[:, 0].last_valid_index()
        if last_row_index is not None:
            start = last_row_index + 1
            end = start + 10
            selected_columns = selected_columns.drop(index=range(start, end), errors="ignore")

        selected_columns.to_excel(output_path, index=False, header=False)

        today = datetime.now().strftime("%Y-%m-%d")
        data_store["total"] = data_store.get("total", 0) + 1
        data_store.setdefault("dates", {})
        data_store["dates"][today] = data_store["dates"].get(today, 0) + 1
        save_data(data_store)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error: {str(e)}"

@app.route("/open/<filename>")
def open_file(filename):
    return send_from_directory(RESULT_FOLDER, filename)

@app.route("/reset", methods=["POST"])
def reset():
    if os.path.exists(STATS_FILE):
        os.remove(STATS_FILE)
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
