from flask import Flask, render_template, request, send_file, redirect, url_for, send_from_directory
import pandas as pd
import os
import json
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "uploads"
app.config['RESULT_FOLDER'] = "results"
app.config['HTML_FOLDER'] = "html_previews"
app.config['STATS_FILE'] = "data.json"
app.config['TEMPLATE_FILE'] = "preview_template.html"

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)
os.makedirs(app.config['HTML_FOLDER'], exist_ok=True)

def load_data():
    if not os.path.exists(app.config['STATS_FILE']):
        return {"total": 0, "dates": {}, "show_count": 3}
    with open(app.config['STATS_FILE'], "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    with open(app.config['STATS_FILE'], "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

data_store = load_data()

def format_file_time(timestamp):
    dt = datetime.fromtimestamp(timestamp)
    return dt.strftime("%d.%m.%Y %H:%M:%S")

@app.route("/", methods=["GET", "POST"])
def index():
    today = datetime.now().strftime("%Y-%m-%d")
    total = data_store.get("total", 0)
    today_count = data_store.get("dates", {}).get(today, 0)
    
    if request.method == "POST" and "show_count" in request.form:
        data_store["show_count"] = int(request.form["show_count"])
        save_data(data_store)
    
    show_count = data_store.get("show_count", 3)

    files_with_time = []
    for filename in os.listdir(app.config['RESULT_FOLDER']):
        filepath = os.path.join(app.config['RESULT_FOLDER'], filename)
        mtime = os.path.getmtime(filepath)
        preview_name = filename.replace('✅Taxrirlangan_fayl_', 'preview_').replace('.xlsx', '.html')
        files_with_time.append({
            'name': filename,
            'preview_name': preview_name,
            'time': format_file_time(mtime),
            'timestamp': mtime
        })
    
    latest_files = sorted(files_with_time, key=lambda x: x['timestamp'], reverse=True)[:show_count]
    
    return render_template("index.html", 
                         stats={"total": total, "today": today_count}, 
                         files=latest_files,
                         show_count=show_count)

@app.route("/process", methods=["POST"])
def process():
    if 'file' not in request.files:
        return redirect(url_for("index"))
    
    file = request.files['file']
    if file.filename == '':
        return redirect(url_for("index"))
    
    if not file.filename.endswith('.xlsx'):
        return "Faqat .xlsx formatidagi fayllarni yuklash mumkin"

    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        
        df = pd.read_excel(filepath, header=None)
        df = df.iloc[9:].reset_index(drop=True)
        df = df.drop(index=1).reset_index(drop=True)
        selected_columns = df.iloc[:, [1, 5, 7]].copy()
        selected_columns.iloc[1:, 2] = "Y"

        last_row_index = selected_columns.iloc[:, 0].last_valid_index()
        if last_row_index is not None:
            selected_columns = selected_columns.iloc[:last_row_index+1]

        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"✅Taxrirlangan_fayl_{current_time}.xlsx"
        output_path = os.path.join(app.config['RESULT_FOLDER'], output_filename)
        selected_columns.to_excel(output_path, index=False, header=False)

        html_filename = f"preview_{current_time}.html"
        html_preview_path = os.path.join(app.config['HTML_FOLDER'], html_filename)
        
        table_html = selected_columns.copy()
        table_html.index += 1
        table_html.reset_index(inplace=True)
        table_html.columns = ["#", "Цена", "Количество", "Код"]
        html_code = table_html.to_html(index=False, escape=False, classes="table table-striped")

        with open(app.config['TEMPLATE_FILE'], "r", encoding="utf-8") as f:
            template = f.read()
        
        formatted_time = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        html_content = template.replace("{table}", html_code).replace("{date}", formatted_time)
        
        with open(html_preview_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        today = datetime.now().strftime("%Y-%m-%d")
        data_store["total"] = data_store.get("total", 0) + 1
        data_store.setdefault("dates", {})
        data_store["dates"][today] = data_store["dates"].get(today, 0) + 1
        save_data(data_store)

        return {
            'success': True,
            'filename': output_filename,
            'preview_name': html_filename,
            'time': formatted_time
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }

@app.route("/get_updated_files", methods=["GET"])
def get_updated_files():
    show_count = data_store.get("show_count", 3)
    
    files_with_time = []
    for filename in os.listdir(app.config['RESULT_FOLDER']):
        filepath = os.path.join(app.config['RESULT_FOLDER'], filename)
        mtime = os.path.getmtime(filepath)
        preview_name = filename.replace('✅Taxrirlangan_fayl_', 'preview_').replace('.xlsx', '.html')
        files_with_time.append({
            'name': filename,
            'preview_name': preview_name,
            'time': format_file_time(mtime),
            'timestamp': mtime
        })
    
    latest_files = sorted(files_with_time, key=lambda x: x['timestamp'], reverse=True)[:show_count]
    
    return {'files': latest_files}

@app.route("/clear_files", methods=["POST"])
def clear_files():
    try:
        # Faqat fayllarni o'chirish
        for filename in os.listdir(app.config['RESULT_FOLDER']):
            file_path = os.path.join(app.config['RESULT_FOLDER'], filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"{file_path} o'chirishda xato: {e}")
        
        for filename in os.listdir(app.config['HTML_FOLDER']):
            file_path = os.path.join(app.config['HTML_FOLDER'], filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"{file_path} o'chirishda xato: {e}")
        
        return {'success': True}
    except Exception as e:
        return {'success': False, 'error': str(e)}

@app.route("/preview/<filename>")
def preview_file(filename):
    if not filename.endswith('.html'):
        filename += '.html'
    try:
        return send_from_directory(app.config['HTML_FOLDER'], filename)
    except FileNotFoundError:
        return "Ko'rish uchun fayl topilmadi. Iltimos, faylni qayta ishlang."

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)

@app.route("/reset_stats", methods=["POST"])
def reset_stats():
    try:
        today = datetime.now().strftime("%Y-%m-%d")
        data_store["total"] = 0
        data_store["dates"] = {today: 0}
        save_data(data_store)
        return {'success': True}
    except Exception as e:
        return {'success': False, 'error': str(e)}

if __name__ == "__main__":
    app.run(debug=True)