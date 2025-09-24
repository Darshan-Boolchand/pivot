from flask import Flask, request, send_file, jsonify, after_this_request
import pandas as pd
import xlrd
import os
import tempfile
import traceback
import uuid

app = Flask(__name__)

@app.route('/pivot', methods=['POST'])
def pivot_service():
    if 'files' not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files")
    if len(files) != 5:
        return jsonify({"error": "Exactly 5 .xls files are required"}), 400

    tmpdir = tempfile.gettempdir()
    unique_id = str(uuid.uuid4())
    output_path = os.path.join(tmpdir, f"pivot_{unique_id}.xlsx")

    try:
        dfs = []

        # Convert each .xls -> DataFrame
        for f in files:
            if not f.filename.lower().endswith('.xls'):
                return jsonify({"error": f"File {f.filename} is not .xls"}), 400

            tmp_input = os.path.join(tmpdir, f.filename)
            f.save(tmp_input)

            book = xlrd.open_workbook(tmp_input, logfile=open(os.devnull, 'w'))
            sheet = book.sheet_by_index(0)

            data = [sheet.row_values(r) for r in range(sheet.nrows)]
            df = pd.DataFrame(data)

            dfs.append(df)

        # Merge all 5 DataFrames
        merged = pd.concat(dfs, ignore_index=True)

        # First row = headers
        merged.columns = merged.iloc[0]
        merged = merged.drop(0).reset_index(drop=True)

        # ⚡ Build Pivot (adjust index/columns/values as per your Stock_Master.xlsx)
        # Example: pivot sales quantity by Store and Product
        pivot = pd.pivot_table(
            merged,
            index=["Store"],          # <-- adjust
            columns=["Product"],      # <-- adjust
            values="Quantity",        # <-- adjust
            aggfunc="sum",
            fill_value=0
        )

        # Save only pivot to Excel
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            pivot.to_excel(writer, sheet_name="Pivot")

        @after_this_request
        def cleanup(response):
            try:
                os.remove(output_path)
            except Exception as e:
                print("Cleanup error:", e)
            return response

        return send_file(output_path, as_attachment=True, download_name="pivot.xlsx")

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/')
def index():
    return "Pivot webservice is running."
