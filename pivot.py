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
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    tmpdir = tempfile.gettempdir()
    unique_id = str(uuid.uuid4())
    output_path = os.path.join(tmpdir, f"pivot_{unique_id}.xlsx")

    try:
        dfs = []

        for f in files:
            if not f.filename.lower().endswith('.xls'):
                return jsonify({"error": f"File {f.filename} is not .xls"}), 400

            tmp_input = os.path.join(tmpdir, f.filename)
            f.save(tmp_input)

            book = xlrd.open_workbook(tmp_input, logfile=open(os.devnull, 'w'))
            sheet = book.sheet_by_index(0)

            data = [sheet.row_values(r) for r in range(sheet.nrows)]
            df = pd.DataFrame(data)

            # first row as headers
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)

            # --- Normalize headers ---
            rename_map = {
                "Product ID": "Product ID",
                "Branch ID": "Branch ID",
                "Qty On Hand": "Qty On Hand",
                "Branch ID ": "Branch ID",  # in case of trailing space
                " Branch ID": "Branch ID",  # in case of leading space
            }
            df = df.rename(columns=rename_map)

            # Keep only the 3 relevant columns
            needed = ["Product ID", "Branch ID", "Qty On Hand"]
            df = df[[c for c in df.columns if c in needed]]

            dfs.append(df)

        merged = pd.concat(dfs, ignore_index=True)

        # Ensure correct dtypes
        merged["Qty On Hand"] = pd.to_numeric(merged["Qty On Hand"], errors="coerce").fillna(0)

        # Build pivot
        pivot = pd.pivot_table(
            merged,
            index=["Product ID"],
            columns=["Branch ID"],
            values="Qty On Hand",
            aggfunc="sum",
            fill_value=0
        )

        # Save pivot to Excel
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
