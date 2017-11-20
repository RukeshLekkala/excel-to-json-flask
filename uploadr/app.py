from flask import Flask, request, redirect, url_for, render_template

import os
import json
import glob

from openpyxl import Workbook
from openpyxl import load_workbook

app = Flask(__name__)

# DIR_JSON = os.path.join("json/")

# READ ALL ROWS IN WORKSHEET AND TRANSFORM INTO JSON
def all_data_to_json(worksheet, filename, sheetname):
    # with open('{}{}_{}.json'.format(DIR_JSON,filename,sheetname), 'w') as file:
    with open('{}_{}.json'.format(filename,sheetname), 'w') as file:
        json_data = []

        for row in range(1, worksheet.max_row):
            item = {}
            for column in range(worksheet.max_column):
                try:
                    item[worksheet.cell(row=1, column=column+1).value.upper()] = worksheet.cell(row=row+1, column=column+1).value.encode('utf-8')
                except:
                    try:
                        item[worksheet.cell(row=1, column=column+1).value.upper()] = worksheet.cell(row=row+1, column=column+1).value
                    except:
                        item[worksheet.cell(row=1, column=column+1).value] = worksheet.cell(row=row+1, column=column+1).value
            json_data.append(item)

        json.dump(json_data, file, indent = 4, ensure_ascii = False) # sort_keys = True
        file.close()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    form = request.form

    # Is the upload using Ajax, or a direct POST by the form?
    is_ajax = False
    if form.get("__ajax", None) == "true":
        is_ajax = True

    # Target folder for these uploads.
    # target = "uploadr/static/uploads/{}".format(upload_key)
    # try:
    #     os.mkdir(target)
    # except:
    #     if is_ajax:
    #         return ajax_response(False, "Couldn't create upload directory: {}".format(target))
    #     else:
    #         return "Couldn't create upload directory: {}".format(target)

    for upload in request.files.getlist("file"):
        try:
            filename = upload.filename.rsplit("/")[0]
            filename =  filename.split('.')[0]

            wb = load_workbook(filename=upload, read_only=True)
            # filename = file.split('.')

            sheets = wb.get_sheet_names()
            for sheet in sheets:
                ws = wb[sheet]
                
                all_data_to_json(ws, filename, sheet)

            return render_template('index.html')
        except:
            return render_template('index.html')

    # try:
    #     return render_template('index.html')
    # except:
    #     return render_template('index.html')

@app.route('/search')
def search():
    return render_template('search.html')
