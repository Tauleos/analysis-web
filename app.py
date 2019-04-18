# -*- coding: utf-8 -*-
import json
from service import ExcelExe
from flask import Flask, render_template, request

app = Flask(__name__)


@app.route('/')
def main_page():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        f = request.files['file']
        # f.save(os.getcwd()+'/uploads/'+secure_filename(f.filename))
        desc = ExcelExe().execute(f)
    obj = json.dumps({'success': True, 'url': desc[1:]})
    return bytes(obj, 'utf-8')
