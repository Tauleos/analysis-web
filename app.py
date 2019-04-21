# -*- coding: utf-8 -*-
import json
import os
import time
from service import ExcelExe
from flask import Flask, render_template, request, make_response
from flask_cors import CORS

app = Flask(__name__)
CORS(app)


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


@app.route('/mg', methods=['POST'])
def upload_img():
    if request.method == 'POST':
        f = request.files['image']
        ext = f.filename.rsplit('.', 1)[1].lower()
        path = '/static/' + str(round(time.time() * 1000)) + '.' + ext
        url = os.getcwd() + path
        f.save(url)
        make_response()
        return bytes('http://127.0.0.1:5000' + path, 'utf-8')
