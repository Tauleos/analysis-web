# -*- coding: utf-8 -*-
import json
import os
import time
from service import ExcelExe
from flask import Flask, render_template, request, make_response, send_from_directory, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)


@app.route('/')
def main_page():
    return render_template('index.html')


@app.route('/server/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        f = request.files['file']
        f.save(os.getcwd() + '/uploads/' + secure_filename(f.filename))
        obj = json.dumps({'success': True, 'filename': secure_filename(f.filename)})
        return bytes(obj, 'utf-8')


@app.route('/server/exe', methods=['POST'])
def exe():
    filename = request.get_json()['filename']
    path = os.getcwd() + '/uploads/' + secure_filename(filename)
    desc = ExcelExe().execute(path)
    return bytes(desc[1:], 'utf-8')


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


@app.route('/server/excel/<filename>')
def favicon(filename):
    return send_from_directory(os.path.join(app.root_path, 'static'), filename)


@app.route('/manifest.json')
def manifest():
    return send_from_directory(os.path.join(app.root_path, 'static'), 'manifest.json')
