from flask import Flask, send_from_directory, jsonify
import os
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.wrappers import Response

app = Flask(__name__, static_folder="../")

@app.route('/excel-files')
def list_excel_files():
    files = [f for f in os.listdir(app.static_folder) if f.endswith('.xlsx')]
    return jsonify(files)

@app.route('/<path:filename>')
def serve_file(filename):
    return send_from_directory(app.static_folder, filename)

def handler(environ, start_response):
    return DispatcherMiddleware(lambda environ, start_response: Response("Not Found", status=404), {
        "/api": app
    })(environ, start_response)
