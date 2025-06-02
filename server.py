import os
from flask import Flask, send_from_directory, jsonify

app = Flask(__name__, static_folder='.')

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/excel-files')
def list_excel_files():
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    return jsonify(files)

@app.route('/<path:filename>')
def serve_file(filename):
    return send_from_directory('.', filename)

if __name__ == '__main__':
    app.run(debug=True)
