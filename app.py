from flask import Flask, request, jsonify
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import tempfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Authenticate and create PyDrive client
gauth = GoogleAuth()
# Creates local webserver and auto handles authentication.
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)


@app.route('/')
def index():
    return app.send_static_file('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    filename = file.filename
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Upload file to Google Drive
    try:
        with open(file_path, "rb") as f:
            gdrive_file = drive.CreateFile({'title': filename})
            gdrive_file.SetContentFile(file_path)
            gdrive_file.Upload()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({'filename': filename}), 200


if __name__ == '__main__':
    app.run(debug=True)
