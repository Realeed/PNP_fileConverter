from flask import Flask, render_template, request, send_file
import os
from pathlib import Path
from zipfile import ZipFile
from converter import convertFile

app = Flask(__name__)
app.secret_key = "secret"

@app.route("/", methods = ["POST", "GET"])
def main():
    return render_template('index.html')

@app.route("/upload_excel", methods = ["POST", "GET"])
def getExcel():
    dir = 'resources\\'
    if not os.path.exists(dir):
        os.mkdir(dir)
    fullPaths = []
    filenames = []
    extensions = []
    excels = request.files.getlist('excelInp')
    for index, excel in enumerate(excels):
        fullPaths.append(dir + excel.filename)
        filenames.append(Path(fullPaths[index]).stem)
        extensions.append(Path(fullPaths[index]).suffix)
        excel.save(fullPaths[index])
        convertFile(fullPaths[index])

    if len(excels)  > 1:
        with ZipFile(dir + 'Formatted_excels.zip', 'w') as zipObj:
            for index, filename in enumerate(filenames):
                zipObj.write(fullPaths[index], arcname = filename + extensions[index])
        return send_file(dir + 'Formatted_excels.zip')
    return send_file(fullPaths[0], as_attachment=True)

if __name__  == "__main__":
    app.run(host='0.0.0.0', port=80, debug=True)