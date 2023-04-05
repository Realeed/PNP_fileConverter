from flask import Flask, render_template, request, send_file
from pathlib import Path
from zipfile import ZipFile
from converter import convertSingle10, convertPanel10

app = Flask(__name__)
app.secret_key = "secret"

@app.route("/", methods = ["POST", "GET"])
def main():
    return render_template('index.html')

@app.route("/upload_excel", methods = ["POST", "GET"])
def getExcel():
    dir = 'resources\\'
    fileType = request.form.get('filetype')
    
    excel = request.files.get('excelInp')

    fullPath = dir + excel.filename
    filename = Path(fullPath).stem
    extension = Path(fullPath).suffix
    excel.save(fullPath)
    if fileType == 'single10':
        layers = convertSingle10(fullPath)
        if layers == 2:
            with ZipFile(dir + filename + '.zip', 'w') as zipObj:
                zipObj.write(dir + filename.replace('Pick Place for ', '').replace('_N10', '').replace('Panel', 'Single') + '_Top_N10' + extension, 
                            arcname = filename.replace('Pick Place for ', '').replace('_N10', '').replace('Panel', 'Single') + '_Top_N10' + extension)
                zipObj.write(dir + filename.replace('Pick Place for ', '').replace('_N10', '').replace('Panel', 'Single') + '_Bottom_N10' + extension, 
                            arcname = filename.replace('Pick Place for ', '').replace('_N10', '').replace('Panel', 'Single') + '_Bottom_N10' + extension)
            return send_file(dir + filename + '.zip')
        elif layers == 1:
            return send_file(dir + filename.replace('Pick Place for ', '').replace('_N10', '').replace('Panel', 'Single') + '_Top_N10' + extension)
    elif fileType == 'panel10':
        layers = convertPanel10(fullPath)
        if layers == 2:
            with ZipFile(dir + filename + '.zip', 'w') as zipObj:
                zipObj.write(dir + filename.replace('Pick Place for ', '').replace('_N10', '') + '_Top_N10' + extension, 
                            arcname = filename.replace('Pick Place for ', '').replace('_N10', '') + '_Top_N10' + extension)
                zipObj.write(dir + filename.replace('Pick Place for ', '').replace('_N10', '') + '_Bottom_N10' + extension, 
                            arcname = filename.replace('Pick Place for ', '').replace('_N10', '') + '_Bottom_N10' + extension)
            return send_file(dir + filename + '.zip')
        elif layers == 1:
            return send_file(dir + filename.replace('Pick Place for ', '').replace('_N10', '') + '_Top_N10' + extension)

if __name__  == "__main__":
    app.run(host='0.0.0.0', port=80)