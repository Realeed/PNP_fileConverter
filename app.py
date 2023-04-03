from flask import Flask, render_template, request, send_file
import os
from pathlib import Path
from zipfile import ZipFile
from converter import convertSingle10

app = Flask(__name__)
app.secret_key = "secret"

@app.route("/", methods = ["POST", "GET"])
def main():
    return render_template('index.html')

@app.route("/upload_excel", methods = ["POST", "GET"])
def getExcel():
    dir = 'resources\\'
    fullPaths = []
    filenames = []
    extensions = []
    fileType = request.form.get('filetype')
    
    excels = request.files.getlist('excelInp')
    if fileType == 'single10':
        for index, excel in enumerate(excels):
            fullPaths.append(dir + excel.filename)
            filenames.append(Path(fullPaths[index]).stem)
            extensions.append(Path(fullPaths[index]).suffix)
            excel.save(fullPaths[index])
            layers = convertSingle10(fullPaths[index])

            if layers == 2:
                with ZipFile(dir + filenames[0] + '.zip', 'w') as zipObj:
                    print(fullPaths[0])
                    zipObj.write(dir + filenames[0].replace('Pick Place for ', '').replace('Panel', 'Single') + '_Top_N10' + extensions[0], 
                                arcname = filenames[0].replace('Pick Place for ', '').replace('Panel', 'Single') + '_Top_N10' + extensions[0])
                    zipObj.write(dir + filenames[0].replace('Pick Place for ', '').replace('Panel', 'Single') + '_Bottom_N10' + extensions[0], 
                                arcname = filenames[0].replace('Pick Place for ', '').replace('Panel', 'Single') + '_Bottom_N10' + extensions[0])
                return send_file(dir + filenames[0] + '.zip')
            # if len(excels) > 1:
            #     with ZipFile(dir + 'Formatted_excels.zip', 'w') as zipObj:
            #         for index, filename in enumerate(filenames):
            #             zipObj.write(fullPaths[index], arcname = filename + extensions[index])
            #     return send_file(dir + 'Formatted_excels.zip')
                
           # return send_file(fullPaths[0], as_attachment=True)

if __name__  == "__main__":
    app.run(host='0.0.0.0', port=80, debug=True)