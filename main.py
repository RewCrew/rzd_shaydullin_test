import os
import json
import pandas
from flask import Flask, render_template, jsonify
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from werkzeug.utils import secure_filename
from wtforms.validators import InputRequired
from threading import Thread
import time

app = Flask(__name__)
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = os.getcwd()+'/storage/'

class UploadFileForm(FlaskForm):
    file = FileField("File", validators=[InputRequired()])
    submit = SubmitField("Upload File")

def to_json(path, file):
    time.sleep(2)
    data = {}
    file_name = file.filename.split('.')[0]
    file.save(os.path.join(os.getcwd() + '/storage/', secure_filename(file.filename)))

    sheet_names = pandas.ExcelFile(path).sheet_names
    for sheet in sheet_names:
            excel_data = pandas.read_excel(path, sheet_name=sheet)
            json_str = excel_data.to_json(orient = 'records', date_format = 'iso')
            data[sheet] = json.loads(json_str)
    # if os.path.exists(os.getcwd()+'\\json_files\\'+file_name+'.json'):
    #     return 'File already exists'
    # else:
    with open(os.getcwd()+'\\json_files\\'+file_name+'.json', 'w', encoding='utf-8') as jn:
        jn.write(json.dumps(data, ensure_ascii=False, indent=4))
        # jn.close()
    os.remove(path)
        # можно убрать клоуз

@app.route('/upload/excel/', methods=['GET',"POST"])

def parser():
    form = UploadFileForm()
    if form.validate_on_submit():
        file = form.file.data
        path = os.path.join(os.getcwd()+'/storage/' + secure_filename(file.filename))
        if os.path.exists(os.getcwd() + '\\json_files\\' + file.filename.split('.')[0] + '.json'):
            return jsonify('File already exist')
        # file.save(os.path.join(os.getcwd()+'/storage/',secure_filename(file.filename)))
        Thread(target=to_json(path, file)).start()
        # to_json(path, file.filename)
        # os.remove(path)
        return jsonify("File has been uploaded.")
    return render_template('index.html', form=form)

if __name__ == '__main__':
    app.run(debug=True)