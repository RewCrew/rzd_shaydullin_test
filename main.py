import os
import json
import pandas
from flask import Flask, render_template
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from werkzeug.utils import secure_filename
from wtforms.validators import InputRequired

app = Flask(__name__)
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = os.getcwd()+'/storage/'

class UploadFileForm(FlaskForm):
    file = FileField("File", validators=[InputRequired()])
    submit = SubmitField("Upload File")

def to_json(file, file_name):
    data = {}
    file_name = file_name.split('.')[0]

    sheet_names = pandas.ExcelFile(file).sheet_names
    for sheet in sheet_names:
            excel_data = pandas.read_excel(file, sheet_name=sheet)
            json_str = excel_data.to_json(orient = 'records', date_format = 'iso')
            data[sheet] = json.loads(json_str)
    if os.path.exists(os.getcwd()+'\\json_files\\'+file_name+'.json'):
        return 'File already exists'
    else:
        with open(os.getcwd()+'\\json_files\\'+file_name+'.json', 'w', encoding='utf-8') as jn:
            jn.write(json.dumps(data, ensure_ascii=False, indent=4))
            jn.close()
            return '  file uploaded to server'

@app.route('/upload/excel/', methods=['GET',"POST"])

def parser():
    form = UploadFileForm()
    if form.validate_on_submit():
        file = form.file.data
        path = os.path.join(os.getcwd()+'/storage/' + secure_filename(file.filename))
        file.save(os.path.join(os.getcwd()+'/storage/',secure_filename(file.filename)))
        result = to_json(path, file.filename)
        os.remove(path)
        return "File has been uploaded." + ' '+result
    return render_template('index.html', form=form)

if __name__ == '__main__':
    app.run(debug=True)