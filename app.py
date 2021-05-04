import os
from werkzeug.utils import secure_filename
from flask import Flask,flash,request,redirect,send_file,render_template, url_for
import pandas as pd
from werkzeug.wrappers import BaseRequest
from werkzeug.wsgi import responder
from werkzeug.exceptions import HTTPException, NotFound
import sqlalchemy
import xlsxwriter
import io
from io import StringIO

app = Flask(__name__)

app.secret_key = "secret key"


# Opens in the web page
@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

# to view the uploaded excel sheet
@app.route('/data', methods=['GET', 'POST'])
def data():
    if request.method == 'POST':
        file = request.form['file']
        #print('The uploaded file name is: ' + file)
        #file = "./Test/" + file 
        #print('Complete path of the file after updating the path is: ' + file)
        data = pd.read_excel(file)
        return render_template('data.html', data=data.to_html())

# to save the excel sheet to database and create table

@app.route('/save', methods=['GET', 'POST'])
def save():
    if request.method == 'POST':
        file = request.form['file']
        df = pd.read_excel(file)
        engine = sqlalchemy.create_engine("postgresql://postgres:admin@localhost/merge")
        con = engine.connect()
        #print(engine.table_names())
        table_name = 'mergetable'
        df.to_sql(table_name, con, if_exists='replace')

        #print("The tables list is:")
        #print(engine.table_names())
        con.close()

    return render_template('data.html', data=df.to_html())


@app.route('/master', methods=['GET', 'POST'])
def master():
    if request.method == 'POST':
        file = request.form['file']
        
        df1 = pd.read_excel(file)
        engine = sqlalchemy.create_engine("postgresql://postgres:admin@localhost/merge")
        con = engine.connect()
        #print(engine.table_names())
        table_name = 'mergetable'
        df1.to_sql(table_name, con, if_exists='append')
        result = pd.read_sql("select * from \"mergetable\"", con)
        con.close()
    return render_template('merge.html', data=result.to_html())


@app.route('/download_excel/')

def download_excel():
    engine = sqlalchemy.create_engine("postgresql://postgres:admin@localhost/test")
    con = engine.connect()
    result = pd.read_sql("select * from \"sleepdata\"", con)
    strIO = io.BytesIO()
    excel_writer = pd.ExcelWriter(strIO, engine="xlsxwriter")
    result.to_excel(excel_writer, sheet_name="sheet1")
    excel_writer.save()
    excel_data = strIO.getvalue()
    strIO.seek(0)
    con.close()

    return send_file(strIO,
                     attachment_filename='master.xlsx',
                     as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)


