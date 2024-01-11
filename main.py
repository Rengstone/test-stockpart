import psycopg2
import openpyxl
import os

from flask import Flask, render_template, redirect
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from config import config
from werkzeug.utils import secure_filename

class UploadFileForm(FlaskForm):
    file = FileField("File")
    submit = SubmitField("UploadFile")

app = Flask(__name__)
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = 'file'

@app.route('/', methods = ['GET', 'POST'])
@app.route('/home', methods = ['GET', 'POST'])

def home():
    form = UploadFileForm()
    if form.validate_on_submit():
        file = form.file.data
        filename = secure_filename(file.filename)
        file.save(os.path.join(os.path.abspath(os.path.dirname(__file__)), app.config['UPLOAD_FOLDER'], secure_filename(file.filename)))
        return connect(filename)
    return render_template('index.html', form = form)

def connect(filename):
    connection = None
    try:
        params = config()
        print('Connecting to the postgreSQL database ...')
        connection = psycopg2.connect(**params)
        crsr = connection.cursor()

        work_book = openpyxl.load_workbook('./file/' + filename)
        work_sheet = work_book['Jalin']
        
        numbers = work_sheet['A']
        number = [numbers[x].value for x in range(len(numbers))]
        brands = work_sheet['B']
        brand = [brands[x].value for x in range(len(brands))]
        part_numbers = work_sheet['C']
        part_number = [part_numbers[x].value for x in range(len(part_numbers))]
        nama_parts = work_sheet['D']
        nama_part = [nama_parts[x].value for x in range(len(nama_parts))]
        serial_numbers = work_sheet['E']
        serial_number = [serial_numbers[x].value for x in range(len(serial_numbers))]
        barcodes = work_sheet['F']
        barcode = [barcodes[x].value for x in range(len(barcodes))]
        qtys = work_sheet['G']
        qty = [qtys[x].value for x in range(len(qtys))]
        ssi_areas = work_sheet['H']
        ssi_area = [ssi_areas[x].value for x in range(len(ssi_areas))]
        sp_codes = work_sheet['I']
        sp_code = [sp_codes[x].value for x in range(len(sp_codes))]
        sp_areas = work_sheet['J']
        sp_area = [sp_areas[x].value for x in range(len(sp_areas))]
        regions = work_sheet['K']
        region = [regions[x].value for x in range(len(regions))]


        for x in range (len(nama_part)):
            if x > 2:
                if nama_part[x] != None:
                    crsr.execute("INSERT INTO tb_stock VALUES({0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}')".format(number[x], brand[x], part_number[x], nama_part[x], serial_number[x], barcode[x], qty[x], ssi_area[x], sp_code[x], sp_area[x], region[x], filename))
        connection.commit()
        crsr.close()

    except(Exception, psycopg2.DatabaseError) as error:
        print(error)

    finally:
        if connection is not None:
            connection.close()
            print('Database connection terminated.')
        return redirect('/list')

@app.route('/list')
def list():

    connection = None
    try:
        params = config()
        print('Connecting to the postgreSQL database ...')
        connection = psycopg2.connect(**params)
        crsr = connection.cursor()
        crsr.execute("SELECT * FROM tb_stock")
        result = crsr.fetchall()
        connection.commit()
        crsr.close()

    except(Exception, psycopg2.DatabaseError) as error:
        print(error)

    finally:
        if connection is not None:
            connection.close()
            print('Database connection terminated.')
        return render_template('list.html', result = result)

if __name__ == "__main__":
    app.run(debug=True)