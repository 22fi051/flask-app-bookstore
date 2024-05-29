from flask import render_template,request,redirect,url_for
from henpin import app
from henpin.excel import excel
from datetime import datetime, timedelta

from henpin import db
from henpin import excelfile

date1 = datetime.now()
date2 = datetime.now() + timedelta(days=1)
datenow = str(date1.month) + '/' + str(date1.day)
datenext = str(date2.month) + '/' + str(date2.day)

@app.route('/')
def index():
    return render_template('henpin/index.html', insert_datenow=datenow, insert_datenext=datenext, sub="登録")

@app.route('/henpinform',methods=['POST'])
def henpinform():
    no1 = request.form['no1']
    no2 = request.form['no2']
    no3 = request.form['no3']
    no4 = request.form['no4']
    no5 = request.form['no5']
    numberList = [no1,no2,no3,no4,no5]

    box1 = int(request.form['box1'])
    box2 = int(request.form['box2'])
    box3 = int(request.form['box3'])
    box4 = int(request.form['box4'])
    box5 = int(request.form['box5'])
    boxList = [box1,box2,box3,box4,box5]

    date = request.form['date']

    datenext = request.form['sdate']

    name = request.form['name']

    e = excel(numberList,boxList,date,datenext,name)
    e.enter_excel()
    return render_template('henpin/index.html', insert_datenow=datenow, insert_datenext=datenext, sub="保存しました")

@app.route('/excel')
def file():
    return render_template('henpin/excel.html')

@app.route('/file_select', methods=['POST'])
def file_select():
    Excelfile = excelfile.query.get(10)
    Excelfile.file_name = request.form.get('filename')
    
    db.session.merge(Excelfile)
    db.session.commit()
    return redirect(url_for('index'))
