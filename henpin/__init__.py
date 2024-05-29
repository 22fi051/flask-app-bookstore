from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config.from_object('henpin.config')

db = SQLAlchemy(app)

class excelfile(db.Model):
    #__tablename__ = 'excelfile'
    id = db.Column(db.Integer, primary_key=True)
    file_name = db.Column(db.String(255), nullable=False)

import henpin.views