import os

from flask import Flask, request
from flask import jsonify
from flask_restplus import Resource, Api
from flask_restplus import reqparse
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from datetime import date
import flask_excel as excel

############# config

project_dir = os.path.dirname(os.path.abspath(__file__))
database_file = "sqlite:///{}".format(os.path.join(project_dir, "currencyDatabase.db"))

app = Flask(__name__)
cors = CORS(app)
app.config["SQLALCHEMY_DATABASE_URI"] = database_file
api = Api(app)

db = SQLAlchemy(app)

############# models

class CurrencyData(db.Model):
    id = db.Column(db.Integer, primary_key=True, unique=True, nullable=False)
    eur = db.Column(db.Float, nullable=False)
    usd = db.Column(db.Float, nullable=False)
    jpy = db.Column(db.Float, nullable=False)
    gbp = db.Column(db.Float, nullable=False)
    dataDateTime = db.Column(db.DateTime, nullable=False)

    def __repr__(self):
        return "<Currency data: {} {} {} {} {} {} >".format(self.id, self.eur, self.usd, self.jpy, self.gbp, self.dataDateTime)

############# routing
@api.route('/currency')
class Currency(Resource):
    def get(self):
        try:
            lastData = CurrencyData.query.order_by(-CurrencyData.id).first()
            #print("Ostatnia aktualizacja danych: ", lastData)
        except Exception as e:
            #print("Failed to load all data from database: ", e)
            return 500

        return jsonify({
            'eur': lastData.eur,
            'usd': lastData.usd,
            'jpy': lastData.jpy,
            'gbp': lastData.gbp,
            'date': lastData.dataDateTime
        })

    @api.param('eur')
    @api.param('usd')
    @api.param('jpy')
    @api.param('gbp')
    def post(self):
        dataFromJson=request.get_json()
        #print("Dane z JSONA: ", dataFromJson)

        if dataFromJson:
            try:
                c = CurrencyData(eur=dataFromJson['eur'], usd=dataFromJson['usd'], jpy=dataFromJson['jpy'], gbp=dataFromJson['gbp'], dataDateTime=date.today())
                #print("Pobrane dane: ", c)
                db.session.add(c)
                db.session.commit()
            
                return 200 #pomyślna operacja
            except Exception as e:
                print("Failed to add new data: ", e)
                return 500 #błąd serwera
        return 400 #jakiś błąd użytkownika

@api.route('/currencyAll')
class Currency(Resource):
    def get(self):
        try:
            cData = CurrencyData.query.all()
        except Exception as e:
            print("Failed to load all data from database: ", e)
            return 500
        all_data = [{
            'eur': cur.eur,
            'usd': cur.usd,
            'jpy': cur.jpy,
            'gbp': cur.gbp,
            'date': cur.dataDateTime 
        } for cur in cData]

        return jsonify(all_data)

@api.route('/saveToFile')
class Currency(Resource):
    def get(self):
        try:
            return excel.make_response_from_tables(db.session, [CurrencyData], "handsontable.html")
        except Exception as e:
            print("Failed to load all data from database: ", e)
            return 500
        return 200


if __name__ == '__main__':
    app.run(debug=True)