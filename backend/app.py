import os

from flask import Flask, request
from flask import jsonify
from flask_restplus import Resource, Api
from flask_restplus import reqparse
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
import datetime
import xlsxwriter
import pandas

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
                c = CurrencyData(eur=dataFromJson['eur'], usd=dataFromJson['usd'], jpy=dataFromJson['jpy'], gbp=dataFromJson['gbp'], dataDateTime=datetime.datetime.now())
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
            workbook = xlsxwriter.Workbook('current_exchange_rates.xlsx')
            worksheet = workbook.add_worksheet('Exchange Rates')

            cData = CurrencyData.query.all()
            all_data = [{
                'eur': cur.eur,
                'usd': cur.usd,
                'jpy': cur.jpy,
                'gbp': cur.gbp,
                'date': cur.dataDateTime 
            } for cur in cData]

            row = 0
            col = 0

            worksheet.write(row, col, "Aktuelle Wechselkurse: Übersicht \nCours de change actuels: aperçu \nCurrent exchange rates: overview \nTassi di cambio attuali: panoramica")
            worksheet.write(row, col + 5, str(datetime.datetime.now()))
            row += 1

            worksheet.write(row, col, "EUR")
            worksheet.write(row, col + 1, "1 EUR in/en CHF")
            row += 1
            worksheet.write(row, col, "USD")
            worksheet.write(row, col + 1, "1 USD in/en CHF")
            row += 1
            worksheet.write(row, col, "JPY")
            worksheet.write(row, col + 1, "100 JPY in/en CHF")
            row += 1
            worksheet.write(row, col, "GBP")
            worksheet.write(row, col + 1, "1 GBP in/en CHF")
            row += 1

            worksheet.write(row, col, "")
            worksheet.write(row, col + 1, "EUR")
            worksheet.write(row, col + 2, "USD")
            worksheet.write(row, col + 3, "JPY")
            worksheet.write(row, col + 4, "GBP")
            row += 1

            for d in all_data:
                print("data: ", d["date"])
                worksheet.write(row, col, str(d['date']))
                worksheet.write(row, col + 1, d['eur'])
                worksheet.write(row, col + 2, d['usd'])
                worksheet.write(row, col + 3, d['jpy'])
                worksheet.write(row, col + 4, d['gbp'])

                row += 1

            workbook.close()

        except Exception as e:
            print("Failed to save all data from database to file: ", e)
            return 500
        return 200


@api.route('/readFromFile')
class Currency(Resource):
    def get(self):
        try:
            excel_data_df = pandas.read_excel('current_exchange_rates.xlsx', sheet_name='Exchange Rates')
            # print("Dane z pliku: \n", excel_data_df)

            date_column=excel_data_df['Aktuelle Wechselkurse: Übersicht\nCours de change actuels: aperçu\nCurrent exchange rates: overview\nTassi di cambio attuali: panoramica'].tolist()
            eur_column=excel_data_df['Unnamed: 1'].tolist()
            usd_column=excel_data_df['Unnamed: 2'].tolist()
            jpy_column=excel_data_df['Unnamed: 3'].tolist()
            gbp_column=excel_data_df['Unnamed: 4'].tolist()

            listLen=len(date_column)
            
            column=5
            for i in range(listLen-column):
                # print("i: ", i+column, " | ", date_column[i+column], " | ", eur_column[i+column], " | ", usd_column[i+column], " | ", jpy_column[i+column], " | ", gbp_column[i+column])
                c = CurrencyData(eur=eur_column[i+column], usd=usd_column[i+column], jpy=jpy_column[i+column], gbp=gbp_column[i+column], dataDateTime=pandas.to_datetime(date_column[i+column]))

                db.session.add(c)
            db.session.commit()
            
        except Exception as e:
            print("Failed to read all data from file to database: ", e)
            return 500
        return 200

if __name__ == '__main__':
    app.run(debug=True)