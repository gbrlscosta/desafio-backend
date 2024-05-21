import os
from pymongo import MongoClient
from dotenv import load_dotenv
from flask import Flask, request, jsonify
import win32com.client as win32
import pythoncom
from datetime import datetime

load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")
client = MongoClient(MONGODB_URI)
db = client['dbpicpay']
collection = db['clients']

app = Flask(__name__)

@app.route("/dbpicpay", methods=["GET"])
def read():
    try:
        output = [
            data for data in collection.find({}, {'_id': 0})
        ]
        return jsonify(output), 200
    except Exception as e:
        return jsonify({"Erro": str(e)}), 500

@app.route("/dbpicpay", methods=["POST"])
def create_user():
    data = request.json
    cpf = data.get('cpf')
    email = data.get('email')
    user_id = data.get('id')

    for i in collection.find().sort({'id': -1}).limit(1):
        if user_id != int(i['id']) + 1:
            return jsonify({"Erro": f"ID Inválido!"}), 400

    campos_necessarios = ['id', 'nome_completo', 'cpf', 'email', 'saldo', 'senha', 'tipo']

    for campo in campos_necessarios:
        if campo not in data:
            return jsonify({"Erro": f"Campo {campo} é obrigatório!"}), 400

    try:
        if collection.find_one({'$or': [{'cpf': cpf}, {'email': email}, {'id': user_id}]}):
            return jsonify({"Erro": "Cliente Existente!"}), 400

        collection.insert_one(data)
        return jsonify({"Status": "Cliente Inserido!"}), 200
    except Exception as e:
        return jsonify({"Erro": str(e)}), 500

@app.route("/dbpicpay/transfer", methods=["PATCH"])
def transfer():
    data = request.json
    value = data.get('value')
    id_payer = data.get('payer')
    id_payee = data.get('payee')

    if not value or not id_payer or not id_payee:
        return jsonify({"Erro": "Valor, pagador e recebedor são obrigatórios!"}), 400

    if validation_antifraud == True:

        if value <= 0:
            return jsonify({"Erro": "O valor da transferência deve ser maior que zero!"}), 400

        try:
            payer = collection.find_one({'id': id_payer})
            payee = collection.find_one({'id': id_payee})

            if payer is None or payee is None:
                return jsonify({"Erro": "Pagador ou recebedor não encontrados"}), 404

            if payer['tipo'] == 'lojista':
                return jsonify({'Erro': 'O Usuário não pode fazer transferência!'}), 403

            saldo_payer = float(payer['saldo'])
            saldo_payee = float(payee['saldo'])

            if saldo_payer < value:
                return jsonify({"Erro": "Saldo Insuficiente!"}), 400

            new_saldo_payer = saldo_payer - value
            new_saldo_payee = saldo_payee + value

            collection.update_one({'id': id_payer}, {'$set': {'saldo': new_saldo_payer}})
            collection.update_one({'id': id_payee}, {'$set': {'saldo': new_saldo_payee}})

            send_email(value, payer, payee)

            return jsonify({'Status': 'Transferência Concluída!'}), 200
        except Exception as e:
            collection.update_one({'id': id_payer}, {'$set': {'saldo': saldo_payer}})
            collection.update_one({'id': id_payee}, {'$set': {'saldo': saldo_payee}})
            return jsonify({"Erro": str(e)}), 500
    else:
        return jsonify({"Erro": "Fraude!"}), 403

def send_email(value, payer, payee):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        email_out = outlook.CreateItem(0)
        email_out.To = payer['email'] + '; ' + payee['email']
        email_out.Subject = "Transferência realizada!"
        data = datetime.now().strftime("%d/%m/%Y")
        email_out.HTMLBody = f"""
        <center><h2> Detalhes da Transferência </h2></center>
        <center><h2 style="color:#00A000"><strong>R${value}</strong></h2></center>
        <hr>
        <p>Data: <strong>{data}</strong>
        <p>Nome Pagador: <strong>{payer['nome_completo']}</strong></p>
        <p>Nome Recebedor: <strong>{payee['nome_completo']}</strong></p>
        """
        email_out.Send()
    except Exception as e:
        print(f"Falha ao enviar email: {str(e)}")

def validation_antifraud(payee, payer):
    return True

if __name__ == "__main__":
    app.run(debug=True, port=5005, host="0.0.0.0")