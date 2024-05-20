import os
from pymongo import MongoClient
from dotenv import load_dotenv
from flask import Flask, request, jsonify, Response
import win32com.client as win32
import pythoncom
# Carregar variáveis de ambiente
load_dotenv()

# Configurar MongoDB
MONGODB_URI = os.getenv("MONGODB_URI")
client = MongoClient(MONGODB_URI)
db = client['dbpicpay']
collection = db['clients']

app = Flask(__name__)

@app.route("/dbpicpay", methods=["GET"])
def read():
    try:
        output = [
            {item: data[item] for item in data if item != "_id"} for data in collection.find()
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

    if not cpf or not email or not user_id:
        return jsonify({"Erro": "CPF, email e ID são obrigatórios!"}), 400

    try:
        if collection.find_one({'$or': [{'cpf': cpf}, {'email': email}, {'id': user_id}]}):
            return jsonify({"Erro": "Cliente Existente!"}), 400

        collection.insert_one(data)
        return jsonify({"Status": "Cliente Inserido!"}), 201
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

    try:
        payer = collection.find_one({'id': id_payer})
        payee = collection.find_one({'id': id_payee})

        if payer is None or payee is None:
            return jsonify({"Erro": "Pagador ou recebedor não encontrados"}), 404

        if payer['tipo'] == 'lojista':
            return jsonify({'Erro': 'O Usuário não pode fazer transferencia!'}), 403

        saldo_payer = float(payer['saldo'])
        if saldo_payer < value:
            return jsonify({"Erro": "Saldo Insuficiente!"}), 400

        new_saldo_payer = saldo_payer - value
        collection.update_one({'id': id_payer}, {'$set': {'saldo': new_saldo_payer}})

        saldo_payee = float(payee['saldo'])
        new_saldo_payee = saldo_payee + value
        collection.update_one({'id': id_payee}, {'$set': {'saldo': new_saldo_payee}})

        try:

            pythoncom.CoInitialize()
            outlook = win32.Dispatch('outlook.application')
            email_out = outlook.CreateItem(0)
            email = payee['email']
            email_out.To = email
            email_out.Subject = "Transferencia recebida!"
            email_out.HTMLBody = f"""
            <p>Valor: {value}</p>
            <p>Nome Pagador: {payer['nome_completo']}</p>
            <p>Nome Recebedor: {payee['nome_completo']}</p>
            """
            email_out.Send()
        except Exception as e:
            return jsonify({"Erro": f"Falha ao enviar email: {str(e)}"}), 500

        return jsonify({'Status': 'Transferencia Concluida!'}), 200
    except Exception as e:
        return jsonify({"Erro": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, port=5005, host="0.0.0.0")
