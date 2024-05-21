import os
from pymongo import MongoClient
from dotenv import load_dotenv
from flask import Flask, request, jsonify
import win32com.client as win32
import pythoncom

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
        return jsonify({"Erro": str(e)}), 500

def send_email(value, payer, payee):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        email_out = outlook.CreateItem(0)
        email_out.To = payer['email'] + '; ' + payee['email']
        email_out.Subject = "Transferência realizada!"
        email_out.HTMLBody = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Detalhes da Transação</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }}
        .container {{
            border-radius: 12px;
            box-shadow: 0 10px 20px rgba(0, 0, 0, 0.1);
            padding: 30px;
            max-width: 500px;
            width: 100%;
            text-align: center;
        }}
        .header {{
            font-size: 1.5em;
            margin-bottom: 20px;
            color: #2c3e50;
        }}
        .value {{
            font-size: 2em;
            font-weight: bold;
            color: #27ae60;
            margin-bottom: 20px;
        }}
        .info {{
            font-size: 1.2em;
            color: #34495e;
            margin-bottom: 10px;
        }}
        .info span {{
            display: block;
            font-weight: bold;
            color: #2980b9;
        }}
        .footer {{
            margin-top: 20px;
            font-size: 0.9em;
            color: #95a5a6;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">Detalhes da Transação</div>
        <div class="value">Valor: {value}</div>
        <div class="info">
            <span>Nome Pagador:</span> {payer['nome_completo']}
        </div>
        <div class="info">
            <span>Nome Recebedor:</span> {payee['nome_completo']}
        </div>
        <div class="footer">Obrigado por utilizar nossos serviços.</div>
    </div>
</body>
</html>
"""

        email_out.Send()
    except Exception as e:
        print(f"Falha ao enviar email: {str(e)}")

def validation_antifraud(payee, payer):
    return True

if __name__ == "__main__":
    app.run(debug=True, port=5005, host="0.0.0.0")