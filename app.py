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

class User:
    def __init__(self, id, nome_completo, cpf, email, saldo, senha, tipo):
        self.id = id
        self.nome_completo = nome_completo
        self.cpf = cpf
        self.email = email
        self.saldo = saldo
        self.senha = senha
        self.tipo = tipo

    def __repr__(self):
        return (f"User(id={self.id}, nome_completo='{self.nome_completo}', cpf='{self.cpf}', "
                f"email='{self.email}', saldo={self.saldo}, senha='{self.senha}', tipo='{self.tipo}')")

@app.route("/dbpicpay", methods=["GET"])
def read():
    try:
        output = list(collection.find({}, {'_id': 0}))
        return jsonify(output), 200
    except Exception as e:
        return jsonify({"Erro": str(e)}), 500



@app.route("/dbpicpay", methods=["POST"])
def create_user():
    try:
        data = request.json
        user_obj = user_validation(data)

        last_id = collection.find_one(sort=[('id', -1)])
        id_correto = last_id['id'] + 1 if last_id else 1

        if user_obj.id != id_correto:
            return jsonify({"Erro": f"ID Inválido! Esperado: {id_correto}"}), 400

        if collection.find_one({'$or': [{'cpf': user_obj.cpf}, {'email': user_obj.email}, {'id': user_obj.id}]}):
            return jsonify({"Erro": "Cliente Existente!"}), 400

        user = {
            'id': user_obj.id,
            'nome_completo': user_obj.nome_completo,
            'cpf': user_obj.cpf,
            'email': user_obj.email,
            'saldo': user_obj.saldo,
            'senha': user_obj.senha,
            'tipo': user_obj.tipo
        }

        collection.insert_one(user)
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

    if not validation_antifraud(id_payee, id_payer):
        return jsonify({"Erro": "Fraude detectada!"}), 403

    payer = collection.find_one({'id': id_payer})
    payee = collection.find_one({'id': id_payee})

    if not payer or not payee or payer['tipo'] == 'lojista':
        return jsonify({"Erro": "Transação inválida!"}), 403

    if payer['saldo'] < value:
        return jsonify({"Erro": "Saldo Insuficiente!"}), 400

    saldo_payer = payer['saldo']
    saldo_payee = payee['saldo']

    try:
        collection.update_one({'id': id_payer}, {'$inc': {'saldo': -value}})
        collection.update_one({'id': id_payee}, {'$inc': {'saldo': value}})
        send_email(value, payer, payee)
        return jsonify({'Status': 'Transferência Concluída!'}), 200
    except Exception as e:
        collection.update_one({'id': id_payer}, {'$set': {'saldo': saldo_payer}})
        collection.update_one({'id': id_payee}, {'$set': {'saldo': saldo_payee}})
        return jsonify({"Erro": f"Falha na transferência!"}), 500


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

def validation_antifraud(payee_id, payer_id):
    return True

def user_validation(data):
    campos_necessario = ['id', 'nome_completo', 'cpf', 'email', 'saldo', 'senha', 'tipo']

    if not all(campos in data for campos in campos_necessario):
        raise ValueError(f"está faltando campos.")

    if not isinstance(data['id'], int):
        raise ValueError("O campo 'id' deve ser um inteiro.")
    if not isinstance(data['nome_completo'], str):
        raise ValueError("O campo 'nome_completo' deve ser uma string.")
    if not isinstance(data['cpf'], str):
        raise ValueError("O campo 'cpf' deve ser uma string.")
    if not isinstance(data['email'], str):
        raise ValueError("O campo 'email' deve ser uma string.")
    if not isinstance(data['saldo'], (int, float)):
        raise ValueError("O campo 'saldo' deve ser um número.")
    if not isinstance(data['senha'], str):
        raise ValueError("O campo 'senha' deve ser uma string.")
    if not isinstance(data['tipo'], str):
        raise ValueError("O campo 'tipo' deve ser uma string.")

    required_keys = ['id', 'nome_completo', 'cpf', 'email', 'saldo', 'senha', 'tipo']

    if not all(key in data for key in required_keys):
        raise ValueError("falta campos necessários.")

    return User(
        id=data['id'],
        nome_completo=data['nome_completo'],
        cpf=data['cpf'],
        email=data['email'],
        saldo=data['saldo'],
        senha=data['senha'],
        tipo=data['tipo']
    )



if __name__ == "__main__":
    app.run(debug=True, port=5005, host="0.0.0.0")




