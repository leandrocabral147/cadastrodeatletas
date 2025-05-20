from flask import Flask, render_template, request, redirect, url_for, flash, session
import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests  # Adicione esta linha no início do arquivo para fazer requisições HTTP

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'  # Necessário para usar flash messages e sessões

# Configurações de e-mail
EMAIL_HOST = 'smtp.gmail.com'  # Servidor SMTP do Gmail
EMAIL_PORT = 587  # Porta SMTP do Gmail
EMAIL_USER = 'leandrocabral147@gmail.com'  # Seu e-mail
EMAIL_PASSWORD = 'wryhliczhhiprntw'  # Sua senha ou senha de aplicativo

# Caminho para o arquivo Excel de atletas
EXCEL_FILE_ATLETAS = 'atletas.xlsx'

# Caminho para o arquivo Excel de usuários
EXCEL_FILE_USUARIOS = 'usuarios.xlsx'

# Verifica se o arquivo Excel de atletas já existe, se não, cria um novo
if not os.path.exists(EXCEL_FILE_ATLETAS):
    df = pd.DataFrame(columns=[
        'Nome', 'Idade', 'Sexo', 'Categoria', 'Telefone', 'Email',
        'Cidade', 'CEP', 'Logradouro', 'Data de Nascimento'
    ])
    df.to_excel(EXCEL_FILE_ATLETAS, index=False)

# Verifica se o arquivo Excel de usuários já existe, se não, cria um novo
if not os.path.exists(EXCEL_FILE_USUARIOS):
    df = pd.DataFrame(columns=['Email', 'Senha'])
    df.to_excel(EXCEL_FILE_USUARIOS, index=False)

def consultar_cep(cep):
    """
    Função para consultar o CEP na API do ViaCEP.
    Retorna um dicionário com os dados do endereço ou None se o CEP for inválido.
    """
    url = f"https://viacep.com.br/ws/{cep}/json/"
    response = requests.get(url)

    if response.status_code == 200:  # Verifica se a requisição foi bem-sucedida
        dados = response.json()
        if 'erro' not in dados:  # Verifica se o CEP é válido
            return dados
    return None

def enviar_email(destinatario, assunto, corpo):
    """
    Função para enviar e-mails via SMTP.
    """
    try:
        # Configuração do e-mail
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = destinatario
        msg['Subject'] = assunto

        # Corpo do e-mail
        msg.attach(MIMEText(corpo, 'plain'))

        # Conexão com o servidor SMTP
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()  # Habilita a criptografia TLS
        server.login(EMAIL_USER, EMAIL_PASSWORD)  # Autenticação
        server.sendmail(EMAIL_USER, destinatario, msg.as_string())  # Envio do e-mail
        server.quit()  # Encerra a conexão

        return True
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        return False

@app.route('/')
def index():
    if 'email' in session:
        return render_template('index.html')
    else:
        return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        senha = request.form['senha']

        # Lendo o arquivo Excel de usuários
        df = pd.read_excel(EXCEL_FILE_USUARIOS)

        # Verificando se o usuário existe
        usuario = df[(df['Email'] == email) & (df['Senha'] == senha)]

        if not usuario.empty:
            session['email'] = email
            flash('Login realizado com sucesso!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Email ou senha incorretos. Tente novamente.', 'error')
            return redirect(url_for('login'))

    return render_template('login.html')

@app.route('/cadastro_usuario', methods=['GET', 'POST'])
def cadastro_usuario():
    if request.method == 'POST':
        email = request.form['email']
        senha = request.form['senha']

        # Lendo o arquivo Excel de usuários
        df = pd.read_excel(EXCEL_FILE_USUARIOS)

        # Verificando se o email já está cadastrado
        if email in df['Email'].values:
            flash('Email já cadastrado. Tente outro email.', 'error')
            return redirect(url_for('cadastro_usuario'))

        # Adicionando novo usuário
        novo_usuario = pd.DataFrame([{
            'Email': email,
            'Senha': senha
        }])

        # Concatenando o novo usuário ao DataFrame existente
        df = pd.concat([df, novo_usuario], ignore_index=True)

        # Salvando no arquivo Excel
        df.to_excel(EXCEL_FILE_USUARIOS, index=False)

        flash('Usuário cadastrado com sucesso! Faça login para continuar.', 'success')
        return redirect(url_for('login'))

    return render_template('cadastro_usuario.html')

@app.route('/logout')
def logout():
    session.pop('email', None)
    flash('Você foi deslogado com sucesso.', 'success')
    return redirect(url_for('login'))

@app.route('/cadastro', methods=['GET', 'POST'])
def cadastro():
    if 'email' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        # Coletando dados do formulário
        nome = request.form['nome']
        idade = request.form['idade']
        sexo = request.form['sexo']
        categoria = request.form['categoria']
        telefone = request.form['telefone']
        email = request.form['email']
        cidade = request.form['cidade']
        cep = request.form['cep']
        logradouro = request.form['logradouro']
        data_nascimento = request.form['data_nascimento']

        # Validando o CEP
        dados_cep = consultar_cep(cep)
        if not dados_cep:
            flash('CEP inválido. Por favor, verifique o CEP e tente novamente.', 'error')
            return redirect(url_for('cadastro'))

        # Preenchendo os campos de endereço com os dados da API
        logradouro = dados_cep.get('logradouro', '')
        cidade = dados_cep.get('localidade', '')
        estado = dados_cep.get('uf', '')

        # Convertendo a data de nascimento para o formato correto
        try:
            data_nascimento = datetime.strptime(data_nascimento, '%Y-%m-%d').date()
        except ValueError:
            flash('Formato de data inválido. Use o formato AAAA-MM-DD.', 'error')
            return redirect(url_for('cadastro'))

        # Lendo o arquivo Excel existente
        df = pd.read_excel(EXCEL_FILE_ATLETAS)

        # Adicionando novo atleta
        novo_atleta = pd.DataFrame([{
            'Nome': nome,
            'Idade': idade,
            'Sexo': sexo,
            'Categoria': categoria,
            'Telefone': telefone,
            'Email': email,
            'Cidade': cidade,
            'CEP': cep,
            'Logradouro': logradouro,
            'Data de Nascimento': data_nascimento
        }])

        # Concatenando o novo atleta ao DataFrame existente
        df = pd.concat([df, novo_atleta], ignore_index=True)

        # Salvando no arquivo Excel
        df.to_excel(EXCEL_FILE_ATLETAS, index=False)

        # Enviando e-mail de confirmação
        assunto = "Confirmação de Cadastro na Corrida"
        corpo = f"""
        Olá {nome},

        Seu cadastro na corrida foi realizado com sucesso!

        Seguem os detalhes do seu cadastro:
        - Nome: {nome}
        - Idade: {idade}
        - Sexo: {sexo}
        - Categoria: {categoria}
        - Telefone: {telefone}
        - Email: {email}
        - Cidade: {cidade}
        - CEP: {cep}
        - Logradouro: {logradouro}
        - Data de Nascimento: {data_nascimento}

        Obrigado por se inscrever!
        """

        if enviar_email(email, assunto, corpo):
            flash('Atleta cadastrado com sucesso! Um e-mail de confirmação foi enviado.', 'success')
        else:
            flash('Atleta cadastrado com sucesso, mas o e-mail de confirmação não pôde ser enviado.', 'warning')

        return redirect(url_for('index'))

    return render_template('cadastro.html')

@app.route('/visualizar', methods=['GET', 'POST'])
def visualizar_cadastro():
    if 'email' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        nome_busca = request.form['nome_busca']

        # Lendo o arquivo Excel
        df = pd.read_excel(EXCEL_FILE_ATLETAS)

        # Buscando o atleta pelo nome
        atleta = df[df['Nome'].str.lower() == nome_busca.lower()]

        if not atleta.empty:
            # Convertendo o DataFrame para um dicionário para facilitar o acesso aos dados
            atleta_info = atleta.to_dict('records')[0]
            return render_template('visualizar.html', atleta=atleta_info, encontrado=True)
        else:
            flash('Atleta não encontrado. Verifique o nome e tente novamente.', 'error')
            return redirect(url_for('visualizar_cadastro'))

    return render_template('visualizar.html', encontrado=False)

if __name__ == '__main__':
    app.run(debug=True)