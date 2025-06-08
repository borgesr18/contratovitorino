import os
import re
import zipfile
from io import BytesIO
from email.message import EmailMessage
import smtplib
from flask import Flask, render_template, request

TEMPLATE_PATH = 'Contrato Vitorino.docx'

FORM_FIELDS = {
    'Comprador': 'Comprador',
    'CPF': 'CPF',
    'RG': 'RG',
    'Emissor': 'Emissor',
    'EstadoCivil': 'EstadoCivil',
    'Profissao': 'Profissao',
    'Endereco': 'Endereço',
    'Numero': 'Numero',
    'Complemento': 'Complemento',
    'Bairro': 'Bairro',
    'Cidade': 'Cidade',
    'CEP': 'CEP',
    'Quadra': 'Quadra',
    'Lote': 'Lote',
    'Testemunha': 'Testemunha',
    'CPFTest': 'CPF Test',
    'Testemunha2': 'Testemunha2',
    'CPFTest2': 'CPF Test2'
}

app = Flask(__name__)

from urllib.parse import parse_qs
from email.message import EmailMessage
import smtplib
from wsgiref.simple_server import make_server

TEMPLATE_PATH = 'Contrato Vitorino.docx'

FORM_HTML = """<!DOCTYPE html>
<html>
<head><meta charset='utf-8'><title>Gerar Contrato</title></head>
<body>
<h1>Gerar Contrato</h1>
<form method='POST' action='/generate'>
<p>Nome completo: <input name='Comprador' required></p>
<p>CPF: <input name='CPF' required></p>
<p>RG: <input name='RG' required></p>
<p>Órgão emissor: <input name='Emissor' required></p>
<p>Estado civil: <input name='EstadoCivil' required></p>
<p>Profissão: <input name='Profissao' required></p>
<p>Endereço: <input name='Endereço' required></p>
<p>Número: <input name='Numero' required></p>
<p>Complemento: <input name='Complemento'></p>
<p>Bairro: <input name='Bairro' required></p>
<p>Cidade: <input name='Cidade' required></p>
<p>CEP: <input name='CEP' required></p>
<p>Quadra: <input name='Quadra' required></p>
<p>Lote: <input name='Lote' required></p>
<p>Testemunha 1 nome: <input name='Testemunha' required></p>
<p>Testemunha 1 CPF: <input name='CPF Test' required></p>
<p>Testemunha 2 nome: <input name='Testemunha2' required></p>
<p>Testemunha 2 CPF: <input name='CPF Test2' required></p>
<p><button type='submit'>Gerar Contrato</button></p>
</form>
</body>
</html>"""

SUCCESS_HTML = """<!DOCTYPE html><html><body><h2>Contrato enviado com sucesso!</h2></body></html>"""
FAIL_HTML = """<!DOCTYPE html><html><body><h2>Falha ao enviar contrato.</h2></body></html>"""

PLACEHOLDERS = [
    'Comprador','CPF','RG','Emissor','EstadoCivil','Profissao',
    'Endereço','Numero','Complemento','Bairro','Cidade','CEP',
    'Quadra','Lote','Testemunha','CPF Test','Testemunha2','CPF Test2'
]

def replace_placeholders(replacements):
    with zipfile.ZipFile(TEMPLATE_PATH) as z:
        xml = z.read('word/document.xml').decode('utf-8')
        others = {n: z.read(n) for n in z.namelist() if n != 'word/document.xml'}
    for key, val in replacements.items():
        pattern = re.compile(r'<w:t>\[</w:t>.*?<w:t>' + re.escape(key) + r'</w:t>.*?<w:t>\]</w:t>', re.DOTALL)
        xml = pattern.sub('<w:t>' + val + '</w:t>', xml)
        xml = xml.replace('[' + key + ']', val)
    bio = BytesIO()
    with zipfile.ZipFile(bio, 'w') as out:
        out.writestr('word/document.xml', xml)
        for n, d in others.items():
            out.writestr(n, d)
    bio.seek(0)
    return bio.read()

def send_email(doc_bytes):
    user = os.environ.get('EMAIL_USER')
    password = os.environ.get('EMAIL_PASS')
    dest = os.environ.get('EMAIL_DEST', 'rba1807@gmail.com')
    if not user or not password:
        raise RuntimeError('Credenciais de e-mail não definidas')
    msg = EmailMessage()
    msg['Subject'] = 'Contrato Gerado'
    msg['From'] = user
    msg['To'] = dest
    msg.set_content('Segue contrato em anexo.')
    msg.add_attachment(
        doc_bytes,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename='contrato.docx'
    )
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(user, password)
        smtp.send_message(msg)


@app.route('/', methods=['GET', 'POST'])
def form():
    status = None
    if request.method == 'POST':
        data = {f: request.form.get(f, '') for f in FORM_FIELDS.keys()}
        replacements = {placeholder: data[field] for field, placeholder in FORM_FIELDS.items()}
        try:
            doc = replace_placeholders(replacements)
            send_email(doc)
            status = 'Contrato enviado com sucesso!'
        except Exception as exc:  # pragma: no cover - best effort
            print('Erro:', exc)
            status = 'Falha ao enviar contrato.'
    return render_template('form.html', status=status)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8000'))
    app.run(host='0.0.0.0', port=port)
def app(environ, start_response):
    path = environ.get('PATH_INFO', '/')
    method = environ.get('REQUEST_METHOD')
    if path == '/' and method == 'GET':
        start_response('200 OK', [('Content-Type', 'text/html; charset=utf-8')])
        return [FORM_HTML.encode('utf-8')]
    if path == '/generate' and method == 'POST':
        size = int(environ.get('CONTENT_LENGTH', 0))
        data = environ['wsgi.input'].read(size).decode('utf-8')
        params = {k: v[0] for k, v in parse_qs(data).items()}
        replacements = {k: params.get(k, '') for k in PLACEHOLDERS}
        try:
            doc = replace_placeholders(replacements)
            send_email(doc)
            start_response('200 OK', [('Content-Type', 'text/html; charset=utf-8')])
            return [SUCCESS_HTML.encode('utf-8')]
        except Exception as e:
            print('Erro:', e)
            start_response('500 Internal Server Error', [('Content-Type', 'text/html; charset=utf-8')])
            return [FAIL_HTML.encode('utf-8')]
    start_response('404 Not Found', [('Content-Type', 'text/plain')])
    return [b'Not found']

if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8000'))
    with make_server('', port, app) as srv:
        print(f'Servidor iniciado na porta {port}...')
        srv.serve_forever()
