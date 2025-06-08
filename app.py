"""Aplicação para geração de contratos e envio por e-mail."""

import os
import re
import zipfile
from email.message import EmailMessage
from io import BytesIO
import smtplib
from flask import Flask, render_template, request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Contrato Vitorino.docx")

FORM_FIELDS = {
    "Comprador": "Comprador",
    "CPF": "CPF",
    "RG": "RG",
    "Emissor": "Emissor",
    "EstadoCivil": "EstadoCivil",
    "Profissao": "Profissão",
    "Endereco": "Endereço",
    "Numero": "Número",
    "Complemento": "Complemento",
    "Bairro": "Bairro",
    "Cidade": "Cidade",
    "CEP": "CEP",
    "Quadra": "Quadra",
    "Lote": "Lote",
    "Testemunha": "Testemunha",
    "CPFTest": "CPF Test",
    "Testemunha2": "Testemunha2",
    "CPFTest2": "CPF Test2",
}

app = Flask(__name__)

def replace_placeholders(replacements):
    """Substitui marcadores no template Word preservando a estrutura do XML."""
    with zipfile.ZipFile(TEMPLATE_PATH) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        others = {n: z.read(n) for n in z.namelist() if n != "word/document.xml"}

    # Colapsa marcadores que foram quebrados em várias tags
    def collapse(match: re.Match) -> str:
        inner = re.sub(r"<[^>]+>", "", match.group(0))[1:-1]
        return "[" + inner + "]"

    xml = re.sub(r"\[.*?\]", collapse, xml, flags=re.DOTALL)

    for key, val in replacements.items():
        xml = xml.replace("[" + key + "]", val)

    with zipfile.ZipFile(TEMPLATE_PATH) as z:
        xml = z.read('word/document.xml').decode('utf-8')
        others = {n: z.read(n) for n in z.namelist() if n != 'word/document.xml'}
    # Merge adjacent text nodes so placeholders are contiguous
    xml = re.sub(r'</w:t>\s*</w:r>\s*<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>', '', xml)
    xml = re.sub(r'</w:t>\s*<w:t[^>]*>', '', xml)
    for key, val in replacements.items():
        xml = xml.replace('[' + key + ']', val)
    bio = BytesIO()
    with zipfile.ZipFile(bio, "w") as out:
        out.writestr("word/document.xml", xml)
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
