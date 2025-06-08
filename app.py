"""Aplicação para geração de contratos e envio por e-mail."""

import os
import re
from email.message import EmailMessage
from io import BytesIO
import smtplib
from flask import Flask, render_template, request
from docx import Document

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

def _replace_in_paragraph(paragraph, replacements):
    text = ''.join(run.text for run in paragraph.runs)
    for key, val in replacements.items():
        placeholder = f'[{key}]'
        while placeholder in text:
            start = text.index(placeholder)
            end = start + len(placeholder)
            count = 0
            start_idx = end_idx = None
            start_off = end_off = 0
            for i, run in enumerate(paragraph.runs):
                run_len = len(run.text)
                if start_idx is None and start < count + run_len:
                    start_idx = i
                    start_off = start - count
                if end <= count + run_len:
                    end_idx = i
                    end_off = end - count
                    break
                count += run_len
            if start_idx is None or end_idx is None:
                break
            first_run = paragraph.runs[start_idx]
            last_run = paragraph.runs[end_idx]
            prefix = first_run.text[:start_off]
            suffix = last_run.text[end_off:]
            first_run.text = prefix + val + suffix
            for j in range(end_idx, start_idx, -1):
                paragraph._p.remove(paragraph.runs[j]._r)
            text = ''.join(run.text for run in paragraph.runs)

def replace_placeholders(replacements):
    """Gera um novo DOCX substituindo marcadores do template de forma segura."""
    doc = Document(TEMPLATE_PATH)

    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_in_paragraph(paragraph, replacements)

    bio = BytesIO()
    doc.save(bio)
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
