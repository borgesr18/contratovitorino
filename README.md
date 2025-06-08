# Gerador de Contratos

Este projeto cria um pequeno servidor em Python para gerar contratos a partir do arquivo `Contrato Vitorino.docx`.

## Como usar

1. Defina variáveis de ambiente com as credenciais de envio de e-mail:
   - `EMAIL_USER` – usuário da conta (ex: gmail).
   - `EMAIL_PASS` – senha ou app password.
   - `EMAIL_DEST` – e-mail que receberá o contrato (opcional, padrão `rba1807@gmail.com`).
2. Execute o servidor:
   ```bash
   python app.py
   ```
3. Acesse `http://localhost:8000` e preencha o formulário. O contrato é gerado e enviado por e-mail como anexo.

O servidor utiliza apenas bibliotecas padrão do Python, então não é necessário instalar dependências adicionais.
