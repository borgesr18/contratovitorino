# Gerador de Contratos

Este projeto cria um pequeno servidor em Python para gerar contratos a partir do arquivo `Contrato Vitorino.docx`.

## Como usar

1. Defina variáveis de ambiente com as credenciais de envio de e-mail (exemplo abaixo usa a conta **contratodovitorino@gmail.com**):
   - `EMAIL_USER` – usuário da conta. Exemplo:

     ```bash
     export EMAIL_USER="contratodovitorino@gmail.com"
     export EMAIL_PASS="gtim cijd snqy ekti"
     ```

   - `EMAIL_DEST` – e-mail que receberá o contrato (opcional, padrão `rba1807@gmail.com`).

2. Instale as dependências necessárias:

   ```bash
   pip install -r requirements.txt
