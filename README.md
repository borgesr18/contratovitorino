# Gerador de Contratos

Este projeto cria um pequeno servidor em Python para gerar contratos a partir do arquivo `Contrato Vitorino.docx`.  
Os dados são preenchidos por um formulário web e o contrato gerado é enviado por e-mail.

## Como usar

1. **Configure as variáveis de ambiente** com as credenciais da conta de e-mail que será usada para o envio.  
   Exemplo usando a conta `contratodovitorino@gmail.com`:

   ```bash
   export EMAIL_USER="contratodovitorino@gmail.com"
   export EMAIL_PASS="gtim cijd snqy ekti"
   export EMAIL_DEST="rba1807@gmail.com"  # Opcional. Se não definido, usa este como padrão.
