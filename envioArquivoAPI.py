# Este script realiza as seguintes etapas:
# 1. Lê um arquivo JSON com a lista de colaboradores e documentos (gerado previamente).
# 2. Para cada item da lista:
#    - Envia o documento via API da ZapSign para que o colaborador assine.
# 3. Registra os envios bem-sucedidos.
# 4. Ao final, envia um e-mail para o responsável contendo um relatório dos documentos enviados com sucesso.
#
# Pré-requisitos:
# - O JSON deve estar no caminho: C:\rpa\departamentoPessoal\Premiacoes\lista Json\lista_envio_zapsign.json
# - Cada item no JSON deve conter: nome, matrícula, e-mail, nome do arquivo e link direto para o PDF.

import json
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import traceback

CAMINHO_JSON = r'C:\rpa\departamentoPessoal\Premiacoes\lista Json\lista_envio_zapsign.json'

ZAPSIGN_API_TOKEN = 'REMOVED_FOR_GITHUB'
URL_ZAPSIGN = 'https://api.zapsign.com.br/api/v1/docs/'

HEADERS = {
    'Authorization': f'Bearer {ZAPSIGN_API_TOKEN}',
    'Content-Type': 'application/json'
}


def enviar_para_zapsign(item):
    body = {
        "name": item['arquivo'],
        "external_id": item['matricula'],
        "url_pdf": item['url_pdf'],
        "folder_path": "/Termo de premiação",
        "signers": [
            {
                "name": item['nome'],
                "email": item['email'],
                "send_automatic_email": True
            }
        ]
    }

    resposta = requests.post(URL_ZAPSIGN, headers=HEADERS, json=body)

    if resposta.status_code in [200, 201]:
        print(f"Enviado: {item['arquivo']} {item['email']}")
        return True
    else:
        print(f"Falha: {item['arquivo']} {item['email']} ({resposta.status_code})")
        print(resposta.text)
        return False


def enviar_email_arquivos(processed_files):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'dti-admin@COMPANY_NAME.com.br'
    smtp_password = 'REMOVED_FOR_GITHUB'

    sender_email = 'dti-admin@COMPANY_NAME.com.br'
    receiver_email = 'raissa.correa@COMPANY_NAME.com.br'
    cc_emails = ['israel.martins@COMPANY_NAME.com.br', 'nicolas.nasario@COMPANY_NAME.com.br', 'lucas.remor@COMPANY_NAME.com.br']
    subject = 'RPA - Premiações'

    corpo_texto = "Prezado(a),\n\nSegue abaixo a lista de arquivos enviados:\n\n"
    corpo_texto += "Arquivo\tNome do Funcionário\tE-mail\n"
    corpo_texto += "-------\t-------------------\t------\n"
    for item in processed_files:
        corpo_texto += f"{item['arquivo']}\t{item['nome']}\t{item['email']}\n"

    corpo_html = """\
    <html>
      <body>
        <p>Prezados,<br><br>
           Segue abaixo a lista de arquivos processados:<br><br>
           <table border="1" cellpadding="5" cellspacing="0">
             <tr>
               <th>Arquivo</th>
               <th>Nome do Funcionário</th>
               <th>E-mail</th>
             </tr>
    """
    for item in processed_files:
        corpo_html += f"""\
             <tr>
               <td>{item['arquivo']}</td>
               <td>{item['nome']}</td>
               <td>{item['email']}</td>
             </tr>
        """
    corpo_html += """\
           </table>
           <br>E-mail automático, favor não responder!<br>
           <br>Atenciosamente,<br>
           Equipe RPA<br>
        </p>
      </body>
    </html>
    """

    mensagem = MIMEMultipart('alternative')
    mensagem['Subject'] = subject
    mensagem['From'] = sender_email
    mensagem['To'] = receiver_email
    mensagem['Cc'] = ", ".join(cc_emails)

    mensagem.attach(MIMEText(corpo_texto, 'plain'))
    mensagem.attach(MIMEText(corpo_html, 'html'))

    destinatarios = [receiver_email] + cc_emails

    try:
        servidor = smtplib.SMTP(smtp_server, smtp_port)
        servidor.starttls()
        servidor.login(smtp_username, smtp_password)
        servidor.sendmail(sender_email, destinatarios, mensagem.as_string())
        servidor.quit()
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        traceback.print_exc()


def main():
    with open(CAMINHO_JSON, 'r', encoding='utf-8') as f:
        lista = json.load(f)

    print(f"{len(lista)} arquivos para envio...\n")

    enviados = []

    for item in lista:
        if enviar_para_zapsign(item):
            enviados.append(item)

    if enviados:
        enviar_email_arquivos(enviados)


if __name__ == '__main__':
    main()
