import os
import json
import traceback
import oracledb
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Configurações fixas
CAMINHO_CREDENCIAL = r'C:\rpa\departamentoPessoal\Premiacoes\credencial json\core-plate-442111-s4-310713805904.json'
ID_PASTA_DRIVE = '14zF7KYR0Irvr_LOgEuUYZE_w1Mv2HHlq'
ID_PASTA_CLONE = '1OHHBKDw9q5tKg2q-JHe38ghul19zq2fM'
DESTINATARIOS = ["israel.martins@COMPANY_NAME.com.br"]
CAMINHO_JSON_SAIDA = r'C:\rpa\departamentoPessoal\Premiacoes\lista Json\lista_envio_zapsign.json'

lista_para_envio = []

# Envio de e-mail
def envia_email(mensagemEmail, destinatarios_email, assunto_email):
    smtp_server = 'mail.COMPANY_NAME.com.br'
    smtp_port = 25
    remetente_email = "rpa@COMPANY_NAME.com.br"
    remetente_senha = 'REMOVED_FOR_GITHUB'

    mensagem = MIMEMultipart()
    mensagem['From'] = remetente_email
    mensagem['To'] = ",".join(destinatarios_email)
    mensagem['Subject'] = assunto_email
    mensagem.attach(MIMEText(mensagemEmail, 'html'))

    try:
        servidor_smtp = smtplib.SMTP(smtp_server, smtp_port)
        servidor_smtp.starttls()
        servidor_smtp.login(remetente_email, remetente_senha)
        servidor_smtp.sendmail(remetente_email, destinatarios_email, mensagem.as_string())
    except Exception as e:
        pass
    finally:
        servidor_smtp.quit()

# Google Drive
def autenticar_google_drive():
    creds = service_account.Credentials.from_service_account_file(CAMINHO_CREDENCIAL, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds)

def listar_arquivos_pdf(service, pasta_id):
    query = f"'{pasta_id}' in parents and mimeType='application/pdf' and trashed=false"
    arquivos = []
    page_token = None

    while True:
        resposta = service.files().list(
            q=query,
            spaces='drive',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            fields='nextPageToken, files(id, name)',
            pageToken=page_token
        ).execute()
        arquivos.extend(resposta.get('files', []))
        page_token = resposta.get('nextPageToken', None)
        if not page_token:
            break
    return arquivos


def gerar_link_direto(file_id):
    return f"https://drive.google.com/uc?export=download&id={file_id}"

# Oracle
def consultar_colaborador_por_matricula(matricula):
    try:
        oracledb.init_oracle_client(lib_dir=None)
        connection = oracledb.connect(
            user = 'REMOVED',
            password = 'REMOVED_FOR_GITHUB',
            dsn="10.1.1.20:1521/pdb1"
        )
        cursor = connection.cursor()
        query = f"""
            SELECT EMACOM, B.NOMFUN 
            FROM VETORH.R034CPL@SAPIENS A, VETORH.R034FUN@SAPIENS B
            WHERE A.NUMCAD = B.NUMCAD
            AND A.NUMEMP = B.NUMEMP
            AND B.NUMCAD = '{matricula}'
            AND VALSAL > 0
            AND A.NUMEMP = 1
            AND A.TIPCOL = 1
            ORDER BY 1
        """
        cursor.execute(query)
        resultado = cursor.fetchone()
        return resultado if resultado else (None, None)
    except Exception as e:
        envia_email(f"Erro Oracle: {e}", DESTINATARIOS, "Erro - Consulta Oracle")
        return None, None
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'connection' in locals():
            connection.close()

def mover_arquivo_para_clone(service, file_id, pasta_destino_id):
    try:
        service.files().update(
            fileId=file_id,
            addParents=pasta_destino_id,
            removeParents=ID_PASTA_DRIVE,
            fields='id, parents',
            supportsAllDrives=True
        ).execute()
    except Exception as e:
        envia_email(f"Erro ao mover arquivo ID {file_id}:<br>{e}", DESTINATARIOS, "Erro - Mover Arquivo")

def processar_arquivo(service, arq):
    nome = arq['name']
    if not nome.lower().endswith('.pdf'):
        return

    matricula = nome.replace('.pdf', '').strip()
    email, nome_funcionario = consultar_colaborador_por_matricula(matricula)

    if email and nome_funcionario:
        link = gerar_link_direto(arq['id'])
        lista_para_envio.append({
            'arquivo': nome,
            'matricula': matricula,
            'nome': nome_funcionario,
            'email': email,
            'url_pdf': link
        })
        mover_arquivo_para_clone(service, arq['id'], ID_PASTA_CLONE)
    else:
        envia_email(f"Matrícula {matricula} não encontrada no banco.", DESTINATARIOS, "Aviso - Matrícula Não Encontrada")

def main():
    global lista_para_envio
    try:
        service = autenticar_google_drive()
        arquivos = listar_arquivos_pdf(service, ID_PASTA_DRIVE)

        for arq in arquivos:
            processar_arquivo(service, arq)

        try:
            with open(CAMINHO_JSON_SAIDA, 'w', encoding='utf-8') as f:
                json.dump(lista_para_envio, f, ensure_ascii=False, indent=2)

            envia_email(
                f"Script finalizado com sucesso.<br>Registros: {len(lista_para_envio)}<br>Arquivo: {CAMINHO_JSON_SAIDA}",
                DESTINATARIOS,
                "Sucesso - JSON Premiações"
            )

        except Exception as erro_arquivo:
            envia_email(f"Erro ao salvar JSON:<br>{erro_arquivo}", DESTINATARIOS, "Erro - Salvando JSON")

    except Exception as erro_geral:
        envia_email(f"Erro geral no script:<br>{erro_geral}", DESTINATARIOS, "Erro - Execução Geral")

if __name__ == '__main__':
    main()