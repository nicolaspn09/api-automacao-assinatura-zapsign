from google.oauth2 import service_account
from googleapiclient.discovery import build

# Caminho para a credencial da conta de serviço
CAMINHO_CREDENCIAL = r'C:\Users\COMPANY_NAME\Documents\envio ianca\credencial json\ornate-method-464813-a7-64be6d2adf33.json'

# Nome da pasta que será criada
NOME_PASTA = f"Uniforme"

# Lista de e-mails que devem ter acesso como editores
EMAILS_COMPARTILHAMENTO = [
    'israel.martins@COMPANY_NAME.com.br',
    'ianca.giassi@COMPANY_NAME.com.br'
]

def autenticar_drive():
    credenciais = service_account.Credentials.from_service_account_file(
        CAMINHO_CREDENCIAL,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=credenciais)

def criar_pasta_em_minha_unidade(service, nome_pasta):
    metadados = {
        'name': nome_pasta,
        'mimeType': 'application/vnd.google-apps.folder'
    }

    pasta = service.files().create(
        body=metadados,
        fields='id'
    ).execute()
    
    return pasta['id']

def compartilhar_com_editores(service, file_id, lista_emails):
    for email in lista_emails:
        permissao = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': email
        }
        service.permissions().create(
            fileId=file_id,
            body=permissao,
            sendNotificationEmail=False
        ).execute()

if __name__ == '__main__':
    service = autenticar_drive()
    id_pasta = criar_pasta_em_minha_unidade(service, NOME_PASTA)
    compartilhar_com_editores(service, id_pasta, EMAILS_COMPARTILHAMENTO)
    print(f'Pasta criada e compartilhada com sucesso. ID: {id_pasta}')
