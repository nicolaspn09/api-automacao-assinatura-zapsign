import openpyxl
import locale
from docx import Document
import os
import mysql.connector
from datetime import datetime
from dateutil.relativedelta import relativedelta
from docx2pdf import convert
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

CAMINHO_CREDENCIAL = r'C:\rpa\departamentoPessoal\Premiacoes\credencial json\core-plate-442111-s4-310713805904.json'
ID_PASTA_DRIVE = '14zF7KYR0Irvr_LOgEuUYZE_w1Mv2HHlq'  # ID pasta onde será enviado

#Função para envio para pasta via API
def upload_pdf_to_drive(file_path, file_name):
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    credentials = service_account.Credentials.from_service_account_file(CAMINHO_CREDENCIAL, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)

    file_metadata = {
        'name': file_name,
        'parents': [ID_PASTA_DRIVE]
    }

    media = MediaFileUpload(file_path, mimetype='application/pdf')

    uploaded_file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id',
        supportsAllDrives=True
    ).execute()

    print(f"Upload concluído: {file_name} ID {uploaded_file.get('id')}")
    return uploaded_file.get('id')

#Envia e-mail para os usuários
def envia_email(mensagemEmail, destinatarios_email, assunto_email):    
    # Configurações do servidor SMTP
    smtp_server = 'mail.COMPANY_NAME.com.br'
    smtp_port = 25  # Porta para comunicação segura com TLS

    # Credenciais do remetente
    remetente_email = "rpa@COMPANY_NAME.com.br"
    remetente_senha = 'REMOVED_FOR_GITHUB'

    destinatarios = destinatarios_email
    #destinatarios = [destinatarios_enviar]

    # Crie uma mensagem MIMEMultipart
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente_email
    mensagem['To'] = ",".join(destinatarios)
    mensagem['Subject'] = assunto_email

    # Adicione o corpo do e-mail
    corpo_email = mensagemEmail
    mensagem.attach(MIMEText(corpo_email, 'html'))  # 'plain' para texto simples ou 'html' para HTML

    try:
        servidor_smtp = smtplib.SMTP(smtp_server, smtp_port)
        servidor_smtp.starttls()  # Ative a criptografia TLS

        # Faça login com suas credenciais
        servidor_smtp.login(remetente_email, remetente_senha)

        # Envie o e-mail
        texto_email = mensagem.as_string()
        servidor_smtp.sendmail(remetente_email, destinatarios, texto_email)


    except Exception as e:
        #Bloco de logs
        locale.setlocale(locale.LC_ALL, 'pt_BR') #Seta o local
        data_hora_atual = datetime.now() #Busca a data atual
        mensagem = f"Erro ao enviar e-mail no primeiro código de clientes antecipados (clienteAntecipado): {str(e)}" #Informa a mensagem do Log

    finally:
        servidor_smtp.quit()  # Encerre a conexão com o servidor SMTP

#Roda query para executar o MySQL
def conecta_my_sql(sql):
    host = 'REMOVED'  # Endereço do servidor MySQL
    database = 'REMOVED'  # Nome do banco de dados
    user = 'REMOVED'  # Nome de usuário para acessar o banco de dados
    password = 'REMOVED_FOR_GITHUB'  # Senha do usuário para acessar o banco de dados

    try:
        # Estabelece a conexão com o banco de dados
        connection = mysql.connector.connect(
            host=host,
            database=database,
            user=user,
            password=password
        )
   
        if connection.is_connected():
            cursor = connection.cursor()
            cursor.execute(sql)
            tabela_sql = cursor.fetchall()
            data_referencia = tabela_sql[0]
            cursor.close()
            connection.close()

            #Retorna o resultado da consulta do SQL para o usuário
            return data_referencia
        
    except mysql.connector.Error as error:
        print("Erro")

sql = f"select * from fiscal.validacao_premiacoes"
retorno_sql = conecta_my_sql(sql=sql)
data_referencia = retorno_sql[0]
data_convertida = data_referencia

locale.setlocale(locale.LC_ALL, 'pt_BR')
# data_hora_atual = datetime.now()
# data_menos_um_mes = data_hora_atual - relativedelta(months=0)
# data_mais_um_mes = data_hora_atual + relativedelta(months=1)
# data_referencia = data_mais_um_mes.strftime('%B/%Y')
# data_convertida = data_menos_um_mes.strftime('%B/%Y')

# Carregar a planilha Excel
excel_path = r'C:\rpa\departamentoPessoal\Premiacoes\Backup\Premiacao_Farma.xlsx'
workbook = openpyxl.load_workbook(excel_path)
sheet = workbook.active

styles = getSampleStyleSheet()
title_font_style = ParagraphStyle('SmallFont', parent=styles['Title'], fontSize=10)
small_font_style = ParagraphStyle('SmallFont', parent=styles['Normal'], fontSize=9)
bold_font_style = ParagraphStyle('BoldFont', parent=styles['Normal'], fontSize=9, textColor='black', fontName='Helvetica-Bold')

def create_pdf(output_path, content):
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    story = []
    
    for paragraph in content:
        story.append(paragraph)
        story.append(Spacer(1, 7))
    
    doc.build(story)

contador_execucao = 0
matriculas_processadas = set()

# Percorrer as linhas da planilha a partir da linha 9
for row in sheet.iter_rows(min_row=8, values_only=True):
    try:
        cadastro = row[2]         
        nome_funcionario = row[3]
        status_funcionario = row[4]
        centro_custo = row[1]
        tipo_colaborador = row[43]

        if nome_funcionario != None and cadastro != None and "gerente" not in str(tipo_colaborador).lower():
            
            if cadastro not in matriculas_processadas:

                venda_farma = row[5]
                if venda_farma == None or venda_farma == "-" or str(venda_farma).lower() == "não se aplica" or venda_farma == "0,00" or venda_farma == "0" or venda_farma == "0.0" or venda_farma == "":
                    venda_farma_informar = 'Não se aplica'
                else:
                    venda_farma_informar = locale.currency(venda_farma, grouping=True, symbol=None)
                    
                valor_individual_venda_farma = row[35]
                if valor_individual_venda_farma == None or valor_individual_venda_farma == "-" or valor_individual_venda_farma == "Não se aplica":
                    valor_individual_venda_farma_informar = 'Não se aplica'
                else:
                    if venda_farma_informar == 'Não se aplica':
                        valor_individual_venda_farma_informar = 'Não se aplica'
                    else:     
                        valor_individual_venda_farma_informar = locale.currency(valor_individual_venda_farma, grouping=True, symbol=None)
                
                superavit_110 = row[7]
                if superavit_110 == None or superavit_110 == "-" or str(superavit_110).lower() == "não se aplica" or superavit_110 == "0,00" or superavit_110 == "0" or superavit_110 == "0.0" or superavit_110 == "":
                    superavit_110_informar = 'Não se aplica'
                else:
                    superavit_110_informar = locale.currency(superavit_110, grouping=True, symbol=None)
                
                valor_individual_superavit_110 = row[37]
                if valor_individual_superavit_110 == None or valor_individual_superavit_110 == "-":
                    valor_individual_superavit_110_informar = 'Não se aplica'
                else:
                    if superavit_110_informar == 'Não se aplica':
                        valor_individual_superavit_110_informar = 'Não se aplica'
                    else:     
                        valor_individual_superavit_110_informar = locale.currency(valor_individual_superavit_110, grouping=True, symbol=None)
                
                mex_meta_industria = row[8]
                if mex_meta_industria == None or mex_meta_industria == "-" or str(mex_meta_industria).lower() == "não se aplica" or mex_meta_industria == "0,00" or mex_meta_industria == "0" or mex_meta_industria == "0.0" or mex_meta_industria == "":
                    mex_meta_industria_informar = 'Não se aplica'
                else:
                    mex_meta_industria_informar = str(mex_meta_industria).replace(".0", "").strip()
                
                valor_individual_mex_meta_industrias = row[38]
                if valor_individual_mex_meta_industrias == None or valor_individual_mex_meta_industrias == "-" or str(valor_individual_mex_meta_industrias).lower() == "não se aplica" or valor_individual_mex_meta_industrias == "0,00" or valor_individual_mex_meta_industrias == "0" or valor_individual_mex_meta_industrias == "0.0" or valor_individual_mex_meta_industrias == "":
                    valor_individual_mex_meta_industrias_informar = 'Não se aplica'
                else:
                    if mex_meta_industria_informar == 'Não se aplica':
                        valor_individual_mex_meta_industrias_informar = 'Não se aplica'
                    else:     
                        valor_individual_mex_meta_industrias_informar = locale.currency(valor_individual_mex_meta_industrias, grouping=True, symbol=None)
                
                mex_meta_faturamento = row[9]
                if mex_meta_faturamento == None or mex_meta_faturamento == "-" or str(mex_meta_faturamento).lower() == "não se aplica" or mex_meta_faturamento == "0,00" or mex_meta_faturamento == "0" or mex_meta_faturamento == "0.0" or mex_meta_faturamento == "":
                    mex_meta_faturamento_informar = 'Não se aplica'
                else:
                    mex_meta_faturamento_informar = locale.currency(mex_meta_faturamento, grouping=True, symbol=None)
                
                valor_individual_mex_meta_faturamento = row[39]
                if valor_individual_mex_meta_faturamento == None or valor_individual_mex_meta_faturamento == "-":
                    valor_individual_mex_meta_faturamento_informar = 'Não se aplica'
                else:
                    if mex_meta_faturamento_informar == 'Não se aplica':
                        valor_individual_mex_meta_faturamento_informar = 'Não se aplica'
                    else:     
                        valor_individual_mex_meta_faturamento_informar = locale.currency(valor_individual_mex_meta_faturamento, grouping=True, symbol=None)

                flex_industrias = row[10]
                if flex_industrias == None or flex_industrias == "-" or str(flex_industrias).lower() == "não se aplica" or flex_industrias == "0,00" or flex_industrias == "0" or flex_industrias == "0.0" or flex_industrias == "":
                    flex_industrias_informar = 'Não se aplica'
                else:
                    flex_industrias_informar = str(flex_industrias).replace(".0", "").strip()
                
                valor_individual_flex_industrias = row[40]
                if valor_individual_flex_industrias == None or valor_individual_flex_industrias == "-":
                    valor_individual_flex_industrias_informar = 'Não se aplica'
                else:
                    if flex_industrias_informar == 'Não se aplica':
                        valor_individual_flex_industrias_informar = 'Não se aplica'
                    else:     
                        valor_individual_flex_industrias_informar = locale.currency(valor_individual_flex_industrias, grouping=True, symbol=None)

                prazo_medio = row[11]
                if prazo_medio == None or prazo_medio == "-" or str(prazo_medio).lower() == "não se aplica" or prazo_medio == "0,00" or prazo_medio == "0" or prazo_medio == "0.0" or prazo_medio == "":
                        prazo_medio_informar = 'Não se aplica'
                else:
                    prazo_medio_informar = str(prazo_medio).replace(".0", "").strip()
                
                valor_prazo_medio = row[41]
                if valor_prazo_medio == None or valor_prazo_medio == "-" or str(valor_prazo_medio).lower() == "não se aplica":
                        valor_individual_prazo_medio_informar_informar = 'Não se aplica'
                else:
                    if prazo_medio_informar == 'Não se aplica':
                        valor_individual_prazo_medio_informar_informar = 'Não se aplica'
                    else:     
                        valor_individual_prazo_medio_informar_informar = locale.currency(valor_prazo_medio, grouping=True, symbol=None)

                inadimplencia = row[12]
                if inadimplencia == None or inadimplencia == "-" or str(inadimplencia).lower() == "não se aplica" or inadimplencia == "0,00" or inadimplencia == "0" or inadimplencia == "0.0" or inadimplencia == "":
                        inadimplencia_informar = 'Não se aplica'
                else:
                    inadimplencia_novo = inadimplencia * 100
                    inadimplencia_informar = locale.currency(inadimplencia_novo, grouping=True, symbol=None) + '%'
                
                valor_inadimplencia = row[42]
                if valor_inadimplencia == None or valor_inadimplencia == "-" or str(valor_inadimplencia).lower() == "não se aplica":
                        valor_individual_inadimplencia_informar = 'Não se aplica'
                else:
                    if inadimplencia_informar == 'Não se aplica':
                        valor_individual_inadimplencia_informar = 'Não se aplica'
                    else:     
                        valor_individual_inadimplencia_informar = locale.currency(valor_inadimplencia, grouping=True, symbol=None)

                if status_funcionario == None or status_funcionario == "":
                    status_funcionario_informar = 'Trabalhando'
                else:
                    if str(status_funcionario) == "0" or "-" not in str(status_funcionario):
                        status_funcionario_informar = 'Trabalhando'
                    else:
                        status_funcionario_informar = f"""Férias {str(status_funcionario).split("-")[1].replace(".0", "").strip()} dias"""

                    if venda_farma_informar != "Não se aplica" or superavit_105_informar != "Não se aplica" or superavit_110_informar != "Não se aplica" or mex_meta_faturamento_informar != "Não se aplica" or prazo_medio_informar != "Não se aplica" or inadimplencia_informar != "Não se aplica":
                        pdf_content = []
                        pdf_content.append(Paragraph('TERMO DE PARTICIPAÇÃO NA PREMIAÇÃO POR COBERTURA DE OBJETIVO', title_font_style))
                        pdf_content.append(Paragraph(f'Venho através do presente, por minha livre e espontânea iniciativa, participar das premiações por cobertura de objetivo conforme abaixo especificado:',  small_font_style))
                        pdf_content.append(Paragraph(f'Mês de Referência: <b>{data_referencia}</b>', small_font_style))
                        pdf_content.append(Paragraph('- Indicadores -', bold_font_style))
                        pdf_content.append(Paragraph(f'1. Vendas Farma - R$: <b>{venda_farma_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'Valor Individual a Receber – R$: <b>{valor_individual_venda_farma_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'2. Superavit 110 - R$: <b>{superavit_110_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'Valor Individual a Receber – R$: <b>{valor_individual_superavit_110_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'3. MEX Meta Indústrias (Atingir 100% das vendas de cada indústria): <b>{mex_meta_industria_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_mex_meta_industrias_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'4. MEX Meta Faturamento - R$: <b>{mex_meta_faturamento_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_mex_meta_faturamento_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'5. Flex Indústrias: <b>{flex_industrias_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_flex_industrias_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'6. Prazo Médio: <b>{prazo_medio_informar}</b>', small_font_style)) 
                        pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_prazo_medio_informar_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'7. Inadimplência - Os títulos vencidos a mais de 30 dias deverão ser movimentados em 10% sobre o saldo devedor em atraso e abaixo de: <b>{inadimplencia_informar}</b>', small_font_style)) 
                        pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_inadimplencia_informar}</b>', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph('Observações/Regras de Apuração', bold_font_style))
                        pdf_content.append(Paragraph(f'● Os valores serão pagos em folha referente a competência do mês de apuração, ou seja, após o fechamento de cada mês, apurando-se os respectivos valores individuais que serão pagos. Por exemplo: Vendas de Fevereiro serão apuradas em Março e pagas em Abril, que se refere à competência de Março.', small_font_style))
                        pdf_content.append(Paragraph(f'● No período de férias os valores são pagos proporcionais aos dias trabalhados.', small_font_style))
                        pdf_content.append(Paragraph(f'● A adesão a presente é livre e a não cobertura da cota não ensejará qualquer tipo de punição ou advertência ao colaborador, ocorrendo apenas o não recebimento da respectiva premiação.', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'', small_font_style))
                        pdf_content.append(Paragraph(f'Estou plenamente ciente de seu teor.',  small_font_style))
                        pdf_content.append(Paragraph(f'Tubarão (SC) / Santa Cruz do Sul (RS) / São José dos Pinhais (PR) / Serra (ES) / Sobradinho (DF),  01   de  {data_convertida}.', small_font_style))
                        pdf_content.append(Spacer(1, 18))  # Adiciona um espaço extra antes do parágrafo específico
                        pdf_content.append(Paragraph(f'Assinatura: _______________________________________________', small_font_style))
                        pdf_content.append(Paragraph(f'Nome : {nome_funcionario}', small_font_style))
                        pdf_content.append(Paragraph(f'Cadastro: {cadastro_informar}', small_font_style))
                        pdf_content.append(Paragraph(f'Centro de Custo: {centro_custo}', small_font_style))
                        pdf_content.append(Paragraph(f'Status – {status_funcionario_informar}', small_font_style))

                        output_folder = r'C:\rpa\departamentoPessoal\Premiacoes\docs temp'
                        os.makedirs(output_folder, exist_ok=True)
                        pdf_output_path = os.path.join(output_folder, f'{cadastro_informar}.pdf')

                        create_pdf(pdf_output_path, pdf_content)
                        upload_pdf_to_drive(pdf_output_path, os.path.basename(pdf_output_path))
                        os.remove(pdf_output_path)

        elif nome_funcionario != None and cadastro != None and "gerente" in str(tipo_colaborador).lower():
            venda_farma = row[5]
            if venda_farma == None or venda_farma == "-" or str(venda_farma).lower() == "não se aplica" or venda_farma == "0,00" or venda_farma == "0" or venda_farma == "0.0" or venda_farma == "":
                venda_farma_informar = 'Não se aplica'
            else:
                venda_farma_informar = locale.currency(venda_farma, grouping=True, symbol=None)
                
            valor_individual_venda_farma = row[35]
            if valor_individual_venda_farma == None or valor_individual_venda_farma == "-" or valor_individual_venda_farma == "Não se aplica":
                valor_individual_venda_farma_informar = 'Não se aplica'
            else:
                if venda_farma_informar == 'Não se aplica':
                    valor_individual_venda_farma_informar = 'Não se aplica'
                else:     
                    valor_individual_venda_farma_informar = locale.currency(valor_individual_venda_farma, grouping=True, symbol=None)
            
            superavit_105 = row[6]
            if superavit_105 == None or superavit_105 == "-" or str(superavit_105).lower() == "não se aplica" or superavit_105 == "0,00" or superavit_105 == "0" or superavit_105 == "0.0" or superavit_105 == "":
                superavit_105_informar = 'Não se aplica'
            else:
                superavit_105_informar = locale.currency(superavit_105, grouping=True, symbol=None)
            
            valor_individual_superavit_105 = row[36]
            if valor_individual_superavit_105 == None or valor_individual_superavit_105 == "-":
                valor_individual_superavit_105_informar = 'Não se aplica'
            else:
                if superavit_105_informar == 'Não se aplica':
                    valor_individual_superavit_105_informar = 'Não se aplica'
                else:     
                    valor_individual_superavit_105_informar = locale.currency(valor_individual_superavit_105, grouping=True, symbol=None)
            
            superavit_110 = row[7]
            if superavit_110 == None or superavit_110 == "-" or str(superavit_110).lower() == "não se aplica" or superavit_110 == "0,00" or superavit_110 == "0" or superavit_110 == "0.0" or superavit_110 == "":
                superavit_110_informar = 'Não se aplica'
            else:
                superavit_110_informar = locale.currency(superavit_110, grouping=True, symbol=None)
            
            valor_individual_superavit_110 = row[37]
            if valor_individual_superavit_110 == None or valor_individual_superavit_110 == "-":
                valor_individual_superavit_110_informar = 'Não se aplica'
            else:
                if superavit_110_informar == 'Não se aplica':
                    valor_individual_superavit_110_informar = 'Não se aplica'
                else:     
                    valor_individual_superavit_110_informar = locale.currency(valor_individual_superavit_110, grouping=True, symbol=None)
            
            mex_meta_industrias = row[8]
            if mex_meta_industrias == None or mex_meta_industrias == "-" or str(mex_meta_industrias).lower() == "não se aplica" or mex_meta_industrias == "0,00" or mex_meta_industrias == "0" or mex_meta_industrias == "0.0" or mex_meta_industrias == "":
                mex_meta_industrias_informar = 'Não se aplica'
            else:
                mex_meta_industrias_informar = str(mex_meta_industrias).replace(".0", "").strip()
            
            valor_individual_mex_meta_industrias = row[38]
            if valor_individual_mex_meta_industrias == None or valor_individual_mex_meta_industrias == "-" or str(valor_individual_mex_meta_industrias).lower() == "não se aplica" or valor_individual_mex_meta_industrias == "0,00" or valor_individual_mex_meta_industrias == "0" or valor_individual_mex_meta_industrias == "0.0" or valor_individual_mex_meta_industrias == "":
                valor_individual_mex_meta_industrias_informar = 'Não se aplica'
            else:
                if mex_meta_industrias_informar == 'Não se aplica':
                    valor_individual_mex_meta_industrias_informar = 'Não se aplica'
                else:     
                    valor_individual_mex_meta_industrias_informar = locale.currency(valor_individual_mex_meta_industrias, grouping=True, symbol=None)
            
            mex_meta_faturamento = row[9]
            if mex_meta_faturamento == None or mex_meta_faturamento == "-" or str(mex_meta_faturamento).lower() == "não se aplica" or mex_meta_faturamento == "0,00" or mex_meta_faturamento == "0" or mex_meta_faturamento == "0.0" or mex_meta_faturamento == "":
                mex_meta_faturamento_informar = 'Não se aplica'
            else:
                mex_meta_faturamento_informar = locale.currency(mex_meta_faturamento, grouping=True, symbol=None)
            
            valor_individual_mex_meta_faturamento = row[39]
            if valor_individual_mex_meta_faturamento == None or valor_individual_mex_meta_faturamento == "-":
                valor_individual_mex_meta_faturamento_informar = 'Não se aplica'
            else:
                if mex_meta_faturamento_informar == 'Não se aplica':
                    valor_individual_mex_meta_faturamento_informar = 'Não se aplica'
                else:     
                    valor_individual_mex_meta_faturamento_informar = locale.currency(valor_individual_mex_meta_faturamento, grouping=True, symbol=None)

            flex_industrias = row[10]
            if flex_industrias == None or flex_industrias == "-" or str(flex_industrias).lower() == "não se aplica" or flex_industrias == "0,00" or flex_industrias == "0" or flex_industrias == "0.0" or flex_industrias == "":
                flex_industrias_informar = 'Não se aplica'
            else:
                flex_industrias_informar = str(flex_industrias).replace(".0", "").strip()
            
            valor_individual_flex_industrias = row[40]
            if valor_individual_flex_industrias == None or valor_individual_flex_industrias == "-":
                valor_individual_flex_industrias_informar = 'Não se aplica'
            else:
                if flex_industrias_informar == 'Não se aplica':
                    valor_individual_flex_industrias_informar = 'Não se aplica'
                else:     
                    valor_individual_flex_industrias_informar = locale.currency(valor_individual_flex_industrias, grouping=True, symbol=None)

            prazo_medio = row[11]
            if prazo_medio == None or prazo_medio == "-" or str(prazo_medio).lower() == "não se aplica" or prazo_medio == "0,00" or prazo_medio == "0" or prazo_medio == "0.0" or prazo_medio == "":
                prazo_medio_informar = 'Não se aplica'
            else:
                prazo_medio_informar = str(prazo_medio).replace(".0", "").strip()
            
            valor_prazo_medio = row[41]
            if valor_prazo_medio == None or valor_prazo_medio == "-" or str(valor_prazo_medio).lower() == "não se aplica":
                    valor_individual_prazo_medio_informar_informar = 'Não se aplica'
            else:
                if prazo_medio_informar == 'Não se aplica':
                    valor_individual_prazo_medio_informar_informar = 'Não se aplica'
                else:     
                    valor_individual_prazo_medio_informar_informar = locale.currency(valor_prazo_medio, grouping=True, symbol=None)

            inadimplencia = row[12]
            if inadimplencia == None or inadimplencia == "-" or str(inadimplencia).lower() == "não se aplica" or inadimplencia == "0,00" or inadimplencia == "0" or inadimplencia == "0.0" or inadimplencia == "":
                    inadimplencia_informar = 'Não se aplica'
            else:
                inadimplencia_novo = inadimplencia * 100
                inadimplencia_informar = locale.currency(inadimplencia_novo, grouping=True, symbol=None) + '%'
            
            valor_inadimplencia = row[42]
            if valor_inadimplencia == None or valor_inadimplencia == "-" or str(valor_inadimplencia).lower() == "não se aplica":
                    valor_individual_inadimplencia_informar = 'Não se aplica'
            else:
                if inadimplencia_informar == 'Não se aplica':
                    valor_individual_inadimplencia_informar = 'Não se aplica'
                else:     
                    valor_individual_inadimplencia_informar = locale.currency(valor_inadimplencia, grouping=True, symbol=None)

            if cadastro != None:
                cadastro_informar = int(cadastro)
                if status_funcionario == None or status_funcionario == "":
                    status_funcionario_informar = 'Trabalhando'
                else:
                    if str(status_funcionario) == "0" or "-" not in str(status_funcionario):
                        status_funcionario_informar = 'Trabalhando'
                    else:
                        status_funcionario_informar = f"""Férias {str(status_funcionario).split("-")[1].replace(".0", "").strip()} dias"""

                if venda_farma_informar != "Não se aplica" or superavit_105_informar != "Não se aplica" or superavit_110_informar != "Não se aplica" or mex_meta_faturamento_informar != "Não se aplica" or prazo_medio_informar != "Não se aplica" or inadimplencia_informar != "Não se aplica":

                    pdf_content = []
                    pdf_content.append(Paragraph('TERMO DE PARTICIPAÇÃO NA PREMIAÇÃO POR COBERTURA DE OBJETIVO', title_font_style))
                    pdf_content.append(Paragraph(f'Venho através do presente, por minha livre e espontânea iniciativa, participar das premiações por cobertura de objetivo conforme abaixo especificado:',  small_font_style))
                    pdf_content.append(Paragraph(f'Mês de Referência: <b>{data_referencia}</b>', small_font_style))
                    pdf_content.append(Paragraph('- Indicadores -', bold_font_style))
                    pdf_content.append(Paragraph(f'1. Vendas Farma - R$: <b>{venda_farma_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'Valor Individual a Receber – R$: <b>{valor_individual_venda_farma_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'2. Superavit 105 - R$: <b>{superavit_105_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'Valor Individual a Receber – R$: <b>{valor_individual_superavit_105_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'3. Superavit 110 - R$: <b>{superavit_110_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_superavit_110_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'4. MEX Meta Indústrias (Atingir 100% das vendas de cada indústria): <b>{mex_meta_industrias_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_mex_meta_industrias_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'5. MEX Meta Faturamento - R$: <b>{mex_meta_faturamento_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_mex_meta_faturamento_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'6. Flex Indústrias: <b>{flex_industrias_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'Premiação máxima de até - R$: <b>{valor_individual_flex_industrias_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'7. Prazo Médio: <b>{prazo_medio_informar}</b>', small_font_style)) 
                    pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_prazo_medio_informar_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'8. Inadimplência - Os títulos vencidos a mais de 30 dias deverão ser movimentados em 10% sobre o saldo devedor em atraso e abaixo de: <b>{inadimplencia_informar}</b>', small_font_style)) 
                    pdf_content.append(Paragraph(f'Valor Individual a Receber - R$: <b>{valor_individual_inadimplencia_informar}</b>', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph('Observações/Regras de Apuração', bold_font_style))
                    pdf_content.append(Paragraph(f'● Os valores serão pagos em folha referente a competência do mês de apuração, ou seja, após o fechamento de cada mês, apurando-se os respectivos valores individuais que serão pagos. Por exemplo: Vendas de Fevereiro serão apuradas em Março e pagas em Abril, que se refere à competência de Março.', small_font_style))
                    pdf_content.append(Paragraph(f'● No período de férias os valores são pagos proporcionais aos dias trabalhados.', small_font_style))
                    pdf_content.append(Paragraph(f'● A adesão a presente é livre e a não cobertura da cota não ensejará qualquer tipo de punição ou advertência ao colaborador, ocorrendo apenas o não recebimento da respectiva premiação.', small_font_style))
                    pdf_content.append(Paragraph(f'● O valor da premiação de INDUSTRIAS FLEX, será dividido entre os KPIS participantes.', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'', small_font_style))
                    pdf_content.append(Paragraph(f'Estou plenamente ciente de seu teor.',  small_font_style))
                    pdf_content.append(Paragraph(f'Tubarão (SC) / Santa Cruz do Sul (RS) / São José dos Pinhais (PR) / Serra (ES) / Sobradinho (DF),  01   de  {data_convertida}.', small_font_style))
                    pdf_content.append(Spacer(1, 18))  # Adiciona um espaço extra antes do parágrafo específico
                    pdf_content.append(Paragraph(f'Assinatura: _______________________________________________', small_font_style))
                    pdf_content.append(Paragraph(f'Nome : {nome_funcionario}', small_font_style))
                    pdf_content.append(Paragraph(f'Cadastro: {cadastro_informar}', small_font_style))
                    pdf_content.append(Paragraph(f'Centro de Custo: {centro_custo}', small_font_style))
                    pdf_content.append(Paragraph(f'Status – {status_funcionario_informar}', small_font_style))

                    output_folder = r'C:\rpa\departamentoPessoal\Premiacoes\docs temp'
                    os.makedirs(output_folder, exist_ok=True)
                    pdf_output_path = os.path.join(output_folder, f'{cadastro_informar}.pdf')

                    create_pdf(pdf_output_path, pdf_content)
                    upload_pdf_to_drive(pdf_output_path, os.path.basename(pdf_output_path))
                    os.remove(pdf_output_path)

            else:
                pass

        else:
            contador_execucao += 1

            #Caso tenha mais de 5 linhas vazias, estoura o laço pra finalizar o código
            if contador_execucao == 5:
                break
    
    except Exception as e:
       destinatarios_email = []
       destinatarios_email.append("nicolas.nasario@COMPANY_NAME.com.br")
       destinatarios_email.append("lucas.remor@COMPANY_NAME.com.br")
       destinatarios_email.append("raissa.correa@COMPANY_NAME.com.br")
       destinatarios_email.append("israel.martins@COMPANY_NAME.com.br")

       mensagemEmail = f"Olá!<br><br> Há erro na parte do farma: {e}"

       assunto_email = f"Erro - RPA Premiações Farma"

       envia_email(mensagemEmail, destinatarios_email, assunto_email) 

print("Processo concluído.")