import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import traceback
import os
import shutil
from datetime import datetime
import locale
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
from selenium.webdriver.common.action_chains import ActionChains

def main():
    # Definir a localidade
    locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')

    # Definir o diretório de download
    download_dir = r"C:\rpa\departamentoPessoal\Premiacoes\Arquivos baixados premiacao"

    # Criar novo diretório com o mês atual e ano no formato MÊS_ANO (em português)
    current_date = datetime.now()
    month_year = current_date.strftime("%B_%Y")  # Exemplo: 'outubro_2024'
    new_directory = os.path.join(r"G:\.shortcut-targets-by-id\0B7kvjs8r64dlWTBnWUozODFOeTg\Premiações\Termos assinados", month_year)

    if not os.path.exists(new_directory):
        os.makedirs(new_directory)  # Cria o diretório se ele não existir

    print(f"Arquivos baixados serão movidos para: {new_directory}")

    # Configurar as preferências do Chrome
    options = uc.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,  # Diretório de download específico
        "download.prompt_for_download": False,  # Não perguntar onde salvar
        "download.directory_upgrade": True,  # Atualiza o diretório de download padrão
        "safebrowsing.enabled": True  # Habilitar downloads seguros
    }
    options.add_experimental_option("prefs", prefs)

    # Definir um User-Agent personalizado (opcional)
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.5938.132 Safari/537.36")

    print("Tentando abrir navegador")
    
    # Inicializar o navegador com as opções configuradas
    driver = uc.Chrome(options=options)

    try:
        login_site(driver)
        print("Login realizado com sucesso..")
        
        navegar_dentro_site(driver)
        print("Navegação dentro site, Ok..")
        
        documentos_baixados = loop_baixar_assinados(driver)  # Agora retorna a quantidade de documentos baixados
        print("Terminei de baixar os documentos..")

        # Passando a variável corretamente
        arquivos_movidos = mover_arquivos_baixados(download_dir, new_directory, documentos_baixados)
        print("Arquivos movidos para pasta do drive")

        
        # Processo de download
        #loop_baixar_assinados(driver)
        #print("Terminei de baixar os documentos..")

        # Mover arquivos baixados
        #arquivos_movidos = mover_arquivos_baixados(download_dir, new_directory, documentos_baixados)
        #print("Arquivos movidos para pasta do drive")
        
        # Exclusão de arquivos na página web
        if arquivos_movidos:
            backup_finalizados(driver)

        # Enviar e-mail com os detalhes
        enviar_email(arquivos_movidos, new_directory)

        print("Processo finalizado com sucesso.")

    except Exception as e:
        print(f"Erro no processo principal: {e}")
        traceback.print_exc()
    finally:
        if driver:
            try:
                driver.quit()
                print("WebDriver encerrado com sucesso.")
            except Exception as e:
                print(f"Erro ao encerrar o WebDriver: {e}")

def login_site(driver):
    driver.get("https://app.zapsign.com.br/acesso/entrar")

    login = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-login/div/div/div[1]/form/zs-text-input/mat-form-field/div/div[1]/div[1]/input")))
    login.send_keys('dp@COMPANY_NAME.com.br')

    senha = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-login/div/div/div[1]/form/zs-password-input/mat-form-field/div/div[1]/div[1]/input")))
    senha.send_keys('COMPANY_NAME@')

    driver.find_element(By.XPATH, "/html/body/app-root/div/app-login/div/div/div[1]/form/div/zs-button/button/span").click()
    time.sleep(2)

def navegar_dentro_site(driver):
    
    try:
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/app-folders/div/div/div/app-folder-tree/div/div/div[2]/app-folder[1]")))
        
        time.sleep(1)
        
        for i in range(1,20):            
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/app-folders/div/div/div/app-folder-tree/div/div/div[2]/app-folder[{i}]")))
            
            departamento_pessoal = driver.find_element(By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/app-folders/div/div/div/app-folder-tree/div/div/div[2]/app-folder[{i}]")
            
            valor = departamento_pessoal.text
            
            time.sleep(1)
            
            if "DEPARTAMENTO PESSOAL" in valor.upper():                
                time.sleep(1)
                departamento_pessoal.click()
                break                
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/app-folders/div/div/div/app-folder-tree/div/div/div[2]/app-folder[2]/div/div[1]/p")))
        time.sleep(2)
        
        for i in range (1,20):
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/app-folders/div/div/div/app-folder-tree/div/div/div[2]/app-folder[{i}]/div")))
            nome_opcao = driver.find_element(By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/app-folders/div/div/div/app-folder-tree/div/div/div[2]/app-folder[{i}]/div")
            valor_texto = nome_opcao.text
            time.sleep(2)
            if "Termo de Premiação" in valor_texto:
                nome_opcao.click()
                break
            #else:
                #print("Não consegui achar o elemento de click -Termo de Premiação-")
                    
        time.sleep(1)
        
        #finalizados
        for i in range (1,3):
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-filter-status/div/div/div[{i}]")))
            
            elemento_finalizado = driver.find_element(By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-filter-status/div/div/div[{i}]")
            
            valor_finalizado = elemento_finalizado.text
            time.sleep(1)
            if "Finalizados" in valor_finalizado:
                elemento_finalizado.click()
                break

        time.sleep(3)

        #Click finalizados
        #WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-filter-status/div/div/div[1]"))).click()
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-action-manager/div/div/app-select-to-input/div/mat-form-field/div/div[1]/div/mat-select/div/div[2]"))).click()

        time.sleep(1)
        
        try:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div/div/mat-option[4]"))).click()
        except:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/div/div/div/mat-option[4]"))).click()

        time.sleep(1)


    except Exception as e:
        print(f"Erro ao navegar no site: {e}")
        traceback.print_exc()
        raise

def loop_baixar_assinados(driver):
    try:
        numero_documentos = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-filter-status/div/div/div[1]/div[2]/h3")))

        total_documentos = int(numero_documentos.text)
        
        documentos_baixados = 0 #contar a quantidade de downloads
        
        if total_documentos == 0:
            raise Exception("Nenhum documento encontrado para download.")
        
        print(f"Total de documentos finalizados para baixar: {total_documentos}")       
        
        for contador in range(1, total_documentos + 1):
            try:
                #Verifica se o documento está com a "bandeira" ASSINADO
                status_assinado = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-documents-list/div/div[{contador}]/app-document-item-list/a/app-documents-item-header/div/div[2]/app-doc-status-chip/div/div")))
                
                status_text = status_assinado.text               
                
                #valida bandeira ASSINADO
                if status_text == "ASSINADO":
                    
                    print( f"Documento de número {contador}, está assinado, seguindo para o download..")
                                
                #time.sleep(2)
                #WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-filter-status/div/div/div[1]"))).click()

                    #documento conforme decorrer o para
                    documento_xpath = f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-documents-list/div/div[{contador}]"
                    time.sleep(2)

                    #Clica conforme for decorrendo o 'FOR'
                    documento = WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, documento_xpath)))
                    documento.click()

                    time.sleep(5)

                    #Coleta URL do download do documento
                    url_download = WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-page-single-document/div/div[3]/div/app-signer-box/div/div[3]/div/div[2]/div/span")))

                    url_texto = url_download.text

                    if not url_texto:
                        raise Exception(f"URL de download não encontrada para o documento {contador}.")

                    driver.get(url_texto)
                    
                    time.sleep(8)
                    
                    #Botao download                    
                    WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/div/app-signers/div/div/app-sign-page/div/div/app-signed-document/div/div/app-box-download-and-create-my-account/div/div[2]/div/div/div[2]/zs-button/button/span/mat-icon"))).click()
                    
                    print(f"Documento de número {contador}, foi baixado com sucesso.")
                    documentos_baixados += 1
                        
                    time.sleep(3)
                    
                    #Volta para pagina de documentos da pasta termo premiação
                    driver.get("https://app.zapsign.com.br/conta/documentos?pasta=339a0489-c74a-48c0-b027-79c2a6571a25")
                    
                    #Clica no status de finalizados novamente, pois quando atualiza a pagina volta os pendentes no meio dos finalizados
                    WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-filter-status/div/div/div[1]"))).click()
                    
                    contador += 1                    
                    
                    time.sleep(2)
                else:
                    print(f"Documento de número {contador} não esta assinado.") 
                       
            except Exception as e:
                print("Erro no processo de download dos arquivos")
                traceback.print_exc
                sys.exit()

        if documentos_baixados != total_documentos:
            raise Exception(f"Faltaram {total_documentos - documentos_baixados} para baixar, encerrando o processo..")     
        return documentos_baixados       
                
    except Exception as e:
        print(f"Erro no processo de download: {e}")
        traceback.print_exc()
        raise

def mover_arquivos_baixados(download_dir, new_directory, documentos_baixados):
    tentativas_maximas = 3  # Defina o número máximo de tentativas
    tentativa = 0
    arquivos_movidos = []
    
    # Lista de arquivos já movidos (mantém o controle)
    arquivos_movidos_anteriores = set()

    while tentativa < tentativas_maximas:
        try:
            arquivos = [f for f in os.listdir(download_dir) if os.path.isfile(os.path.join(download_dir, f))]
            
            # Filtra os arquivos que ainda não foram movidos
            arquivos_restantes = [f for f in arquivos if f not in arquivos_movidos_anteriores]

            if not arquivos_restantes:
                raise Exception("Nenhum arquivo restante para mover.")

            for arquivo in arquivos_restantes:
                origem = os.path.join(download_dir, arquivo)
                destino = os.path.join(new_directory, arquivo)
                shutil.move(origem, destino)
                arquivos_movidos.append(arquivo)
                arquivos_movidos_anteriores.add(arquivo)  # Adiciona à lista de movidos
                print(f"Arquivo {arquivo} movido para {new_directory}")

            if not arquivos_movidos:
                raise Exception("Nenhum arquivo foi movido.")
            
            if len(arquivos_movidos) != documentos_baixados:
                raise Exception(f"Faltaram {documentos_baixados - len(arquivos_movidos)} documentos para mover, tentarei novamente..")
            
            return arquivos_movidos

        except Exception as e:
            print(f"Erro ao mover arquivos: {e}")
            tentativa += 1
            print(f"Tentativa {tentativa} falhou. Tentando novamente...")
            time.sleep(5)  # Aguarda antes de tentar novamente

    raise Exception(f"Falha ao mover arquivos após {tentativas_maximas} tentativas.")

def backup_finalizados(driver):
    
    try:
        # Clique em "selecionar tudo"
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-action-manager/div/div/div/div[1]/button/span"))).click()

        # Mover para
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "btn-open"))).click()

        # Procura pelo "Departamento Pessoal" nas pastas e clica na primeira opção encontrada
        for i in range(1, 20):            
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-action-manager/app-modal-folders/div/div/div[2]/app-folder-tree/div/div/div[2]/app-folder[{i}]/div")))

            departamento_pessoal = driver.find_element(By.XPATH, f"/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-action-manager/app-modal-folders/div/div/div[2]/app-folder-tree/div/div/div[2]/app-folder[{i}]/div")
            
            valor = departamento_pessoal.text
            
            if "Termo premiação Backup" in valor:
                #pasta termo premiação backup
                elemento = driver.find_element(By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-action-manager/app-modal-folders/div/div/div[2]/app-folder-tree/div/div/div[2]/app-folder[2]/div").click()
                
                # botão mover arquivos
                driver.find_element(By.XPATH, "/html/body/app-root/div/app-client/div/div/div/app-my-documents/app-documents/div/div/div[2]/div/div/app-action-manager/app-modal-folders/div/div/div[3]/zs-button[2]/button").click()
                
                print("Movendo arquivos para pasta backup..")
                time.sleep(20)
                break

    except Exception as e:
        print(f"Erro ao excluir documentos: {e}")
        traceback.print_exc()
        raise


def enviar_email(arquivos_movidos, new_directory):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'dti-admin@COMPANY_NAME.com.br'
    smtp_password = 'REMOVED_FOR_GITHUB'

    sender_email = 'dti-admin@COMPANY_NAME.com.br'
    receiver_email = 'israel.martins@COMPANY_NAME.com.br, raissa.correa@COMPANY_NAME.com.br, lucas.remor@COMPANY_NAME.com.br, nicolas.nasario@COMPANY_NAME.com.br'
    cc_emails = ['israel.martins@COMPANY_NAME.com.br']
    subject = 'RPA - Premiações: Arquivos Armazenados'

    corpo_texto = f"Prezados,\n\nOs arquivos abaixo foram armazenados no diretório: {new_directory}.\n\n"
    corpo_texto += "Arquivos:\n"
    for arquivo in arquivos_movidos:
        corpo_texto += f"- {arquivo}\n"
    corpo_texto += "\nAtenciosamente,\nEquipe RPA\n"

    mensagem = MIMEMultipart()
    mensagem['Subject'] = subject
    mensagem['From'] = sender_email
    mensagem['To'] = receiver_email
    mensagem['Cc'] = ", ".join(cc_emails)
    mensagem.attach(MIMEText(corpo_texto, 'plain'))

    try:
        servidor = smtplib.SMTP(smtp_server, smtp_port)
        servidor.starttls()
        servidor.login(smtp_username, smtp_password)
        servidor.sendmail(
            sender_email, [receiver_email] + cc_emails, mensagem.as_string())
        servidor.quit()
        print(f"E-mail enviado com sucesso para {receiver_email}.")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        traceback.print_exc()
        raise

if __name__ == "__main__":
    main()
