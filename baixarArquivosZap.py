import os
import requests
from datetime import datetime, timedelta

# Constantes globais
API_TOKEN = 'REMOVED_FOR_GITHUB'
BASE_URL = 'https://api.zapsign.com.br/api/v1/docs/'
HEADERS = {
    'Authorization': f'Bearer {API_TOKEN}' 
}

# Consulta data de hoje para gerar uma pasta nova no caminho
def diretorio():
    caminho = r'G:\.shortcut-targets-by-id\0B7kvjs8r64dlWTBnWUozODFOeTg\Premiações\Termos assinados'
    data = datetime.now()
    mes_ano = data.strftime("%m-%y")  # Exemplo: "03-25" para março de 2025
    print(data)
    print(mes_ano)
    
    # Cria o caminho completo com a subpasta mês-ano
    caminho_completo = os.path.join(caminho, mes_ano)
    print(caminho_completo)
    
    if not os.path.exists(caminho_completo):
        os.makedirs(caminho_completo)
    
    return caminho_completo

# Função que busca uma página de documentos da API
def buscar_documentos(pagina):
    # Calcula a data de 30 dias atrás
    data_hoje = datetime.now()
    data_limite = data_hoje - timedelta(days=30)
    
    # Formata as datas no formato YYYY-MM-DD que a API espera
    created_from = data_limite.strftime("%Y-%m-%d")
    created_to = data_hoje.strftime("%Y-%m-%d")
    
    params = {
        'folder_path': '/Termo de premiação',  # Pasta que queremos acessar
        'status': 'signed',                     # Somente documentos assinados
        'page': pagina,                         # Número da página atual
        'created_from': created_from,           # Data inicial (30 dias atrás)
        'created_to': created_to                # Data final (hoje)
    }

    print(f'Buscando documentos criados entre {created_from} e {created_to}')
    
    resposta = requests.get(BASE_URL, headers=HEADERS, params=params)

    if resposta.status_code == 200:
        return resposta.json()
    else:
        print(f'Erro ao buscar página {pagina}: {resposta.status_code}')
        print(f'Resposta da API: {resposta.text}')
        return None

# Função auxiliar para converter data ISO para formato legível
def formatar_data(data_iso):
    """Converte data ISO 8601 para formato dd/mm/yyyy HH:MM:SS"""
    if data_iso:
        try:
            data_obj = datetime.fromisoformat(data_iso.replace('Z', '+00:00'))
            return data_obj.strftime("%d/%m/%Y %H:%M:%S")
        except:
            return data_iso
    return "Data não disponível"

# Função auxiliar para calcular dias desde a criação
def calcular_dias_desde_criacao(created_at):
    """Calcula quantos dias se passaram desde a criação do documento"""
    if created_at:
        try:
            data_criacao = datetime.fromisoformat(created_at.replace('Z', '+00:00'))
            data_atual = datetime.now(data_criacao.tzinfo)
            diferenca = data_atual - data_criacao
            return diferenca.days
        except:
            return None
    return None

# Função que realiza o download do arquivo e exclui o documento
def baixar_arquivo(nome, url, caminho_completo, doc_token, created_at, last_update_at):
    # Adiciona o mês e ano ao nome do arquivo, mas sem duplicar a extensão
    data = datetime.now()
    mes_ano = data.strftime("%m-%y")
    
    # Separa o nome do arquivo da extensão
    nome_base, extensao = os.path.splitext(nome)
    
    # Cria o nome completo com o mês/ano
    nome_completo = f"{nome_base}_{mes_ano}{extensao}"
    
    caminho_arquivo = os.path.join(caminho_completo, nome_completo)

    # Calcula quantos dias desde a criação
    dias_desde_criacao = calcular_dias_desde_criacao(created_at)
    
    # Verifica se o arquivo já foi baixado
    if os.path.exists(caminho_arquivo):
        print(f'Arquivo {nome_completo} já existe, pulando download.')
        print(f'Data de criação: {formatar_data(created_at)} ({dias_desde_criacao} dias atrás)')
        return
    
    # Exibe informações de data do documento
    print(f'Data de criação: {formatar_data(created_at)} ({dias_desde_criacao} dias atrás)')
    print(f'Última atualização: {formatar_data(last_update_at)}')
    
    # Baixa o arquivo PDF
    resposta = requests.get(url)
    if resposta.status_code == 200:
        with open(caminho_arquivo, 'wb') as f:
            f.write(resposta.content)
        print(f'Arquivo salvo: {nome_completo}')

        # Após o download, excluir o documento
        excluir_documento(doc_token)
    else:
        print(f'Erro ao baixar {nome_completo}: {resposta.status_code}')


# Função que faz a requisição DELETE para excluir o documento
def excluir_documento(doc_token):
    url_exclusao = f"{BASE_URL}{doc_token}/"
    resposta = requests.delete(url_exclusao, headers=HEADERS)

    if resposta.status_code == 200:
        print(f'Documento com token {doc_token} excluído com sucesso.')
    else:
        print(f'Erro ao excluir documento {doc_token}: {resposta.status_code}')

# Função para contar a quantidade de arquivos assinados na pasta
def contar_arquivos_assinados(pasta):
    # Calcula a data de 30 dias atrás
    data_hoje = datetime.now()
    data_limite = data_hoje - timedelta(days=30)
    
    params = {
        'folder_path': pasta,
        'status': 'signed',
        'page': 1,
        'created_from': data_limite.strftime("%Y-%m-%d"),
        'created_to': data_hoje.strftime("%Y-%m-%d")
    }

    resposta = requests.get(BASE_URL, headers=HEADERS, params=params)

    if resposta.status_code == 200:
        dados = resposta.json()
        total_arquivos = dados['count']
        print(f'Total de arquivos assinados na pasta "{pasta}" (últimos 30 dias): {total_arquivos}')
        return total_arquivos
    else:
        print(f'Erro ao contar documentos assinados: {resposta.status_code}')
        return None

# Função principal que executa o processo
def main():
    pasta = '/Termo de premiação'
    caminho_completo = diretorio()

    total_assinados = contar_arquivos_assinados(pasta)

    if total_assinados is not None:
        if total_assinados == 0:
            print('\nNenhum documento assinado encontrado nos últimos 30 dias.')
            return
            
        pagina = 1
        total_baixados = 0
        total_ja_existentes = 0
        
        while True:
            dados = buscar_documentos(pagina)
            if not dados:
                break

            print(f'\n{"="*70}')
            print(f'Página {pagina} - {len(dados["results"])} arquivos encontrados')
            print(f'{"="*70}')

            for doc in dados['results']:
                print(f"\n{'-'*70}")
                print(f"Verificando arquivo: {doc['name']}")
                nome_arquivo = doc['name']
                url_arquivo = doc['signed_file']
                created_at = doc.get('created_at')
                last_update_at = doc.get('last_update_at')

                # Conta se foi baixado ou já existia
                caminho_teste = os.path.join(caminho_completo, nome_arquivo)
                if os.path.exists(caminho_teste):
                    total_ja_existentes += 1
                else:
                    total_baixados += 1

                # Realiza o download se o arquivo não existir ainda
                baixar_arquivo(nome_arquivo, url_arquivo, caminho_completo, 
                             doc['token'], created_at, last_update_at)

            if not dados['next']:
                break

            pagina += 1
        
        print(f'\n{"="*70}')
        print(f'RESUMO:')
        print(f'Total de arquivos processados: {total_assinados}')
        print(f'Arquivos baixados: {total_baixados}')
        print(f'Arquivos que já existiam: {total_ja_existentes}')
        print(f'{"="*70}')

# Bloco que executa o script
if __name__ == '__main__':
    main()