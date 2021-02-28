from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import tabula
from numpy import nan
import pandas as pd
import argparse
from os import path
import sys
# ----------------------------------
# Uso do Selenium para scrapy no BNMP
# ----------------------------------
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from fake_useragent import UserAgent
import time
# ---------------------------------



# ----------------------------------
# Tratando os argumentos da linha de comando
# ----------------------------------

# Setando as rotas do dia
route_name = ['leste','oeste','sul']

# Lendo os argumentos da linha de comando
parser = argparse.ArgumentParser(description='Processa a lista de presos do SISPEN e preenche o arquivo world '
                                             'com o modelo de pesquisa diária de incidência penal, mandados a cumprir'
                                             'e antecedentes por crimes sexuais, conforme as rotas indicadas.')
parser.add_argument("-c", "--cabecalho", help= "Caso deseje inserir dados de cabecalho: Nome do Agente, Matricula e Equipe", default=False)
parser.add_argument("-r", "--rotas", help= "Nomes das rotas do dia separadas por ',' no formato: rota1,rota2,etc...", default='leste,oeste,sul')
parser.add_argument("--pdf", required= True, help= "caminho relativo do arquivo pdf oriundo do SISPEN")
parser.add_argument("--model", required= True, help= "caminho relativo do arquivo modelo em world para composição das rotas")
args = parser.parse_args()

# Setando o arquivo world com o modelo
template = args.model
# template = 'dados/originais/Pesquisa-Diaria-DCCP-modelo02.docx'

# Setando o arquivo PDF
pdf_path = args.pdf
# pdf_path = 'dados/originais/solicitacao_escolta_solicitacoes_dia25.pdf'

if not (path.exists(template) and path.exists(pdf_path)):
    print('Erro: Arquivo PDF ou MODEL inexistentes. \nInforme o caminho relativo completo dos dois arquivos.', end='\n')
    sys.exit(1)

if args.rotas:
    # Setando as rotas do dia
    route_name = args.rotas.split(',')

# Setando as variaveis para o cabeçalho do modelo World
agente = 'Renata'
matricula = '590010'
delivery = '3'
data_plantao = '{:%d/%m/%Y}'.format(date.today())

if args.cabecalho:
    agente = input('Insira o nome do Agente: ')
    matricula = input('Insira a matricula do Agente: ')
    delivery = input('Insira o num. da equipe: ')


# ----------------------------------
# Funções de tratamento do arquivo PDF
# ----------------------------------
def verifica_presos_folha_anterior(df_folha_atual, df_delegacias):
    '''
    Funcao que verifica se na pagina atual do PDF exitem presos da delegacia da pagina anterior
    :param df_folha_atual: TabulaDataFrame da pagina atual
    :param df_delegacias: Series das delegacias da pagina atual
    :return: -1 ou index do totalizados de presos
    '''

    str_total_presos = df_folha_atual.iloc[:, 0].where(
        df_folha_atual.iloc[:, 0].str.contains('^Total de presos para escolta na Delegacia', na=False)).dropna()

    presos_folha_anterior = -1
    for idx_total in str_total_presos.index.tolist():
        if idx_total < df_delegacias.index.min():
            presos_folha_anterior = idx_total

    return presos_folha_anterior


def trata_df_pdf(dfs_list):
    '''
    Funcao para tratar uma lista de dataframes criada pela bibliteca tabula
    :param dfs_list: tabula DataFrame List
    :return: DataFrame List
    '''
    dfs = dfs_list
    df_delegacias = []

    for i in range(len(dfs)):

        dfs[i].columns = ['nome_preso', 'nome_mae', 'dt_nascimento', 'ocorrencia', 'dt_cadastro']
        dfs[i]['delegacia'] = nan
        df_delegacias.append(dfs[i].iloc[:,0].where(dfs[i].iloc[:, 0].str.contains('^Delegacia', na=False)).dropna())

        if len(df_delegacias[i]) > 0:

            pfa = verifica_presos_folha_anterior(dfs[i], df_delegacias[i])  # presos da folha anterior
            if pfa > -1:
                dfs[i].iloc[:pfa, 5] = df_delegacias[i-1].iloc[-1].split(' : ')[1]

            idx_del = df_delegacias[i].index.to_list()
            lista = iter(idx_del)
            primeiro = next(lista, 'fim')
            while True:
                next_val = next(lista, 'fim')
                if next_val == 'fim':
                    dfs[i].iloc[primeiro:, 5] = df_delegacias[i][primeiro].split(' : ')[1]
                    break
                else:
                    dfs[i].iloc[primeiro:next_val, 5] = df_delegacias[i][primeiro].split(' : ')[1]
                    primeiro = next_val
        else:
            dfs[i].iloc[:, 5] = df_delegacias[i-1].iloc[-1].split(' : ')[1]

        dfs[i] = dfs[i].loc[dfs[i]['nome_preso'] != 'Nome do Preso']
        dfs[i].dropna(inplace=True)

        # trantando as colunas
        dfs[i]['nome_preso'] = dfs[i]['nome_preso'].str.strip()

    return dfs


# ----------------------------------
# Funções para uso do Scrapy
# ----------------------------------
def wait_element(drv, expr, timeout=8, by_tag=By.ID, to_sleep=0):
    '''
    Função para controlar o tempo de espera de carregamento da página pelo bot
    :param drv: selenium web driver
    :param expr: expressão xpath utilizada para encontrar o elemento desejado
    :param timeout: tempo maximo em segundos que o Selenium ira aguardar para um elemento ser encontrado dado um criterio de busca expr
    :param by_tag: tipo da expressão (ID | XPATH)
    :param to_sleep: Tempo adicional de espera, em segundos
    :return: boolean
    '''
    try:
        element_present = EC.presence_of_element_located((by_tag, expr))
        WebDriverWait(drv, timeout).until(element_present)
    except TimeoutException:
        print("Timed out waiting for page to load")
        pass
        return False
    if to_sleep > 0:
        time.sleep(to_sleep)
    return True


def scrapy_bnmp(drv, nome_preso:str, nome_mae:str) -> str:
    '''
    Função para realizar a consulta ao BNMP
    :param drv: selenium web driver
    :param nome_preso: Nome do Preso
    :param nome_mae: Nome da Mãe do Preso
    :return: str
    '''

    wait_element(drv, '//app-version/span[@class="cssClass"]', by_tag=By.XPATH)
    btn_pesquisar = drv.find_element_by_xpath('//button[contains(@label,"Pesquisar")]')
    input_nomepessoa = drv.find_element_by_xpath('//input[@name="nomePessoa"]')
    input_nomemae = drv.find_element_by_xpath('//input[@name="nomeMae"]')
    nome_preso = nome_preso.split('(')[0].strip()
    input_nomepessoa.send_keys(nome_preso)
    input_nomemae.send_keys(nome_mae)
    btn_pesquisar.click()
    wait_element(drv, '//app-version/span[@class="cssClass"]', by_tag=By.XPATH, to_sleep=3)
    try:
        semresultado = drv.find_element_by_xpath('//app-sem-resultado[contains(@class, "ng-star-inserted")]')
        btn_voltar = drv.find_element_by_xpath('//button[contains(@label,"Voltar")]')
        btn_voltar.click()
        wait_element(drv, '//button[contains(@label,"Pesquisar")]', by_tag=By.XPATH)
        return 'NC'
    except NoSuchElementException:
        pass

    linhas_proc = drv.find_elements_by_xpath('//div[@class="ui-datatable-tablewrapper ng-star-inserted"]/table/tbody/child::tr')
    str_content = ''
    for linha in linhas_proc:
        nome = linha.find_element_by_xpath('.//td[2]/span[contains(@class, "ui-cell-data")]').text
        if nome.upper() == nome_preso:
            numero = linha.find_element_by_xpath('.//td[1]/span[contains(@class, "ui-cell-data")]').text
            orgao = linha.find_element_by_xpath('.//td[5]/span[contains(@class, "ui-cell-data")]').text
            if len(str_content) > 1: str_content += '\n'
            str_content += numero + '\n' + orgao
        else:
            continue

    input_nomepessoa.clear()
    input_nomemae.clear()
    return str_content if len(str_content) > 0 else 'NC'


# ---------------------------------
# Tratando o PDF
# ---------------------------------
tab_dfs = tabula.read_pdf(pdf_path, columns=[300, 500, 600, 700, 800, 900], guess=False, pages='all')
df_tratados = trata_df_pdf(tab_dfs)
df_final = pd.concat(df_tratados, ignore_index=True)

# ---------------------------------
# Tratando o World
# ---------------------------------
document = MailMerge(template)

# Setando os valores no cabeçalho do documento
document.merge(
    nome_agente=agente,
    matr_agente=matricula,
    equipe=delivery,
    date_doc=data_plantao)

# Setando o dicionario de rotas
rotas = {
    'leste': ['p05', 'p06', 'p13', 'p16', 'p30'],
    'oeste': ['p12', 'p15', 'p21', 'p26'],
    'sul': ['p01', 'p20', 'p27']
}

# Criando dicionario de listas de nomes de cada DP
conteudo = {}

# Criando as chaves com os nomes de cada DP
for p in range(1, 40):
    chave = 'p0' + str(p) if p < 10 else 'p' + str(p)
    conteudo[chave] = []

# Preenchendo cada chave com a lista dos nomes dos presos de cada DP
# conforme regra de negocio
lst_delegacias = df_final['delegacia'].unique().tolist()
for delegacia in lst_delegacias:

    if delegacia == '1a DP':
        conteudo['p01'] = conteudo['p01'] + df_final.query('delegacia == "1a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '4a DP':
        conteudo['p01'] = conteudo['p01'] + df_final.query('delegacia == "4a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '5a DP':
        conteudo['p05'] = conteudo['p05'] + df_final.query('delegacia == "5a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '6a DP':
        conteudo['p06'] = conteudo['p06'] + df_final.query('delegacia == "6a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '8a DP':
        conteudo['p01'] = conteudo['p01'] + df_final.query('delegacia == "8a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '10a DP':
        conteudo['p01'] = conteudo['p01'] + df_final.query('delegacia == "10a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '11a DP':
        conteudo['p21'] = conteudo['p21'] + df_final.query('delegacia == "11a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '12a DP':
        conteudo['p12'] = conteudo['p12'] + df_final.query('delegacia == "12a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '13a DP':
        conteudo['p13'] = conteudo['p13'] + df_final.query('delegacia == "13a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '14a DP':
        conteudo['p20'] = conteudo['p20'] + df_final.query('delegacia == "14a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '15a DP':
        conteudo['p15'] = conteudo['p15'] + df_final.query('delegacia == "15a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '16a DP':
        conteudo['p16'] = conteudo['p16'] + df_final.query('delegacia == "16a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '17a DP':
        conteudo['p12'] = conteudo['p12'] + df_final.query('delegacia == "17a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '18a DP':
        conteudo['p15'] = conteudo['p15'] + df_final.query('delegacia == "18a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '19a DP':
        conteudo['p15'] = conteudo['p15'] + df_final.query('delegacia == "19a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '20a DP':
        conteudo['p20'] = conteudo['p20'] + df_final.query('delegacia == "20a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '21a DP':
        conteudo['p21'] = conteudo['p21'] + df_final.query('delegacia == "21a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '23a DP':
        conteudo['p15'] = conteudo['p15'] + df_final.query('delegacia == "23a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '26a DP':
        conteudo['p26'] = conteudo['p26'] + df_final.query('delegacia == "26a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '27a DP':
        conteudo['p27'] = conteudo['p27'] + df_final.query('delegacia == "27a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '29a DP':
        conteudo['p27'] = conteudo['p27'] + df_final.query('delegacia == "29a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '30a DP':
        conteudo['p30'] = conteudo['p30'] + df_final.query('delegacia == "30a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '31a DP':
        conteudo['p16'] = conteudo['p16'] + df_final.query('delegacia == "31a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '32a DP':
        conteudo['p26'] = conteudo['p26'] + df_final.query('delegacia == "32a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '33a DP':
        conteudo['p20'] = conteudo['p20'] + df_final.query('delegacia == "33a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == '35a DP':
        conteudo['p13'] = conteudo['p13'] + df_final.query('delegacia == "35a DP"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

    elif delegacia == 'DEAM II':
        conteudo['p15'] = conteudo['p15'] + df_final.query('delegacia == "DEAM II"')[['nome_preso', 'nome_mae']].to_dict(orient='records')

# Excluindo as chaves que ficaram vazias
delete = [key for key in conteudo if len(conteudo[key]) == 0]
for key in delete: del conteudo[key]


# ---------------------------------
# Inciando consulta web ao BNMP
# ---------------------------------
options = Options()
ua = UserAgent()
userAgent = ua.random
options.add_argument(f'user-agent={userAgent}')
driver = webdriver.Chrome(options=options, executable_path="C:\\webdrivers\\chromedriver.exe")
url = 'https://portalbnmp.cnj.jus.br/#/pesquisa-peca'
driver.get(url)
time.sleep(80)


# ---------------------------------
# Preenchendo o documento Word modelo com o conteudo dos nomes dos presos de cada DP,
# conforme as rotas indicadas
for rota in route_name:

    for delegacia in rotas[rota]:
        merge_content = []
        if delegacia in conteudo:

            for i in range(len(conteudo[delegacia])):
                dict_content = {
                    str(delegacia) + '_idx': '0' + str(i + 1) if (i + 1) < 10 else str(i + 1),
                    str(delegacia) + '_nome': conteudo[delegacia][i]['nome_preso'],
                    str(delegacia) + '_ip': '',
                    str(delegacia) + '_mp': 'NÃO',
                    str(delegacia) + '_mlj': 'NÃO',
                    str(delegacia) + '_cs': 'NÃO',
                    str(delegacia) + '_bnmp': scrapy_bnmp(driver, conteudo[delegacia][i]['nome_preso'], conteudo[delegacia][i]['nome_mae'])
                }
                merge_content.append(dict_content)

            document.merge_rows(str(delegacia) + '_idx', merge_content)


driver.quit()
document.write('./{:%d.%m.%Y}.docx'.format(date.today()))
print('\nExecucao finalizada com sucesso!', end='\n')
print('Arquivo criado:    {:%d.%m.%Y}.docx'.format(date.today()), end='\n')