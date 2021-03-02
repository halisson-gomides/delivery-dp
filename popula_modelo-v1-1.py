from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import tabula
import numpy as np
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
route_name = ['leste', 'oeste', 'sul']

# Lendo os argumentos da linha de comando
parser = argparse.ArgumentParser(description='Processa a lista de presos do SISPEN e preenche o arquivo world '
                                             'com o modelo de pesquisa diária de incidência penal, mandados a cumprir'
                                             'e antecedentes por crimes sexuais, conforme as rotas indicadas.')
parser.add_argument("-c", "--cabecalho",
                    help="Caso deseje inserir dados de cabecalho: Nome do Agente, Matricula e Equipe", default=False)
parser.add_argument("-r", "--rotas", help="Nomes das rotas do dia separadas por ',' no formato: rota1,rota2,etc...",
                    default='leste,oeste,sul')
parser.add_argument("--pdf", required=True, help="caminho relativo do arquivo pdf oriundo do SISPEN")
parser.add_argument("--model", required=True,
                    help="caminho relativo do arquivo modelo em world para composição das rotas")
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
        dfs[i]['delegacia'] = np.nan
        df_delegacias.append(dfs[i].iloc[:, 0].where(dfs[i].iloc[:, 0].str.contains('^Delegacia', na=False)).dropna())

        if len(df_delegacias[i]) > 0:

            pfa = verifica_presos_folha_anterior(dfs[i], df_delegacias[i])  # presos da folha anterior
            if pfa > -1:
                dfs[i].iloc[:pfa, 5] = df_delegacias[i - 1].iloc[-1].split(' : ')[1].strip()

            idx_del = df_delegacias[i].index.to_list()
            lista = iter(idx_del)
            primeiro = next(lista, 'fim')
            while True:
                next_val = next(lista, 'fim')
                if next_val == 'fim':
                    dfs[i].iloc[primeiro:, 5] = df_delegacias[i][primeiro].split(' : ')[1].strip()
                    break
                else:
                    dfs[i].iloc[primeiro:next_val, 5] = df_delegacias[i][primeiro].split(' : ')[1].strip()
                    primeiro = next_val
        else:
            dfs[i].iloc[:, 5] = df_delegacias[i - 1].iloc[-1].split(' : ')[1].strip()

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


def scrapy_bnmp(drv, nome_preso: str, nome_mae: str) -> str:
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

    linhas_proc = drv.find_elements_by_xpath(
        '//div[@class="ui-datatable-tablewrapper ng-star-inserted"]/table/tbody/child::tr')
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

# Criando coluna de DP's
dp_patherns = [
    (df_final['delegacia'].isin(["1a DP", "4a DP", "8a DP", "10a DP"]), 'p01'),
    (df_final['delegacia'].isin(["2a DP", "5a DP"]), 'p05'),
    (df_final['delegacia'] == '6a DP', 'p06'),
    (df_final['delegacia'].isin(["12a DP", "17a DP"]), 'p12'),
    (df_final['delegacia'].isin(["13a DP", "35a DP"]), 'p13'),
    (df_final['delegacia'].isin(["15a DP", "18a DP", "19a DP", "23a DP", "DEAM II"]), 'p15'),
    (df_final['delegacia'].isin(["16a DP", "31a DP"]), 'p16'),
    (df_final['delegacia'].isin(["14a DP", "20a DP", "33a DP"]), 'p20'),
    (df_final['delegacia'].isin(["11a DP", "21a DP"]), 'p21'),
    (df_final['delegacia'].isin(["26a DP", "32a DP"]), 'p26'),
    (df_final['delegacia'].isin(["27a DP", "29a DP"]), 'p27'),
    (df_final['delegacia'] == '30a DP', 'p30')
]

dp_criteria, dp_values = zip(*dp_patherns)
df_final['dp'] = np.select(dp_criteria, dp_values, None)

# Criando coluna de Rotas
rotas_patherns = [
    (df_final['dp'].isin(['p05', 'p06', 'p13', 'p16', 'p30']), 'leste'),
    (df_final['dp'].isin(['p12', 'p15', 'p21', 'p26']), 'oeste'),
    (df_final['dp'].isin(['p01', 'p20', 'p27']), 'sul')
]

rotas_criteria, rotas_values = zip(*rotas_patherns)
df_final['rota'] = np.select(rotas_criteria, rotas_values, None)

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
time.sleep(90)

# ---------------------------------
# Preenchendo o documento Word modelo com o conteudo dos nomes dos presos de cada DP,
# conforme as rotas indicadas

df_ = df_final[df_final['rota'].isin(route_name)]  # Filtra pelas rotas

for dp in np.unique(df_['dp'].values):

    df_filtrado = df_.loc[df_['dp'] == dp].reset_index()
    merge_content = []
    for i, row in df_filtrado.iterrows():
        dict_content = {
            str(row['dp']) + '_idx': '0' + str(i + 1) if (i + 1) < 10 else str(i + 1),
            str(row['dp']) + '_nome': row['nome_preso'],
            str(row['dp']) + '_ip': '',
            str(row['dp']) + '_mp': 'NÃO',
            str(row['dp']) + '_mlj': 'NÃO',
            str(row['dp']) + '_cs': 'NÃO',
            str(row['dp']) + '_bnmp': scrapy_bnmp(driver, row['nome_preso'], row['nome_mae'])
        }
        merge_content.append(dict_content)

    document.merge_rows(str(dp) + '_idx', merge_content)

driver.quit()
document.write('./{:%d.%m.%Y}.docx'.format(date.today()))
print('\nExecucao finalizada com sucesso!', end='\n')
print('Arquivo criado:    {:%d.%m.%Y}.docx'.format(date.today()), end='\n')
