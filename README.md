## PLANTAO DELIVERY

**Autor:** Halisson S. Gomides
**Contato:** halisson.gomides@gmail.com


 > Processa a lista de presos do SISPEN e preenche o arquivo world (.docx) com o modelo de pesquisa diária de incidência penal, 
mandados a cumprir e antecedentes por crimes sexuais, conforme as rotas indicadas.

#### Requisitos:
- Python 3.8.5
- [Python Anaconda](https://anaconda.org/anaconda/python) ou [Miniconda](https://docs.conda.io/en/latest/miniconda.html) instalado
- [Chromedriver](https://chromedriver.chromium.org/): colocar o arquivo no caminho `C:\webdrivers\chromedriver_92.exe`
- As bibliotecas utilizadas, as quais constam especificadas no arquivo `environment.yml`

#### Passos necessários para rodar o programa:

1. Recriar o ambiente virtual python: `conda env create -f environment.yml`
2. Ativar o ambiente virtual python: `conda activate plantao_delivery`
2. Exemplo comando: `(plantao_delivery)$ python popula_modelo.py -c 1 -r sul,leste --pdf dados/originais/solicitacao_escolta_solicitacoes_dia07.pdf --model dados/originais/Pesquisa-Diaria-DCCP-modelo.docx` 
     Ao executar com sucesso, emite a seguinte mensagem ao final:
    > Execucao finalizada com sucesso!
    > Arquivo criado:    07.03.2021.docx
- O programa cria um arquivo em formato word (.docx) contendo a relação dos presos organizados por rota, já com o resultado da consulta ao BNMP (caso não encontre mandado em aberto, preenche o campo com 'NC'). No me do arquivo gerado segue o padrão: `dd.mm.aaaa.docx` -> [data em que o programa foi executado]
3. Para entender os parâmetros: `(plantao_delivery)$ python popula_modelo.py --help`

#### Observações importantes:
- Essa versão do programa não lida automaticamente com o captcha do portal BNMP. Ele aguarda 90 segundos para que um ser humano realize a validação do captcha, antes de iniciar as pesquisas.
- Essa versão do programa não realiza consulta no sistema PROCED. Ao invés disso, ele preenche o arquivo word (.docx) com **NÃO** nos campos **|Mandado de Prisão|**	**|Mandado de Localização Judicial|**	**|Crime Sexual|**