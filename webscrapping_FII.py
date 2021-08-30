# 01 - Importar as bibliotecas necessarias
import pandas as pd
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from urllib.error import URLError
from openpyxl import *

# 02 - Obter o HTML da pagina e separar a tabela em um dataframe
url = 'https://www.fundsexplorer.com.br/ranking'

try:
  req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
except HTTPError as e:
  print(e)
except URLError:
  print('Error')
else:
  web_byte = urlopen(req).read()
  webpage = web_byte.decode('utf-8')
  html = pd.read_html(webpage, match='ABCP11')
  ranking_table_df = html[0]

# 03 - Separar as colunas a serem exportadas
ranking_table_final_df = ranking_table_df [
  [
    'Códigodo fundo', 'Setor', 'Preço Atual', 
    'Liquidez Diária', 'Dividendo',
    'DividendYield', 'DY (12M)Acumulado',
    'PatrimônioLíq.','P/VPA',
    'VacânciaFísica'
  ]
]

# 04 - Exportar para um arquivo xlsx
ranking_table_final_df.to_excel('Ranking_table_FII.xlsx', index=False)