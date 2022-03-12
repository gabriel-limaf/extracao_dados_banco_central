import requests
import csv
import pandas as pd
from openpyxl import load_workbook

mercado_mensais = 'ExpectativaMercadoMensais'
mercado_trimestrais = 'ExpectativasMercadoTrimestrais'
mercado_anuais = 'ExpectativasMercadoAnuais'
inflacao = 'ExpectativasMercadoInflacao12Meses'
top_mensal = 'ExpectativasMercadoTop5Mensais'
top_anual = 'ExpectativasMercadoTop5Anuais'
instituicao = 'ExpectativasMercadoInstituicoes'
format_csv = 'text/csv'
format_json = 'json'
format_html = 'text/html'
format_xml = 'xml'
output_file = 'arquivo.xlsx'
url = 'https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/' + top_anual + '?$format=' + format_csv
with requests.Session() as s:
    download = s.get(url)
    decoded_content = download.content.decode('utf-8')
    cr = csv.reader(decoded_content.splitlines(), delimiter=',')
    datas = pd.DataFrame(cr)
    book = load_workbook(output_file)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    datas.to_excel(writer, "Sheet1", header=True, index=False)
    writer.save()
