
# https://www.youtube.com/watch?v=VIY6oFPkSqU&t=42s

import pandas as pd
import requests
import tabula


urlPdf = " link"

fazendo = pdf_to_excel(urlPdf)

def pdf_to_excel (urlPdf):
    
    # Faz donwload
    download = requests.get(urlPdf)
    with open("temp.pdf", "wb") as f:
        f.write(download.content)
    
    # Converte o pdf em tabelas e vc pode escolher as páginas assim: pages = '1-5, 20-25, 30-50'
    tabelas = tabula.read_pdf("temp.pdf", pages='all')


    # Ao converter ele vai colocar cada página como uma aba do excel mas não queremos isso e sim em 
    # apenas uma Aba
    
    for tabela in tabelas:
        df = tabela.copy()
        dataframe_combinado = pd.concat


