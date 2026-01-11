# %%
import tabula
import pandas as pd
from tabula import convert_into
import glob
import os
import re
import openpyxl
import tkinter as tk
from tkinter import filedialog, Label, ttk
import subprocess
import threading
import sys

# %%
janela_fechada = False
def fechar_janela():
    global janela_fechada
    janela_fechada = True
    janela_reservas.destroy()
    sys.exit()

def definir_pasta():
    global pasta
    diretorio = filedialog.askdirectory()
    if diretorio:
        pasta = diretorio
    label_pasta.config(text=f'\{os.path.basename(pasta)}')

def codigo():
    t = threading.Thread(target=code)
    t.start()

def code():
    
    extraction_area = [330.00, 0.00, 800.00, 600.00]
    extraction_area_valores = [550.00, 0.00, 800.00, 600.00]
    extraction_area_passeios = [340.00, 0.00, 800.00, 600.00]

    pdf_files = (pasta + "\\*.pdf")

    pdf_files = glob.glob(pdf_files)

    # Lista para armazenar as informações encontradas
    resultados_valores = pd.DataFrame()
    resultados_passeios = pd.DataFrame()

    for pdf_file in pdf_files:
        pdf_base_name = os.path.basename(pdf_file)
        lbl_final.config(text="Análisando: " + pdf_base_name)
        # Define the extraction area

        # Use Tabula to extract the text from the first page within the specified area
        pdf_text = tabula.read_pdf(pdf_file, pages='1', area=extraction_area, output_format="json")

        # Define the text to search for
        valor_text = 'VALOR TOTAL'
        valor = None

        # Loop through the extracted JSON data to find the 'top' value where 'Canais de Atendimento:' is found
        for item in pdf_text[0]['data']:  # Access the 'data' key of the first item in the JSON list
            for cell in item:
                if 'text' in cell and re.search(valor_text, cell['text'], re.IGNORECASE):
                    valor = float(cell['top'])  # Set y2 to the 'top' value
                    break

        
        # Define the text to search for
        passeio_text = 'ROTEIRO DETALHADO'
        passeio = None

        # Loop through the extracted JSON data to find the 'top' value where 'Canais de Atendimento:' is found
        for item in pdf_text[0]['data']:  # Access the 'data' key of the first item in the JSON list
            for cell in item:
                if 'text' in cell and re.search(passeio_text, cell['text'], re.IGNORECASE):
                    passeio = float(cell['top'])  # Set y2 to the 'top' value
                    break

            
        if valor is not None:
            # Extraindo Valores
            extraction_area_valores = [valor, 0.00, (valor + 50.00), 600.00]

            df_valores = tabula.read_pdf(pdf_file, pages=1, area=extraction_area_valores)[0]

            df_valores = df_valores.drop(columns=['Unnamed: 0'])
            df_valores = df_valores.drop(columns=['VALOR PAGO'])
            df_valores = df_valores.drop(columns=['SALDO'])

            # df_valores['VALOR TOTAL'] = df_valores['VALOR TOTAL'].str.replace('R$', '', regex=False)
            # # Também é possível remover espaços em branco subsequentes, se houver
            # df_valores['VALOR TOTAL'] = df_valores['VALOR TOTAL'].str.strip()

            try:

                df_valores['VALOR TOTAL'] = df_valores['VALOR TOTAL'].str.replace('R$', '', regex=False)
                df_valores['VALOR TOTAL'] = df_valores['VALOR TOTAL'].str.replace('.', '', regex=False)  # removendo ponto
                df_valores['VALOR TOTAL'] = df_valores['VALOR TOTAL'].str.replace(',', '.', regex=False)  # substituindo vírgula por ponto
            except:
                pass

            # Convertendo a coluna para tipo numérico
            df_valores['VALOR TOTAL'] = pd.to_numeric(df_valores['VALOR TOTAL'], errors='coerce')

            df_valores = df_valores.dropna()

            pdf_base_name = os.path.basename(pdf_file).replace('.pdf', '')

            # Adicione a coluna com o nome do PDF no início do DataFrame
            df_valores.insert(0, 'Nome Arquivo', pdf_base_name)

            resultados_valores = pd.concat([resultados_valores, df_valores], ignore_index=True)

            # Extraindo Passeios
            extraction_area_passeios = [passeio, 0.00, (passeio + 50.00), 600.00]

            df_passeios = tabula.read_pdf(pdf_file, pages=1, area=extraction_area_passeios)[0]

            df_passeios = df_passeios.drop(columns=['Unnamed: 0'])
            df_passeios = df_passeios.drop(columns=['DATA'])
            df_passeios = df_passeios.drop(columns=['LINK (ROTEIRO DETALHADO)'])
            df_passeios = df_passeios.dropna()
            
            resultados_passeios = pd.concat([resultados_passeios, df_passeios], ignore_index=True)

    df_resultado = pd.DataFrame(resultados_passeios)

    # Contando as ocorrências de cada nome e armazenando em um dicionário
    contagem = df_resultado['PASSEIO'].value_counts().to_dict()

    # Adicionando a contagem como uma nova coluna no DataFrame
    df_resultado['Quantidade'] = df_resultado['PASSEIO'].map(contagem)

    # Removendo as linhas duplicadas mantendo apenas a primeira ocorrência de cada nome
    df_resultado.drop_duplicates(subset='PASSEIO', keep='first', inplace=True)

    # Exibindo o DataFrame resultante
    df_resultado

    df = pd.DataFrame(resultados_valores)

    nome_pasta = os.path.basename(pasta)
    nome_arquivo = nome_pasta + ' - Valores.xlsx'
    path_completo = os.path.join(pasta, nome_arquivo)

    with pd.ExcelWriter(path_completo, engine='openpyxl') as writer:
        
        # Salve o df na coluna A
        df.to_excel(writer, sheet_name='Sheet1', startcol=0, startrow=0, index=False)
        
        # Salve o df_resultado na coluna D (iniciando na coluna D, linha 1)
        df_resultado.to_excel(writer, sheet_name='Sheet1', startcol=3, startrow=0, index=False)

        lbl_final.config(text="Análise Concluída, Arquivo salvo na pasta: " + nome_pasta)


# %%
janela_reservas = tk.Tk()
janela_reservas.configure(bg='#0b2a4a')
janela_reservas.title("Confirmação de Reservas")
janela_reservas.geometry("400x250")  # Adjusted height to accommodate the progress bar
janela_reservas.protocol("WM_DELETE_WINDOW", fechar_janela)  # Define o tratamento de evento para o fechamento da janela

style = ttk.Style(janela_reservas)
style.configure('TFrame', background='#0b2a4a')  # Configura a cor de fundo para todos os frames ttk.

estilo = ttk.Style()
estilo.configure('Botao.TButton', padding=(30, 15), background="red", color='blue', borderwidth=5, font=("Calibri", 12,))

# Adicionar um campo de entrada para exibir o diretório selecionado

lbl_espaço = Label(janela_reservas, bg="#0b2a4a", font=("Calibri", 12,))
lbl_espaço.pack()

selecionar_pasta_button = ttk.Button(janela_reservas, text="Selecionar a Pasta do mês", style='Botao.TButton', command=definir_pasta)
selecionar_pasta_button.pack(padx=5, pady=5)
label_pasta = Label(janela_reservas, text='', bg="#0b2a4a",fg="white",  font=("Calibri", 12,))
label_pasta.pack(pady=5)

executar_button = ttk.Button(janela_reservas, text="Executar", style='Botao.TButton', command=codigo)
executar_button.pack(padx=10, pady=10)

lbl_final = Label(janela_reservas, bg="#0b2a4a",fg="white", font=("Calibri", 12,))
lbl_final.pack()

try: 
    janela_reservas.iconbitmap(r"_internal\Motion.ico")  # Substitua "icone.ico" pelo caminho do seu arquivo de ícone
except:
    pass

# Inicie o loop principal da interface gráfica
janela_reservas.mainloop()


