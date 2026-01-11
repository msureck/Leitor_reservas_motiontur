import streamlit as st
import tabula
import pandas as pd
from tabula import convert_into
import glob
import os
import re
import openpyxl
import tempfile
import io
import PyPDF2
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Confirmação de Reservas",
    page_icon="Motion.ico",
    layout="centered"
)

# Estilo customizado
st.markdown("""
    <style>
    .main {
        background-color: #0b2a4a;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 15px 30px;
        font-size: 16px;
        border-radius: 5px;
        border: none;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    h1 {
        color: white;
        text-align: center;
    }
    .uploadedFile {
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

# Título da aplicação
st.title("Leitor de Reservas")

# Função para extrair idades do PDF
def extrair_idades_do_pdf(caminho_pdf):
    """
    Extrai informações de idades de um arquivo PDF.
    Procura pelo padrão: Nome Data_Nascimento Idade Contato
    """
    pessoas = []
    
    try:
        with open(caminho_pdf, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            texto_completo = ""
            for pagina in pdf_reader.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"

            linhas = texto_completo.split("\n")
            palavras_ignorar = ['NOME', 'DATA', 'VOUCHER', 'EMISSÃO', 'CNPJ', 'BRASIL',
                                'CURITIBA', 'PARANÁ', 'MOTION', 'TURISMO', 'LOCAL']
            for i, linha in enumerate(linhas):
                # Procurar por padrão de data e idade na linha
                match = re.search(r"(\d{2}/\d{2}/\d{4})\s+(\d{1,3})", linha)
                if match:
                    data_nasc_str = match.group(1)
                    idade = int(match.group(2))
                    # Nome pode estar antes da data na mesma linha ou na linha anterior
                    nome = linha[:match.start()].strip()
                    if not nome and i > 0:
                        nome = linhas[i-1].strip()
                    # Validação básica
                    if 0 <= idade <= 120 and len(nome) > 3:
                        if not any(palavra in nome.upper() for palavra in palavras_ignorar):
                            pessoas.append({
                                'nome': nome,
                                'data_nascimento': data_nasc_str,
                                'idade': idade
                            })
    except Exception as e:
        pass
    return pessoas


def classificar_por_faixa_etaria(idade):
    """
    Classifica a idade em faixas:
    Criança (até 12), Adolescente (13-17), Jovem (18-25), Adulto (26-59), Idoso (60+)
    """
    if idade <= 12:
        return 'Criança'
    elif idade <= 17:
        return 'Adolescente'
    elif idade <= 25:
        return 'Jovem'
    elif idade <= 59:
        return 'Adulto'
    else:
        return 'Idoso'

# Função principal de processamento
def processar_pdfs(uploaded_files):
    
    extraction_area = [330.00, 0.00, 800.00, 600.00]
    extraction_area_valores = [550.00, 0.00, 800.00, 600.00]
    extraction_area_passeios = [340.00, 0.00, 800.00, 600.00]

    if not uploaded_files:
        st.error("⚠️ Nenhum arquivo PDF foi enviado!")
        return None, None

    # Lista para armazenar as informações encontradas
    resultados_valores = pd.DataFrame()
    resultados_passeios = pd.DataFrame()
    resultados_idades = []
    pessoas_passeios = []

    # Barra de progresso
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_files = len(uploaded_files)

    # Criar um diretório temporário para salvar os PDFs
    with tempfile.TemporaryDirectory() as temp_dir:
        
        for idx, uploaded_file in enumerate(uploaded_files):
            pdf_base_name = uploaded_file.name
            status_text.text(f"📄 Analisando: {pdf_base_name} ({idx + 1}/{total_files})")
            
            # Atualizar barra de progresso
            progress_bar.progress((idx + 1) / total_files)

            # Salvar o arquivo temporariamente
            temp_pdf_path = os.path.join(temp_dir, pdf_base_name)
            with open(temp_pdf_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            # Extrair informações de idade do PDF
            pessoas = extrair_idades_do_pdf(temp_pdf_path)
            # Tentar extrair passeios do DataFrame de passeios (se já existir)
            passeios_pdf = []
            try:
                # Tenta extrair passeios do arquivo PDF atual
                if 'df_passeios' in locals():
                    passeios_pdf = df_passeios['PASSEIO'].tolist() if 'PASSEIO' in df_passeios.columns else []
            except Exception:
                passeios_pdf = []
            for pessoa in pessoas:
                resultados_idades.append({
                    'Arquivo': pdf_base_name.replace('.pdf', ''),
                    'Nome': pessoa['nome'],
                    'Data de Nascimento': pessoa['data_nascimento'],
                    'Idade': pessoa['idade'],
                    'Classificação': classificar_por_faixa_etaria(pessoa['idade'])
                })
                # Adiciona uma linha para cada passeio encontrado, senão None
                if passeios_pdf:
                    for passeio in passeios_pdf:
                        pessoas_passeios.append({
                            'Arquivo': pdf_base_name.replace('.pdf', ''),
                            'Nome': pessoa['nome'],
                            'Idade': pessoa['idade'],
                            'Passeio': passeio
                        })
                else:
                    pessoas_passeios.append({
                        'Arquivo': pdf_base_name.replace('.pdf', ''),
                        'Nome': pessoa['nome'],
                        'Idade': pessoa['idade'],
                        'Passeio': None
                    })

            # Use Tabula to extract the text from the first page within the specified area
            pdf_text = tabula.read_pdf(temp_pdf_path, pages='1', area=extraction_area, output_format="json")

            # Define the text to search for
            valor_text = 'VALOR TOTAL'
            valor = None

            # Loop through the extracted JSON data to find the 'top' value where 'VALOR TOTAL' is found
            for item in pdf_text[0]['data']:
                for cell in item:
                    if 'text' in cell and re.search(valor_text, cell['text'], re.IGNORECASE):
                        valor = float(cell['top'])
                        break

            # Define the text to search for
            passeio_text = 'ROTEIRO DETALHADO'
            passeio = None

            # Loop through the extracted JSON data to find the 'top' value where 'ROTEIRO DETALHADO' is found
            for item in pdf_text[0]['data']:
                for cell in item:
                    if 'text' in cell and re.search(passeio_text, cell['text'], re.IGNORECASE):
                        passeio = float(cell['top'])
                        break

            if valor is not None:
                # Extraindo Valores
                extraction_area_valores = [valor, 0.00, (valor + 50.00), 600.00]

                df_valores = tabula.read_pdf(temp_pdf_path, pages=1, area=extraction_area_valores)[0]

                df_valores = df_valores.drop(columns=['Unnamed: 0'])
                df_valores = df_valores.drop(columns=['VALOR PAGO'])
                df_valores = df_valores.drop(columns=['SALDO'])
                
                # Renomear a coluna VALOR TOTAL para Valor por Voucher
                df_valores = df_valores.rename(columns={'VALOR TOTAL': 'Valor por Voucher'})

                try:
                    df_valores['Valor por Voucher'] = df_valores['Valor por Voucher'].str.replace('R$', '', regex=False)
                    df_valores['Valor por Voucher'] = df_valores['Valor por Voucher'].str.replace('.', '', regex=False)
                    df_valores['Valor por Voucher'] = df_valores['Valor por Voucher'].str.replace(',', '.', regex=False)
                except:
                    pass

                # Convertendo a coluna para tipo numérico
                df_valores['Valor por Voucher'] = pd.to_numeric(df_valores['Valor por Voucher'], errors='coerce')

                df_valores = df_valores.dropna()

                pdf_base_name_clean = pdf_base_name.replace('.pdf', '')

                # Adicione a coluna com o nome do PDF no início do DataFrame
                df_valores.insert(0, 'Nome Arquivo', pdf_base_name_clean)

                resultados_valores = pd.concat([resultados_valores, df_valores], ignore_index=True)

                # Extraindo Passeios
                extraction_area_passeios = [passeio, 0.00, (passeio + 50.00), 600.00]

                df_passeios = tabula.read_pdf(temp_pdf_path, pages=1, area=extraction_area_passeios)[0]

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

    df = pd.DataFrame(resultados_valores)
    
    # Adicionar linha de total
    if not df.empty and 'Valor por Voucher' in df.columns:
        # Calcular a soma antes de formatar
        total_valor = df['Valor por Voucher'].sum()
        
        # Criar linha vazia
        linha_vazia = pd.DataFrame([{col: '' for col in df.columns}])
        
        # Criar linha de total
        linha_total = pd.DataFrame([{col: '' for col in df.columns}])
        linha_total.loc[0, 'Nome Arquivo'] = 'Total:'
        linha_total.loc[0, 'Valor por Voucher'] = total_valor
        
        # Concatenar as linhas
        df = pd.concat([df, linha_vazia, linha_total], ignore_index=True)
        
        # Formatar a coluna Valor por Voucher para R$ 1.000,00
        df['Valor por Voucher'] = df['Valor por Voucher'].apply(
            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if pd.notna(x) and x != '' else x
        )

    status_text.empty()
    progress_bar.empty()

    # Criar dataframe de idades e resumo
    df_detalhes_idades = pd.DataFrame(resultados_idades)
    df_resumo_idades = None
    
    if not df_detalhes_idades.empty:
        todas_idades = df_detalhes_idades['Idade'].tolist()
        criancas = sum(1 for i in todas_idades if i <= 12)
        adolescentes = sum(1 for i in todas_idades if 13 <= i <= 17)
        jovens = sum(1 for i in todas_idades if 18 <= i <= 25)
        adultos = sum(1 for i in todas_idades if 26 <= i <= 59)
        idosos = sum(1 for i in todas_idades if i >= 60)

        total = len(todas_idades)
        df_resumo_idades = pd.DataFrame({
            'Faixa Etária': [
                'Criança (até 12 anos)',
                'Adolescente (13-17 anos)',
                'Jovem (18-25 anos)',
                'Adulto (26-59 anos)',
                'Idoso (60+ anos)',
                '',
                'TOTAL'
            ],
            'Quantidade': [
                criancas,
                adolescentes,
                jovens,
                adultos,
                idosos,
                '',
                total
            ],
            'Percentual': [
                f"{(criancas/total*100):.1f}%" if total > 0 else "0%",
                f"{(adolescentes/total*100):.1f}%" if total > 0 else "0%",
                f"{(jovens/total*100):.1f}%" if total > 0 else "0%",
                f"{(adultos/total*100):.1f}%" if total > 0 else "0%",
                f"{(idosos/total*100):.1f}%" if total > 0 else "0%",
                '',
                '100%'
            ]
        })

    df_pessoas_passeios = pd.DataFrame(pessoas_passeios)

    # Adicionar contagem de faixas etárias por passeio
    if not df_pessoas_passeios.empty and 'Passeio' in df_pessoas_passeios.columns:
        # Adiciona coluna de faixa etária
        df_pessoas_passeios['Faixa Etária'] = df_pessoas_passeios['Idade'].apply(classificar_por_faixa_etaria)
        # Pivot para contar por passeio e faixa
        pivot = pd.pivot_table(
            df_pessoas_passeios,
            index='Passeio',
            columns='Faixa Etária',
            values='Nome',
            aggfunc='count',
            fill_value=0
        ).reset_index()
        # Renomear colunas para garantir ordem e nomes
        for col in ['Criança', 'Adolescente', 'Jovem', 'Adulto', 'Idoso']:
            if col not in pivot.columns:
                pivot[col] = 0
        # Mesclar com df_resultado (df_passeios)
        if not df_resultado.empty and 'PASSEIO' in df_resultado.columns:
            df_resultado = df_resultado.merge(pivot[['Passeio','Criança','Adolescente','Jovem','Adulto','Idoso']],
                                              left_on='PASSEIO', right_on='Passeio', how='left')
            df_resultado.drop(columns=['Passeio'], inplace=True)
    return df, df_resultado, df_resumo_idades, df_detalhes_idades, df_pessoas_passeios


# Interface do usuário
st.markdown("---")

# Upload de arquivos
uploaded_files = st.file_uploader(
    "Envie os arquivos PDF de reserva:",
    type=['pdf'],
    accept_multiple_files=True,
    help="Selecione um ou mais arquivos PDF para processar"
)


# Botão para executar
if st.button("Executar Análise", disabled=not uploaded_files):
    if uploaded_files:
        with st.spinner("Processando arquivos..."):
            try:
                df_valores, df_passeios, df_resumo_idades, df_detalhes_idades, df_pessoas_passeios = processar_pdfs(uploaded_files)

                if df_valores is not None and df_passeios is not None:
                    # Criar o Excel em memória
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_valores.to_excel(writer, sheet_name='Sheet1', startcol=0, startrow=0, index=False)
                        df_passeios.to_excel(writer, sheet_name='Sheet1', startcol=3, startrow=0, index=False)
                        # Adicionar análise de idades a partir da coluna L (índice 11)
                        if df_resumo_idades is not None:
                            df_resumo_idades.to_excel(writer, sheet_name='Sheet1', startcol=11, startrow=0, index=False)
                        # Nova sheet com pessoas e passeios
                        if not df_pessoas_passeios.empty:
                            df_pessoas_passeios.to_excel(writer, sheet_name='PessoasPasseios', index=False)
                    excel_data = output.getvalue()

                    st.success("✅ Análise Concluída!")

                    # Mostrar preview dos dados
                    st.markdown("### Preview dos Valores")
                    st.dataframe(df_valores, use_container_width=True)

                    st.markdown("### Preview dos Passeios")
                    st.dataframe(df_passeios, use_container_width=True)

                    # Mostrar análise de idades
                    if df_resumo_idades is not None:
                        st.markdown("### Análise de Idades")
                        st.dataframe(df_resumo_idades, use_container_width=True)

                    # Preview sheet PessoasPasseios
                    if not df_pessoas_passeios.empty:
                        st.markdown("### Pessoas e Passeios")
                        st.dataframe(df_pessoas_passeios, use_container_width=True)

                    # Botão de download
                    st.download_button(
                        label="📥 Download do Excel",
                        data=excel_data,
                        file_name="Confirmacao_Reservas_Valores.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Erro ao processar os arquivos: {str(e)}")
    else:
        st.warning("⚠️ Por favor, envie pelo menos um arquivo PDF!")