import pandas as pd
from pdf2image import convert_from_path
from PIL import ImageDraw, ImageFont
import os
import sys
import textwrap
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if sys.platform.startswith('win'):
    CAMINHO_POPPLER = os.path.join(BASE_DIR, 'poppler-windows', 'Library', 'bin')
    FONTE_CAMINHO = os.path.join(BASE_DIR, 'arialbd.ttf') 
else:
    # No seu Fedora:
    CAMINHO_POPPLER = None 
    FONTE_CAMINHO = "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf"

ARQUIVO_DADOS = os.path.join(BASE_DIR, "dados.xlsx")
TAMANHO_FONTE = 24
POSICAO_DATA_HOJE = (1100, 2200)

# Dicionários de Posições
posicoes_base = {
    "Nome": (660, 645),
    "CPF": (1380, 645),
    "DataOcorrencia": (250, 688),
}

def obter_data_extenso():
    meses = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
             7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
    hoje = datetime.now()
    return f"Vilhena / RO, {hoje.day:02d} de {meses[hoje.month]} de {hoje.year}"

def processar_geral(df_processar, pdf_nome, dicionario_posicoes, sufixo_nome):
    caminho_pdf = os.path.join(BASE_DIR, pdf_nome)
    if not os.path.exists(caminho_pdf):
        print(f"Erro: Arquivo {pdf_nome} não encontrado!")
        return

    try:
        # Passa o caminho exato para o binário do Poppler no Windows
        pages = convert_from_path(caminho_pdf, poppler_path=CAMINHO_POPPLER)
    except Exception as e:
        print(f"Erro ao converter PDF: {e}")
        print(f"Verifique se o Poppler está em: {CAMINHO_POPPLER}")
        return
    
    try:
        fonte = ImageFont.truetype(FONTE_CAMINHO, TAMANHO_FONTE)
    except OSError:
        fonte = ImageFont.load_default()

    pasta_saida = os.path.join(BASE_DIR, 'Alertas_Gerados')
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    for i, linha in df_processar.iterrows():
        imagem = pages[0].copy()
        draw = ImageDraw.Draw(imagem)

        for campo, posicao in dicionario_posicoes.items():
            if campo in df_processar.columns:
                valor_exibir = str(linha[campo])
                wrapper = textwrap.TextWrapper(width=70)
                texto_formatado = wrapper.fill(text=valor_exibir)
                draw.multiline_text(posicao, texto_formatado, font=fonte, fill=(0, 0, 0), spacing=5)

        draw.text(POSICAO_DATA_HOJE, obter_data_extenso(), font=fonte, fill=(0, 0, 0))

        nome_pessoa = str(linha['Nome']).replace(" ", "_").strip()
        caminho_salvar = os.path.join(pasta_saida, f'Alerta_{sufixo_nome}_{nome_pessoa}.pdf')
        imagem.save(caminho_salvar, "PDF", resolution=100.0)
        print(f"[{i + 1}] Sucesso: {nome_pessoa}")

def carregar_e_padronizar():
    df = pd.read_excel(ARQUIVO_DADOS)
    df.columns = df.columns.str.strip()
    df['Nome'] = df['Nome'].str.strip().str.upper()
    for col in ['DataOcorrencia', 'DataEstouroInvertalo']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')
    return df

def alerta_02horas():
    df = carregar_e_padronizar()
    processar_geral(df, 'AlertaEducativo02horas.pdf', posicoes_base, "02Horas")

def alerta_04horas():
    df = carregar_e_padronizar()
    processar_geral(df, 'AlertaEducativo04horas.pdf', posicoes_base, "04Horas")

def alerta_intervalo():
    df = carregar_e_padronizar()
    if 'DataEstouroInvertalo' in df.columns:
        df_agrupado = df.dropna(subset=['DataEstouroInvertalo']).groupby(['Nome', 'CPF'])['DataEstouroInvertalo'].apply(
            lambda x: ', '.join(x)
        ).reset_index()
        
        posicoes_intervalo = {
            "Nome": (660, 645),
            "CPF": (1380, 645),
            "DataEstouroInvertalo": (300, 720),
        }
        processar_geral(df_agrupado, 'AlertaEducativoIntervalo.pdf', posicoes_intervalo, "Intervalo")
    else:
        print("Coluna 'DataEstouroInvertalo' não encontrada!")

if __name__ == "__main__":
    escolha = input('Informe o Alerta:\n1 = 02 Horas\n2 = 04 Horas\n3 = Intervalo\n> ')
    if escolha == '1': alerta_02horas()
    elif escolha == '2': alerta_04horas()
    elif escolha == '3': alerta_intervalo()
    else: print("Opção inválida.")