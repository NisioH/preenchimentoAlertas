import pandas as pd
from pdf2image import convert_from_path
from PIL import ImageDraw, ImageFont
import os
import textwrap
from datetime import datetime

ARQUIVO_DADOS = "dados.xlsx"
#FONTE_CAMINHO = "/usr/share/fonts/liberation-sans/LiberationSans-Bold.ttf" # Uso no Linux
FONTE_CAMINHO = "arialbd.ttf" # Uso no Windows
TAMANHO_FONTE = 24
POSICAO_DATA_HOJE = (1100, 2200)

posicoes_base = {
    "Nome": (660, 645),
    "CPF": (1380, 645),
    "DataOcorrencia": (250, 688),
}


def obter_data_extenso():
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    hoje = datetime.now()
    return f"Vilhena / RO, {hoje.day:02d} de {meses[hoje.month]} de {hoje.year}"


def processar_geral(df_processar, pdf_nome, dicionario_posicoes, sufixo_nome):
    if not os.path.exists(pdf_nome):
        print(f"Erro: Arquivo {pdf_nome} não encontrado!")
        return

    pages = convert_from_path(pdf_nome)

    try:
        fonte = ImageFont.truetype(FONTE_CAMINHO, TAMANHO_FONTE)
    except OSError:
        print(f"Aviso: Fonte {FONTE_CAMINHO} não encontrada. Usando padrão.")
        fonte = ImageFont.load_default()

    if not os.path.exists('Alertas_Gerados'):
        os.makedirs('Alertas_Gerados')

    for i, linha in df_processar.iterrows():
        imagem = pages[0].copy()
        draw = ImageDraw.Draw(imagem)

        for campo, posicao in dicionario_posicoes.items():
            if campo in df_processar.columns:
                valor_exibir = str(linha[campo])

                wrapper = textwrap.TextWrapper(width=110)
                texto_formatado = wrapper.fill(text=valor_exibir)

                draw.multiline_text(posicao, texto_formatado, font=fonte, fill=(0, 0, 0), spacing=5)

        draw.text(POSICAO_DATA_HOJE, obter_data_extenso(), font=fonte, fill=(0, 0, 0))

        nome_pessoa = str(linha['Nome']).replace(" ", "_").strip()
        caminho_salvar = f'Alertas_Gerados/Alerta_{sufixo_nome}_{nome_pessoa}.pdf'

        imagem.save(caminho_salvar, "PDF", resolution=100.0)
        print(f"[{i + 1}] PDF Gerado: {nome_pessoa}")


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
        print("Coluna 'DataEstouroInvertalo' não encontrada na planilha!")


if __name__ == "__main__":
    print("-" * 30)
    print("SISTEMA DE GERAÇÃO DE ALERTAS")
    print("-" * 30)
    escolha = input('Informe o Alerta:\n1 = 02 Horas\n2 = 04 Horas\n3 = Intervalo\n> ')

    if escolha == '1':
        alerta_02horas()
    elif escolha == '2':
        alerta_04horas()
    elif escolha == '3':
        alerta_intervalo()
    else:
        print("Opção inválida!")