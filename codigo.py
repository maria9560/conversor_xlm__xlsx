import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

# Namespace do Excel XML
NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "urn:schemas-microsoft-com:office:spreadsheet"


def converter_xml_para_df(arquivo_xml):
    conteudo = arquivo_xml.read()
    root = ET.fromstring(conteudo)

    linhas_dict = {}
    max_colunas = 0
    contador_linhas = 1

    for row in root.findall(".//ss:Row", NS):
        linha = []
        col_atual = 1

        row_index = row.get(f"{{{SS}}}Index")
        row_index = int(row_index) if row_index else contador_linhas
        contador_linhas += 1

        for cell in row.findall("ss:Cell", NS):
            cell_index = cell.get(f"{{{SS}}}Index")

            if cell_index:
                cell_index = int(cell_index)
                while col_atual < cell_index:
                    linha.append("")
                    col_atual += 1

            data = cell.find("ss:Data", NS)
            valor = data.text if data is not None else ""
            linha.append(str(valor))
            col_atual += 1

        max_colunas = max(max_colunas, len(linha))
        linhas_dict[row_index] = linha

    linhas_ordenadas = []
    for i in sorted(linhas_dict.keys()):
        linha = linhas_dict[i]
        while len(linha) < max_colunas:
            linha.append("")
        linhas_ordenadas.append(linha)

    df = pd.DataFrame(linhas_ordenadas, dtype=str)
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)

    return df

def converter_colunas_float(df):
    colunas_float = ["Valor Faturas", "Quantidade Faturas"]

    for col in colunas_float:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
                .replace("", "0")
                .astype(float)
            )
    return df


    


def formatar_ptbr(valor):
    try:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return valor


def front():
    st.set_page_config(
        page_title="Conversor XML - XLSX",
        layout="wide"
    )

    st.title("Conversor XML - XLSX (fiel à base)")

    arquivo = st.file_uploader("Faça upload do arquivo XML", type=["xml"])

    if arquivo:
        with st.spinner("Convertendo arquivo..."):
            df = converter_xml_para_df(arquivo)
            df = converter_colunas_float(df)

        st.success("Conversão concluída com sucesso!")
        st.subheader("Pré-visualização dos dados")

        df_exibicao = df.copy()

        for col in df_exibicao.columns:
            if pd.api.types.is_numeric_dtype(df_exibicao[col]):
                df_exibicao[col] = df_exibicao[col].apply(formatar_ptbr)

        st.dataframe(df_exibicao, use_container_width=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        output.seek(0)

        st.download_button(
            label="Baixar arquivo XLSX",
            data=output,
            file_name="arquivo_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


front()
