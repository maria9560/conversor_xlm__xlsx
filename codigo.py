import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

# Namespace do Excel XML
NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "urn:schemas-microsoft-com:office:spreadsheet"


def converter_xml_para_df(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    linhas_dict = {}
    max_colunas = 0
    contador_linhas = 1

    for row in root.findall(".//ss:Row", NS):
        linha = []
        col_atual = 1

        # Index real da linha
        row_index = row.get(f"{{{SS}}}Index")
        if row_index:
            row_index = int(row_index)
        else:
            row_index = contador_linhas

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

    # Reconstrói respeitando ordem real
    linhas_ordenadas = []
    for i in sorted(linhas_dict.keys()):
        linha = linhas_dict[i]
        while len(linha) < max_colunas:
            linha.append("")
        linhas_ordenadas.append(linha)

    df = pd.DataFrame(linhas_ordenadas, dtype=str)

    # Cabeçalho
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

        # Exibição pt-BR (sem quebrar tipo float)
       formatacao = {}if "Valor Faturas" in df.columns:
    formatacao["Valor Faturas"] = (
        "{:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )if "Quantidade Faturas" in df.columns:
    formatacao["Quantidade Faturas"] = (
        "{:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )if formatacao:
    st.dataframe(
        df.style.format(formatacao),
        use_container_width=True else:
    st.dataframe(df, use_container_width=True)

        # Download XLSX
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
