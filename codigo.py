# conversor_xml__xlsx
import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

# Namespace do Excel XML
ns = {
    "ss": "urn:schemas-microsoft-com:office:spreadsheet"
}

def converter_xml_para_df(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    linhas = []

    for row in root.findall(".//ss:Row", ns):
        linha = []
        col_atual = 1  # Excel começa na coluna 1

        for cell in row.findall("ss:Cell", ns):
            index = cell.get(
                "{urn:schemas-microsoft-com:office:spreadsheet}Index"
            )

            if index:
                index = int(index)
                while col_atual < index:
                    linha.append("")
                    col_atual += 1

            data = cell.find("ss:Data", ns)
            linha.append(data.text if data is not None else "")
            col_atual += 1

        linhas.append(linha)

    df = pd.DataFrame(linhas)

    # Primeira linha como cabeçalho
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    return df


def front():
    st.set_page_config(
        page_title="Conversor XML → XLSX",
        layout="wide"
    )

    st.title("Conversor de XML (Excel) para XLSX")

    arquivo = st.file_uploader(
        "Faça upload do arquivo XML",
        type=["xml"]
    )

    if arquivo is not None:
        with st.spinner("Convertendo arquivo..."):
            df = converter_xml_para_df(arquivo)

        st.success("Arquivo convertido com sucesso!")

        st.subheader("Pré-visualização dos dados")
        st.dataframe(df, use_container_width=True)

        # Salva em memória
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="⬇️ Baixar arquivo XLSX",
            data=output,
            file_name="arquivo_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


front()
