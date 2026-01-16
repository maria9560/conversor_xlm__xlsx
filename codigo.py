import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

ns = {
    "ss": "urn:schemas-microsoft-com:office:spreadsheet"
}

def converter_xml_para_df(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    linhas = []

    for row in root.findall(".//ss:Row", ns):
        linha = []
        col_atual = 1

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

            # 🔒 NUNCA converter tipo
            valor = data.text if data is not None else ""
            linha.append(str(valor))

            col_atual += 1

        linhas.append(linha)

    # 🔒 FORÇA tudo como string
    df = pd.DataFrame(linhas, dtype=str)

    # Cabeçalho SEM alterar conteúdo
    df.columns = df.iloc[0].astype(str)
    df = df.iloc[1:].reset_index(drop=True)

    return df


def front():
    st.set_page_config(
        page_title="Conversor XML → XLSX (Seguro)",
        layout="wide"
    )

    st.title("Conversor XML → XLSX (sem alterar dados)")

    arquivo = st.file_uploader(
        "Faça upload do XML",
        type=["xml"]
    )

    if arquivo:
        df = converter_xml_para_df(arquivo)

        st.success("Conversão concluída sem alteração de dados")

        st.dataframe(df, use_container_width=True)

        output = BytesIO()

        # 🔒 Garante que o Excel receba tudo como TEXTO
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        output.seek(0)

        st.download_button(
            "⬇️ Baixar XLSX (dados preservados)",
            data=output,
            file_name="arquivo_preservado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

front()
