import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "urn:schemas-microsoft-com:office:spreadsheet"

def converter_xml_para_df(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    linhas_dict = {}
    max_colunas = 0

    for row in root.findall(".//ss:Row", ns):
        linha = []
        col_atual = 1

        # 🔥 RESPEITA O INDEX DA LINHA
        row_index = row.get(f"{{{SS}}}Index")
        row_index = int(row_index) if row_index else None

        for cell in row.findall("ss:Cell", ns):
            cell_index = cell.get(f"{{{SS}}}Index")
            if cell_index:
                cell_index = int(cell_index)
                while col_atual < cell_index:
                    linha.append("")
                    col_atual += 1

            data = cell.find("ss:Data", ns)
            valor = data.text if data is not None else ""
            linha.append(str(valor))
            col_atual += 1

        max_colunas = max(max_colunas, len(linha))

        # Guarda pela posição REAL da linha
        linhas_dict[row_index if row_index else len(linhas_dict) + 1] = linha

    # 🔒 Reconstrói respeitando ordem REAL
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


def front():
    st.set_page_config(
        page_title="Conversor XML → XLSX (linhas corretas)",
        layout="wide"
    )

    st.title("Conversor XML → XLSX (ordem real preservada)")

    arquivo = st.file_uploader("Upload do XML", type=["xml"])

    if arquivo:
        df = converter_xml_para_df(arquivo)

        st.success("Conversão concluída — linhas e valores corretos")

        st.dataframe(df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        output.seek(0)

        st.download_button(
            "⬇️ Baixar XLSX fiel à base",
            data=output,
            file_name="arquivo_fiel_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

front()
