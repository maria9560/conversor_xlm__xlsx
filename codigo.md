# conversor_xml__xlsx
import xml.etree.ElementTree as ET
import pandas as pd

ns = {
    "ss": "urn:schemas-microsoft-com:office:spreadsheet"
}

tree = ET.parse("Base de corte 16.01.xml")
root = tree.getroot()

linhas = []

for row in root.findall(".//ss:Row", ns):
    linha = []
    col_atual = 1  # Excel começa na coluna 1

    for cell in row.findall("ss:Cell", ns):
        index = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")

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

# Define cabeçalho
df.columns = df.iloc[0]
df = df[1:].reset_index(drop=True)

df.to_excel("dados16.01.2026.xlsx", index=False)

print("Arquivo convertido com sucesso!")
