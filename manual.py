from bs4 import BeautifulSoup
import pandas as pd
import os

period = input("Digite o período (MM_YYYY): ")

# HTML
with open("input_data/nfe_fasouto.html", "r", encoding="utf-8") as file:
    html_content = file.read()

# Parsear o HTML
soup = BeautifulSoup(html_content, "html.parser")

# Encontrar todas as tabelas com a classe "toggle box"
tables = soup.find_all("table", class_="toggle box")

# Lista para armazenar os dados extraídos
data = []

# Para cada tabela, extrair os dados desejados
for table in tables:
    rows = table.find_all("tr")  # Encontrar todas as linhas da tabela
    for row in rows:
        # Extrair dados das colunas desejadas
        descricao = row.find("td", class_="fixo-prod-serv-descricao")
        quantidade = row.find("td", class_="fixo-prod-serv-qtd")
        valor = row.find("td", class_="fixo-prod-serv-vb")
        unidade = row.find("td", class_="fixo-prod-serv-uc")
        
        # Garantir que todas as colunas foram encontradas antes de adicionar ao dataset
        if descricao and quantidade and valor:
            data.append({
                "Descrição": descricao.get_text(strip=True),
                "Quantidade": quantidade.get_text(strip=True),
                "Unidade": unidade.get_text(strip=True),
                "Valor": valor.get_text(strip=True),
            })

# Criar o dataframe com os dados extraídos
df = pd.DataFrame(data)

df['Valor'] = df['Valor'].str.replace(',', '.')
df["Valor"] = df["Valor"].astype(float)

df['Quantidade'] = df['Quantidade'].str.replace(',', '.')
df["Quantidade"] = df["Quantidade"].astype(float)

# Exportar para um arquivo Excel
output_file = f"output_data/feira_{period}.xlsx"
df.to_excel(output_file, index=False)

print(f"Dados exportados para {output_file}")

### Parte 2 - Unificação dos arquivos Excel ###

# Caminho para a pasta com os arquivos Excel
folder_path = "output_data"
output_file = "output_data/base_de_dados.xlsx"

# Lista para armazenar os DataFrames
dataframes = []

# Percorrer todos os arquivos na pasta
for filename in os.listdir(folder_path):
    if filename.startswith("feira") and filename.endswith(".xlsx"):
        # Criar o rótulo do período a partir do nome do arquivo
        parts = filename.split("_")
        if len(parts) > 2:
            rotulo = parts[1] + "/" + parts[2].split(".")[0]
        else:
            rotulo = "Indefinido"

        # Caminho completo do arquivo
        file_path = os.path.join(folder_path, filename)

        # Ler o arquivo Excel
        try:
            df = pd.read_excel(file_path)
            
            # Adicionar a coluna "Período" ao DataFrame
            df["Período"] = rotulo
            
            # Adicionar o DataFrame à lista
            dataframes.append(df)
        except Exception as e:
            print(f"Erro ao processar o arquivo {filename}: {e}")

# Concatenar todos os DataFrames em um único DataFrame
if dataframes:
    base_de_dados = pd.concat(dataframes, ignore_index=True)
    
    # Salvar o DataFrame resultante em um único arquivo Excel
    base_de_dados.to_excel(output_file, index=False)
    print(f"Base de dados consolidada salva em {output_file}")

    # Salvar o DataFrame resultante em outro diretório
    base_de_dados.to_excel(r"C:\Users\Daniel\Desktop\Minha nuvem\base_de_dados.xlsx", index=False)
    print(f"Base de dados consolidada salva em Minha nuvem")
else:
    print("Nenhum arquivo foi processado.")