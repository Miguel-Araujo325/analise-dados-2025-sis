import pandas as pd
import unicodedata
import chardet
from datetime import datetime

#NÃO ESQUECER DE ATUALIZAR O CAMINHO DO CSV E DO OUTPUT DO XLSX

# Função para remover acentos
def remover_acentos(texto):
    if isinstance(texto, str):
        texto_normalizado = unicodedata.normalize('NFD', texto)
        return ''.join([c for c in texto_normalizado if not unicodedata.combining(c)])
    return texto

# Caminho do arquivo CSV de entrada
file_path = r"C:\substituir_pelo_seu_caminho\detalhe_votacao.csv"

# Detectar o encoding do arquivo
with open(file_path, "rb") as f:
    result = chardet.detect(f.read(100000))
encoding_detectado = result["encoding"]
print(f"Encoding detectado: {encoding_detectado}")

# Ler o CSV com encoding correto
df = pd.read_csv(file_path, sep=";", encoding=encoding_detectado, dtype=str)

# --- Métricas de limpeza ---
registros_com_acentos = 0
espacos_removidos = 0

# Limpeza e padronização dos dados
for col in df.columns:
    # Contar registros com acentuação antes da limpeza
    registros_com_acentos += df[col].astype(str).str.contains(r"[À-ÿ]", regex=True).sum()
    
    # Contar espaços extras antes da limpeza
    espacos_removidos += df[col].astype(str).str.contains(r"^\s+|\s+$|\s{2,}", regex=True).sum()
    
    # Aplicar tratamentos
    df[col] = df[col].astype(str).str.strip()    # remover espaços extras
    df[col] = df[col].apply(remover_acentos)     # remover acentuação
    df[col] = df[col].str.upper()                # converter para maiúsculas

# Remover duplicatas
df = df.drop_duplicates()

# Preencher valores ausentes
df = df.fillna("DESCONHECIDO")

# Contar quantidade de "DESCONHECIDO"
qtd_desconhecido = (df == "DESCONHECIDO").sum().sum()

# Nome do arquivo de saída com padrão YYYYmmDD-HHMMSS
timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
output_path = fr"C:\substituir_pelo_seu_caminho\detalhe_votacao_{timestamp}.xlsx"

# Exportar para Excel
df.to_excel(output_path, index=False)

# --- Relatório Sysout ---
print("\n==== RELATORIO DE TRATAMENTO ====")
print(f"QNT DE REGISTROS COM ACENTUACAO REMOVIDA: {registros_com_acentos}")
print(f"QNT DE REGISTROS COM ESPACOS EXTRAS REMOVIDOS: {espacos_removidos}")
print(f"QNT DE REGISTROS 'DESCONHECIDO': {qtd_desconhecido}")
print(f"Arquivo tratado salvo em: {output_path}")
