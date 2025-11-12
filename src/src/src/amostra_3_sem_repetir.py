import os
import truststore
import pandas as pd
import random

truststore.inject_into_ssl()

BASE_DIR = r"C:\Users\b8ms69\OneDrive - Linde Group\Documents\Projetos Ana\Projetos em ordem\Projeto 7 - Pesquisa nacional tratamento\Tratamento Ana"
PASTA_ENTRADAS = os.path.join(BASE_DIR, "Entradas")
PASTA_SAIDAS = os.path.join(BASE_DIR, "Saidas")

TABELAO = os.path.join(PASTA_ENTRADAS, "Tabelão_Final.xlsx")
AMOSTRA1 = os.path.join(PASTA_ENTRADAS, "Amostra 1.xlsx")
AMOSTRA2 = os.path.join(PASTA_ENTRADAS, "Amostra_2_SemRepetir.xlsx")
OUTPUT = os.path.join(PASTA_SAIDAS, "Amostra_3_SemRepetir.xlsx")

metas_origem = {
    "Large Bulk": 277,
    "Medicinal": 347,
    "Homecare": 310,
    "On Site": 141,
    "Packaged": 375,
    "Small Bulk": 308
}

print("Lendo planilhas de entrada...")
df = pd.read_excel(TABELAO)
amostra1 = pd.read_excel(AMOSTRA1)
amostra2 = pd.read_excel(AMOSTRA2)

col_id = "ID externo"

# Remove IDs já usados nas amostras anteriores
ids_usados = (
    pd.concat([amostra1[col_id], amostra2[col_id]])
    .drop_duplicates()
    .tolist()
)
df = df[~df[col_id].isin(ids_usados)].copy()

print(f"IDs removidos das Amostras 1 e 2: {len(ids_usados)}")
print(f"Base restante: {df[col_id].nunique()} IDs únicos disponíveis\n")

def selecionar_contatos(grupo):
    if grupo["Função"].str.contains("Comprador", case=False, na=False).any():
        return grupo[grupo["Função"].str.contains("Comprador", case=False, na=False)]
    return grupo.sample(1, random_state=42)

amostras = []

for origem, meta in metas_origem.items():
    subset = df[df["Origem"] == origem].copy()
    if subset.empty:
        continue
    total_ids = subset[col_id].nunique()
    ids_amostra = random.sample(list(subset[col_id].unique()), min(meta, total_ids))
    subset_filtrado = subset[subset[col_id].isin(ids_amostra)]
    amostra_final = subset_filtrado.groupby(col_id, group_keys=False).apply(selecionar_contatos)
    amostras.append(amostra_final)

amostra_3 = pd.concat(amostras, ignore_index=True)
amostra_3.to_excel(OUTPUT, index=False)

print("Amostra 3 gerada com sucesso.")
print(f"Caminho: {OUTPUT}")
print(f"IDs únicos sorteados: {amostra_3[col_id].nunique()}")
print(f"Linhas totais (contatos): {len(amostra_3)}")
