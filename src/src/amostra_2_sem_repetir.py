import os
import truststore
truststore.inject_into_ssl()
import pandas as pd
import random

BASE_DIR = r"C:\Users\b8ms69\OneDrive - Linde Group\Documents\Projetos Ana\Projetos em ordem\Projeto 7 - Pesquisa nacional tratamento\Tratamento Ana"
PASTA_ENTRADAS = os.path.join(BASE_DIR, "Entradas")
PASTA_SAIDAS = os.path.join(BASE_DIR, "Saidas")

TABELAO = os.path.join(PASTA_ENTRADAS, "Tabelão_Final.xlsx")
AMOSTRA1 = os.path.join(PASTA_ENTRADAS, "Amostra 1.xlsx")
OUTPUT = os.path.join(PASTA_SAIDAS, "Amostra_2_SemRepetir.xlsx")

metas_origem = {
    "Large Bulk": 277,
    "Medicinal": 347,
    "Homecare": 310,
    "On Site": 141,
    "Packaged": 375,
    "Small Bulk": 308
}

print("Lendo planilhas...")
df = pd.read_excel(TABELAO)
amostra1 = pd.read_excel(AMOSTRA1)

col_id = "ID externo"
ids_usados = amostra1[col_id].drop_duplicates().tolist()
df = df[~df[col_id].isin(ids_usados)].copy()

print(f"Removidos {len(ids_usados)} IDs já sorteados.")
print(f"Base restante: {df[col_id].nunique()} IDs únicos.\n")

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

amostra_2 = pd.concat(amostras, ignore_index=True)
amostra_2.to_excel(OUTPUT, index=False)

print(f"Amostra 2 gerada com sucesso: {OUTPUT}")
print(f"IDs únicos sorteados: {amostra_2[col_id].nunique()}")
print(f"Linhas totais (contatos): {len(amostra_2)}")
