import os
import pandas as pd
import numpy as np
import truststore

truststore.inject_into_ssl()

BASE_DIR = r"C:\Users\b8ms69\OneDrive - Linde Group\Documents\Projetos Ana\Projetos em ordem\Projeto 7 - Pesquisa nacional tratamento\Tratamento Ana"
INPUT = os.path.join(BASE_DIR, "Entradas", "Tabelão_Final.xlsx")
OUTPUT = os.path.join(BASE_DIR, "Saidas", "Amostra_Final_OrigemFixa.xlsx")

np.random.seed(42)

META_ORIGEM = {
    "Large Bulk": 277,
    "Medicinal": 347,
    "Homecare": 310,
    "On Site": 141,
    "Packaged": 375,
    "Small Bulk": 308
}

print("Lendo planilha...")
df = pd.read_excel(INPUT)
df.columns = [c.strip() for c in df.columns]

expected = [
    "ID externo",
    "Contato",
    "E-mail",
    "Telefone",
    "Celular",
    "Função",
    "Conta",
    "Segmento",
    "Região",
    "Origem"
]

if not all(col in df.columns for col in expected):
    print("Colunas esperadas não encontradas. Verifique a estrutura da planilha.")
    raise SystemExit(1)

df = df[expected].copy()
df["Origem"] = df["Origem"].astype(str).str.strip()
df["Região"] = df["Região"].astype(str).str.strip()
df["Função"] = df["Função"].astype(str).str.strip()

avail = (
    df.groupby(["Origem", "Região"])["ID externo"]
    .nunique()
    .reset_index(name="Disponível")
)
pivot_avail = avail.pivot(index="Origem", columns="Região", values="Disponível").fillna(0).astype(int)

print("IDs disponíveis por Origem x Região:")
print(pivot_avail)

alloc = pivot_avail.copy().astype(float)
for origem, meta in META_ORIGEM.items():
    if origem not in pivot_avail.index:
        continue
    total_disp = pivot_avail.loc[origem].sum()
    if total_disp == 0:
        continue
    proporcoes = pivot_avail.loc[origem] / total_disp
    alocadas = np.floor(proporcoes * meta)
    resto = meta - alocadas.sum()
    decimais = (proporcoes * meta) - alocadas
    while resto > 0:
        idx = decimais.idxmax()
        alocadas[idx] += 1
        decimais[idx] = 0
        resto -= 1
    alloc.loc[origem] = alocadas

alloc = alloc.astype(int)

print("Alocação final por Origem x Região:")
print(alloc)
print("Total geral:", alloc.values.sum())

selected_ids = []
for origem in alloc.index:
    for regiao in alloc.columns:
        n = alloc.loc[origem, regiao]
        if n <= 0:
            continue
        ids_pool = df[(df["Origem"] == origem) & (df["Região"] == regiao)]["ID externo"].drop_duplicates().tolist()
        if not ids_pool:
            continue
        n = min(n, len(ids_pool))
        sorteados = np.random.choice(ids_pool, size=n, replace=False)
        selected_ids.extend(sorteados)

selected_ids = list(dict.fromkeys(selected_ids))
print(f"Total de IDs sorteados: {len(selected_ids)}")

df_sel = df[df["ID externo"].isin(selected_ids)].copy()
final_rows = []
for idv, bloco in df_sel.groupby("ID externo"):
    compradores = bloco[bloco["Função"].str.contains("comprador", case=False, na=False)]
    if not compradores.empty:
        final_rows.append(compradores)
    else:
        final_rows.append(bloco.sample(n=1, random_state=42))

df_final = pd.concat(final_rows, ignore_index=True).drop_duplicates()

check_origem = df_final.groupby("Origem")["ID externo"].nunique().reset_index(name="IDs")
check_regiao = df_final.groupby("Região")["ID externo"].nunique().reset_index(name="IDs")
check_cross = df_final.groupby(["Origem", "Região"])["ID externo"].nunique().reset_index(name="IDs")

with pd.ExcelWriter(OUTPUT, engine="xlsxwriter") as writer:
    df_final.to_excel(writer, sheet_name="Amostra Final", index=False)
    alloc.reset_index().melt(id_vars="Origem", var_name="Região", value_name="IDs_meta").to_excel(writer, sheet_name="Meta_OrigemFixa", index=False)
    check_cross.to_excel(writer, sheet_name="Check_OrigemxRegiao", index=False)
    check_origem.to_excel(writer, sheet_name="Check_Origem", index=False)
    check_regiao.to_excel(writer, sheet_name="Check_Regiao", index=False)

print("Arquivo gerado com sucesso.")
print(f"Caminho: {OUTPUT}")
print(f"IDs únicos sorteados: {df_final['ID externo'].nunique()}")
print(f"Linhas totais (contatos): {len(df_final)}")
