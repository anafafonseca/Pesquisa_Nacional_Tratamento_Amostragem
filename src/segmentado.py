import os
import truststore
truststore.inject_into_ssl()

import pandas as pd
import numpy as np

def gerar_amostra_segmentada(caminho_entrada, caminho_saida):
    """
    Gera a primeira amostra proporcional por origem e região.
    Aplica alocação proporcional, sorteio aleatório e priorização de compradores.
    """

    # Leitura da base consolidada
    base = pd.read_excel(caminho_entrada)

    # Exemplo de campos esperados
    if not {"Origem", "Região", "Comprador", "ID"}.issubset(base.columns):
        raise ValueError("As colunas obrigatórias não foram encontradas na base de entrada.")

    # Definição de proporção alvo por origem
    proporcoes = (
        base.groupby("Origem")["ID"]
        .count()
        .reset_index(name="qtde")
    )
    proporcoes["proporcao"] = proporcoes["qtde"] / proporcoes["qtde"].sum()

    # Amostra proporcional
    amostras = []
    for _, linha in proporcoes.iterrows():
        origem = linha["Origem"]
        proporcao = linha["proporcao"]

        subset = base[base["Origem"] == origem]
        n = int(len(base) * proporcao)

        # Sorteio com prioridade pra compradores
        compradores = subset[subset["Comprador"] == "Sim"]
        nao_compradores = subset[subset["Comprador"] != "Sim"]

        amostra = pd.concat([
            compradores.sample(frac=0.7, random_state=42, replace=False),
            nao_compradores.sample(frac=0.3, random_state=42, replace=False)
        ]).sample(frac=1, random_state=42)

        amostras.append(amostra.head(n))

    amostra_final = pd.concat(amostras).reset_index(drop=True)

    # Exportação
    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    amostra_final.to_excel(caminho_saida, index=False)

    print(f"Amostra segmentada gerada com sucesso: {caminho_saida}")
    return amostra_final


if __name__ == "__main__":
    entrada = r"C:\Users\b8ms69\OneDrive - Linde Group\Documents\Projetos Ana\Projetos em ordem\Projeto 7 - Pesquisa nacional tratamento\Tratamento Ana\Saidas\Tabelão_Final.xlsx"
    saida = r"C:\Users\b8ms69\OneDrive - Linde Group\Documents\Projetos Ana\Projetos em ordem\Projeto 7 - Pesquisa nacional tratamento\Tratamento Ana\Saidas\Amostra_Final_OrigemFixa.xlsx"
    gerar_amostra_segmentada(entrada, saida)
