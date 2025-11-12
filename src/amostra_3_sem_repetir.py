import os
import truststore
import pandas as pd
import numpy as np
from datetime import datetime

truststore.inject_into_ssl()

# Caminhos principais
base_dir = os.path.dirname(os.path.abspath(__file__))
entrada_dir = os.path.join(base_dir, "entradas")
saida_dir = os.path.join(base_dir, "saidas")
os.makedirs(saida_dir, exist_ok=True)

def gerar_amostra_sem_repetir(df, coluna_id, tamanho_amostra, semente=7):
    """
    Gera uma amostra aleatória sem repetir registros, garantindo unicidade pela coluna_id.
    """
    df_unico = df.drop_duplicates(subset=[coluna_id]).copy()
    amostra = df_unico.sample(n=tamanho_amostra, random_state=semente)
    return amostra

def main():
    print("\n=== Iniciando geração da Amostra 3 (sem repetir) ===")

    arquivos = [f for f in os.listdir(entrada_dir) if f.endswith(".xlsx") or f.endswith(".csv")]
    if not arquivos:
        print("Nenhum arquivo encontrado na pasta de entrada.")
        return

    # Leitura consolidada
    bases = []
    for arquivo in arquivos:
        caminho = os.path.join(entrada_dir, arquivo)
        print(f"Lendo arquivo: {arquivo}")

        if arquivo.endswith(".xlsx"):
            df = pd.read_excel(caminho, dtype=str)
        else:
            df = pd.read_csv(caminho, dtype=str, sep=";", encoding="utf-8")

        bases.append(df)

    base_final = pd.concat(bases, ignore_index=True)
    print(f"Total de registros originais: {len(base_final):,}")

    # Amostra 3 com tamanho maior e nova semente
    amostra = gerar_amostra_sem_repetir(base_final, coluna_id="ID", tamanho_amostra=1000, semente=11)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"amostra_3_sem_repetir_{timestamp}.xlsx"
    caminho_saida = os.path.join(saida_dir, nome_arquivo)

    amostra.to_excel(caminho_saida, index=False)
    print(f"Amostra gerada com sucesso: {nome_arquivo}")
    print("Registros finais:", len(amostra))

if __name__ == "__main__":
    main()
