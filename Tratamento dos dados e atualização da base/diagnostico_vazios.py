"""
Script de DIAGNÓSTICO DE VAZIOS de uma base.

Mostra, para cada coluna: quantidade de valores vazios e o percentual que
isso representa do total de registros. Também detalha se os vazios de uma
coluna específica (ex: idade / faixa etária) estão concentrados em algum
grupo (ano, natureza do crime, município) ou espalhados aleatoriamente.

COMO USAR:
1. Troque ARQUIVO pelo caminho da base que quer analisar agora.
2. Ajuste COLUNA_PRINCIPAL para a coluna que mais te interessa (geralmente
   a de idade/faixa etária), e COLUNAS_CRUZAMENTO para as colunas que você
   quer usar pra checar se a ausência tem padrão (ano, natureza, município).
3. Rode: python diagnostico_vazios.py
"""

import pandas as pd
from pathlib import Path

# AJUSTE AQUI: caminho da base que você quer analisar agora
ARQUIVO = "VIOLENCIA DOMESTICA.xlsx"

ABA= "Plan1"

# AJUSTE AQUI: coluna que você mais quer investigar (idade, faixa etária, etc.)
# Deixe como "" se só quiser o panorama geral de vazios por coluna.
COLUNA_PRINCIPAL = "FAIXA ETARIA"

# AJUSTE AQUI: colunas pelas quais você quer cruzar a ausência da COLUNA_PRINCIPAL
COLUNAS_CRUZAMENTO = ["ANO", "NATUREZA JURIDICA", "MUNICIPIO DO FATO"]

# Valores que devem ser tratados como "vazio" além de NaN real
# (algumas bases usam texto tipo "(vazio)", "Desconhecido", "Não informado")
VALORES_VAZIOS_EXTRA = ["(vazio)", "vazio", "n/a", "na", "nan", ""]

def carregar_base(caminho: str, aba=None) -> pd.DataFrame:
    caminho = Path(caminho)
    if caminho.suffix.lower() in [".xlsx", ".xls"]:
        xls = pd.ExcelFile(caminho)
        if len(xls.sheet_names) > 1 and aba is None:
            print(f"[AVISO] Esse arquivo tem {len(xls.sheet_names)} abas: {xls.sheet_names}")
            print("        Usando a primeira aba por padrão "
                  f"('{xls.sheet_names[0]}'). Se não for a aba certa (com os\n"
                  "        microdados linha a linha), defina ABA no topo do script "
                  "e rode de novo.\n")
        aba_usar = aba if aba is not None else 0
        return pd.read_excel(caminho, sheet_name=aba_usar)
    elif caminho.suffix.lower() == ".csv":
        # AJUSTE AQUI o separador/encoding se a base não abrir certo
        return pd.read_csv(caminho, sep=";", encoding="latin1")
    raise ValueError(f"Formato não suportado: {caminho.suffix}")
 
 
def eh_vazio(serie: pd.Series) -> pd.Series:
    """Marca como vazio tanto NaN real quanto strings tipo '(vazio)', 'Desconhecido' etc."""
    valores = serie.astype(str).str.strip().str.lower()
    return serie.isna() | valores.isin(VALORES_VAZIOS_EXTRA)
 
 
def main():
    df = carregar_base(ARQUIVO, aba=ABA)
    total = len(df)
    print(f"Base: {ARQUIVO}")
    print(f"Total de registros: {total}\n")
 
    # ---- Panorama geral: vazios por coluna --------------------------
    print("=== Vazios por coluna ===")
    resultado = []
    for col in df.columns:
        n_vazios = int(eh_vazio(df[col]).sum())
        pct = (n_vazios / total) * 100 if total else 0
        resultado.append((col, n_vazios, pct))
 
    resultado.sort(key=lambda x: x[1], reverse=True)
    for col, n_vazios, pct in resultado:
        print(f"{col:35s} {n_vazios:8d} vazios  ({pct:5.2f}%)")
 
    # ---- Detalhamento da coluna principal ----------------------------
    if COLUNA_PRINCIPAL and COLUNA_PRINCIPAL in df.columns:
        print(f"\n=== Detalhamento de vazios em '{COLUNA_PRINCIPAL}' ===")
        mask_vazio = eh_vazio(df[COLUNA_PRINCIPAL])
        n_vazio_col = int(mask_vazio.sum())
        print(f"Total de vazios: {n_vazio_col} ({n_vazio_col / total * 100:.2f}% da base)\n")
 
        if n_vazio_col == 0:
            print("Nenhum vazio encontrado nessa coluna - nada a investigar.")
            return
 
        for col_cruz in COLUNAS_CRUZAMENTO:
            if col_cruz not in df.columns:
                print(f"(coluna de cruzamento '{col_cruz}' não encontrada na base - pulando)")
                continue
 
            print(f"--- Distribuição dos vazios por '{col_cruz}' ---")
            # % que cada grupo representa DENTRO dos vazios
            distrib_vazios = df.loc[mask_vazio, col_cruz].value_counts(dropna=False)
            pct_vazios = (distrib_vazios / n_vazio_col * 100).round(2)
 
            # % que cada grupo representa na base INTEIRA (pra comparar)
            distrib_geral = df[col_cruz].value_counts(dropna=False)
            pct_geral = (distrib_geral / total * 100).round(2)
 
            for categoria, qtd in distrib_vazios.items():
                pct_geral_cat = pct_geral.get(categoria, 0)
                print(f"  {str(categoria):40s} {qtd:6d} vazios "
                      f"({pct_vazios[categoria]}% dos vazios | "
                      f"grupo representa {pct_geral_cat}% da base toda)")
            print()
 
        print("Como ler: se o % 'dos vazios' de uma categoria for bem maior que\n"
              "o % que ela representa 'na base toda', é sinal de que a ausência\n"
              "não é aleatória - está concentrada ali (ex: um ano que não\n"
              "coletava essa informação). Se os dois percentuais forem parecidos\n"
              "em todas as categorias, a ausência está espalhada e excluir os\n"
              "vazios tende a ser seguro.")
 
 
if __name__ == "__main__":
    main()
 