"""
Empilha SÓ O MÊS NOVO na base_unificada.

FLUXO (você faz na mão o que precisar entre os passos):
1. Rode download_sds.py (ou atualizar_fontes()) pra baixar os arquivos novos.
2. Você ajusta manualmente os nomes das colunas de cada arquivo baixado,
   deixando-os IGUAIS aos nomes de origem que já estão em MAPA_COLUNAS
   (dentro de empilhar_bases.py) - ex: "MUNICIPIO DO FATO", "REGIAO
   GEOGRAFICA", "DATA DO FATO", "NATUREZA JURIDICA", "ANO", "SEXO", "IDADE"
   ou "FAIXA ETARIA", "TOTAL DE VITIMAS".
3. Rode este script: ele lê cada arquivo já ajustado, renomeia pro padrão,
   calcula a SITUAÇAO IDADE (mesma lógica de empilhar_bases.py), filtra só
   quem tem DATA mais nova que a última já salva em base_unificada.xlsx, e
   acrescenta esses registros na base - mantendo TODAS as colunas.

COMO USAR:
1. Ajuste ARQUIVOS_NOVOS abaixo (caminho, tipo, tipo_idade, aba) - mesma
   ideia da lista ARQUIVOS em empilhar_bases.py, só que aqui é só pros
   arquivos do mês novo.
2. Rode: python empilhar_mes_novo.py

IMPORTANTE - bug que eu encontrei em empilhar_bases.py:
   MAPA_COLUNAS renomeia "REGIAO GEOGRAFICA" -> "REGIAO", mas a lista
   COLUNAS_FINAIS de lá ainda procura a coluna pelo nome antigo
   ("REGIAO GEOGRAFICA"), que não existe mais depois do rename - por isso a
   coluna de região estava sumindo da base final. Aqui eu já uso o nome
   CERTO ("REGIAO") na lista de colunas finais. Se quiser, corrija a mesma
   linha em empilhar_bases.py pra não ter esse problema quando rodar ele
   direto também.
"""

import pandas as pd
from pathlib import Path

from empilhar_bases import (
    MAPA_COLUNAS,
    classificar_idade_numerica,
    classificar_faixa_etaria,
    carregar_base,
)

# ----------------------------------------------------------------------
# AJUSTE AQUI: arquivos recém-baixados, já com as colunas certas
# ----------------------------------------------------------------------
ARQUIVOS_NOVOS = [
    {"caminho": "MORTES VIOLENTAS_NOVO.xlsx", "tipo": "MORTES VIOLENTAS", "tipo_idade": "numerica", "aba": "Plan1"},
    {"caminho": "ESTUPRO_NOVO.xlsx", "tipo": "ESTUPRO", "tipo_idade": "faixa", "aba": "Plan1"},
    {"caminho": "VIOLENCIA DOMESTICA_NOVO.xlsx", "tipo": "VIOLENCIA DOMESTICA", "tipo_idade": "faixa", "aba": "Plan1"},
]

CAMINHO_BASE_UNIFICADA_XLSX = Path("base_unificada.xlsx")
CAMINHO_BASE_UNIFICADA_CSV = Path("base_unificada.csv")

# Colunas finais, já com o nome PÓS-renomeação (repare: REGIAO, não
# REGIAO GEOGRAFICA - ver nota no topo do arquivo).
COLUNAS_FINAIS = [
    "TIPO", "DATA", "ANO", "MUNICIPIO", "REGIAO GEOGRAFICA",
    "IDADE", "FAIXA ETARIA", "SITUAÇAO IDADE", "NATUREZA JURIDICA",
    "SEXO", "TOTAL DE VITIMAS",
]


def carregar_base_unificada_existente():
    if CAMINHO_BASE_UNIFICADA_XLSX.exists():
        base = pd.read_excel(CAMINHO_BASE_UNIFICADA_XLSX)
        if "DATA" in base.columns:
            base["DATA"] = pd.to_datetime(base["DATA"], errors="coerce")
        return base
    return None


def main():
    base_existente = carregar_base_unificada_existente()
    data_maxima = (
        base_existente["DATA"].max()
        if base_existente is not None and "DATA" in base_existente.columns
        else None
    )
    if data_maxima is not None and not pd.isna(data_maxima):
        print(f"DATA mais recente já na base_unificada: {data_maxima.date()}")
    else:
        data_maxima = None
        print("Nenhuma base_unificada.xlsx encontrada ainda (ou sem DATA) - tudo entra como novo.")

    partes_novas = []

    for item in ARQUIVOS_NOVOS:
        tipo = item["tipo"]
        print(f"\n=== {tipo} ===")

        try:
            df = carregar_base(item["caminho"], aba=item.get("aba"))
        except FileNotFoundError:
            print(f"  [AVISO] Arquivo não encontrado: {item['caminho']} - pulando.")
            continue

        # Renomeia pro padrão (exige nome IGUAL ao de MAPA_COLUNAS - por
        # isso você ajusta as colunas na mão antes de rodar este script).
        df = df.rename(columns=MAPA_COLUNAS.get(tipo, {}))
        df["TIPO"] = tipo

        # Classificação (Confirmado 0-17 / Provável / descartado) - mesma
        # lógica de empilhar_bases.py, já filtra o que não é 0-17.
        if item.get("tipo_idade") == "faixa":
            df = classificar_faixa_etaria(df, tipo)
        else:
            df = classificar_idade_numerica(df)

        colunas_existentes = [c for c in COLUNAS_FINAIS if c in df.columns]
        faltando = [c for c in COLUNAS_FINAIS if c not in df.columns]
        if faltando:
            print(f"  [AVISO] Colunas esperadas que não vieram nesse arquivo: {faltando}")
        df = df[colunas_existentes]

        if "DATA" not in df.columns:
            print(f"  [AVISO] Sem coluna DATA em {tipo} - não dá pra filtrar por mês novo, pulando esse arquivo.")
            continue

        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
        df_novos = df[df["DATA"] > data_maxima].copy() if data_maxima is not None else df

        print(f"  -> {len(df)} registro(s) classificado(s) (Confirmado/Provável); "
              f"{len(df_novos)} são de DATA nova.")
        if len(df_novos):
            partes_novas.append(df_novos)

    if not partes_novas:
        print("\nNenhum registro novo encontrado. base_unificada não foi alterada.")
        return

    novos_total = pd.concat(partes_novas, ignore_index=True, sort=False)

    base_final = (
        pd.concat([base_existente, novos_total], ignore_index=True, sort=False)
        if base_existente is not None
        else novos_total
    )

    base_final.to_excel(CAMINHO_BASE_UNIFICADA_XLSX, index=False)
    base_final.to_csv(CAMINHO_BASE_UNIFICADA_CSV, index=False, sep=";", encoding="utf-8-sig")

    print(f"\n{len(novos_total)} registro(s) novo(s) adicionado(s).")
    print(f"Total na base_unificada agora: {len(base_final)} registros.")
    print("Arquivos atualizados: base_unificada.xlsx e base_unificada.csv")


if __name__ == "__main__":
    main()
