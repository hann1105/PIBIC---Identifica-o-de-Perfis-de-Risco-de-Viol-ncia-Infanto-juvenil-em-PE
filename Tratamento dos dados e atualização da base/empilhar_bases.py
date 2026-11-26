"""
Script para EMPILHAR (concatenar) as bases:
- Mortes Violentas Intencionais (idade numérica)
- Estupro (faixa etária categórica)
- Violência Doméstica (faixa etária categórica)

O script já calcula sozinho a coluna SITUAÇAO IDADE para cada registro,
usando o critério que já validamos nas conversas anteriores:

- "Confirmado 0-17"        -> idade/faixa etária dentro do recorte infanto-juvenil
- "Provável vulnerável..." -> faixa etária vazia, mas a Natureza Jurídica tem
                               vínculo legal com vítima menor de idade
                               (Art. 217-A ou Art. 136 do Código Penal)
- (linha é DESCARTADA)      -> fora da faixa 0-17 e sem nenhum indício de
                               menoridade (vazio "comum", sem vínculo legal)

Só entram na base final os registros "Confirmado" e "Provável" - o resto
não é empilhado (mas continua intacto no arquivo original, nada é apagado
da planilha-fonte).

COMO USAR:
1. Ajuste ARQUIVOS: caminho, tipo, tipo_idade ("numerica" ou "faixa") e aba
   (nome da aba com os microdados, ou None se só tiver uma aba).
2. Ajuste MAPA_COLUNAS de cada base (nome original -> nome padrão).
3. Ajuste FAIXAS_INFANTO_JUVENIL: rótulos de faixa etária que valem como
   0-17 nessa base (varia de base para base).
4. Ajuste NATUREZAS_PROVAVEL: para bases "faixa", quais Naturezas Jurídicas
   contam como indício de vulnerabilidade e qual rótulo usar para cada uma
   (pode ter mais de uma, com "confianças" diferentes).
5. Rode: python empilhar_bases.py
6. Resultado: base_unificada.xlsx e base_unificada.csv
"""

import pandas as pd
from pathlib import Path

# ----------------------------------------------------------------------
# 1) CONFIGURAÇÃO DOS ARQUIVOS
# ----------------------------------------------------------------------
# AJUSTE AQUI: caminho, tipo, tipo_idade e aba de cada base.
# "aba": nome da aba com os microdados (linha a linha) dentro do .xlsx.
# A base de Mortes Violentas, por exemplo, tem os microdados na aba "Plan1"
# (as outras abas são tabelas dinâmicas de resumo, não os dados crus).
ARQUIVOS = [
    {"caminho": "dados/brutos/MORTES VIOLENTAS.xlsx", "tipo": "MORTES VIOLENTAS", "tipo_idade": "numerica", "aba": "Plan1"},
    {"caminho": "dados/brutos/ESTUPRO.xlsx", "tipo": "ESTUPRO", "tipo_idade": "faixa", "aba": "Plan1"},
    {"caminho": "dados/brutos/VIOLENCIA DOMESTICA.xlsx", "tipo": "VIOLENCIA DOMESTICA", "tipo_idade": "faixa", "aba": "Plan1"},
]

# ----------------------------------------------------------------------
# 2) MAPEAMENTO DE COLUNAS
# ----------------------------------------------------------------------
# AJUSTE AQUI: nome ORIGINAL da coluna (esquerda) -> nome PADRÃO (direita).
# Base "numerica": a coluna de idade deve virar "IDADE".
# Base "faixa": a coluna de faixa etária deve virar "FAIXA ETARIA".
# Não precisa mapear SITUAÇAO IDADE - o script cria essa coluna sozinho.
MAPA_COLUNAS = {
    "MORTES VIOLENTAS": {
        "MUNICIPIO DO FATO": "MUNICIPIO",
        "REGIAO GEOGRAFICA": "REGIAO GEOGRAFICA",
        "NATUREZA JURIDICA": "NATUREZA JURIDICA",
        "DATA DO FATO": "DATA",
        "ANO":"ANO",
        "SEXO": "SEXO",
        "IDADE": "IDADE",
        "SITUAÇAO IDADE": "SITUAÇAO IDADE",
        "TOTAL DE VITIMAS": "TOTAL DE VITIMAS",
    },
    "ESTUPRO": {
        "MUNICIPIO DO FATO": "MUNICIPIO",
        "REGIAO GEOGRAFICA": "REGIAO GEOGRAFICA",
        "NATUREZA JURIDICA": "NATUREZA JURIDICA",
        "DATA DO FATO": "DATA",
        "ANO":"ANO",
        "SEXO": "SEXO",
        "FAIXA ETARIA": "FAIXA ETARIA",
        "SITUAÇAO IDADE": "SITUAÇAO IDADE",
        "TOTAL DE VITIMAS": "TOTAL DE VITIMAS",

    },
    "VIOLENCIA DOMESTICA": {
        "MUNICIPIO DO FATO": "MUNICIPIO",
        "REGIAO GEOGRAFICA": "REGIAO GEOGRAFICA",
        "NATUREZA JURIDICA": "NATUREZA JURIDICA",
        "DATA DO FATO": "DATA",
        "ANO":"ANO",
        "SEXO": "SEXO",
        "FAIXA ETARIA": "FAIXA ETARIA",
        "SITUAÇAO IDADE": "SITUAÇAO IDADE",
        "TOTAL DE VITIMAS": "TOTAL DE VITIMAS",
        
    },
}

# ----------------------------------------------------------------------
# 3) FAIXAS CONSIDERADAS INFANTO-JUVENIL (bases tipo_idade="faixa")
# ----------------------------------------------------------------------
# AJUSTE AQUI: confira se os rótulos batem exatamente com os da sua planilha
# (maiúsculas, acentos, espaços) - se não bater, a linha não é reconhecida.
FAIXAS_INFANTO_JUVENIL = {
    "ESTUPRO": ["1) 00-11", "2) 12-17"],
    "VIOLENCIA DOMESTICA": ["1) 00-11", "2) 12-17"],
}

# ----------------------------------------------------------------------
# 4) NATUREZAS JURÍDICAS QUE INDICAM PROVÁVEL VULNERABILIDADE
# ----------------------------------------------------------------------
# AJUSTE AQUI: só usado quando FAIXA ETARIA está vazia. Para cada Natureza
# Jurídica que tem vínculo legal com menoridade, diga qual rótulo usar.
# Pode ter naturezas com "confiança" diferente (ex: Art. 217-A é presunção
# legal forte; Art. 136 é indício moderado, pode incluir idoso/PcD).
NATUREZAS_PROVAVEL = {
    "ESTUPRO": {
        "ESTUPRO DE VULNERÁVEL":
            "Provável vulnerável (Art. 217-A CP)",
        "ESTUPRO DE VULNERÁVEL POR VIOLÊNCIA DOMÉSTICA/FAMILIAR":
            "Provável vulnerável (Art. 217-A CP)",
    },
    "VIOLENCIA DOMESTICA": {
        "ESTUPRO DE VULNERÁVEL POR VIOLÊNCIA DOMÉSTICA/FAMILIAR":
            "Provável vulnerável (Art. 217-A CP)",
        "MAUS TRATOS POR VIOLÊNCIA DOMÉSTICA/FAMILIAR":
            "Provável vulnerável (Art. 136 CP - confiança moderada)",
    },
}

# Valores tratados como "vazio" além de NaN real (ex: "(vazio)", "Não informado")
VALORES_VAZIOS_EXTRA = ["(vazio)", "vazio", "não informado", "nao informado", "nan", ""]

# Colunas finais mantidas na base unificada, nessa ordem.
COLUNAS_FINAIS = [
    "TIPO", "DATA", "ANO", "MUNICIPIO", "REGIAO GEOGRAFICA",
    "IDADE", "FAIXA ETARIA", "SITUAÇAO IDADE", "NATUREZA JURIDICA",
    "SEXO", "TOTAL DE VITIMAS",
]


# ----------------------------------------------------------------------
# 5) FUNÇÕES
# ----------------------------------------------------------------------
def carregar_base(caminho: str, aba=None) -> pd.DataFrame:
    """Lê .xlsx (com controle de aba) ou .csv automaticamente pela extensão."""
    caminho = Path(caminho)
    if caminho.suffix.lower() in [".xlsx", ".xls"]:
        xls = pd.ExcelFile(caminho)
        if len(xls.sheet_names) > 1 and aba is None:
            print(f"  [AVISO] '{caminho.name}' tem {len(xls.sheet_names)} abas: {xls.sheet_names}")
            print(f"          Usando a primeira ('{xls.sheet_names[0]}'). Se não for a aba com\n"
                  "          os microdados, defina \"aba\" desse arquivo em ARQUIVOS.")
        aba_usar = aba if aba is not None else 0
        return pd.read_excel(caminho, sheet_name=aba_usar)
    elif caminho.suffix.lower() == ".csv":
        return pd.read_csv(caminho, sep=";", encoding="latin1")
    else:
        raise ValueError(f"Formato não suportado: {caminho.suffix}")


def _eh_vazio(serie: pd.Series) -> pd.Series:
    """NaN real OU strings tipo '(vazio)'/'Não informado' contam como vazio."""
    valores = serie.astype(str).str.strip().str.lower()
    return serie.isna() | valores.isin(VALORES_VAZIOS_EXTRA)


def classificar_idade_numerica(df: pd.DataFrame) -> pd.DataFrame:
    """Bases com IDADE em anos (ex: Mortes Violentas).
    Confirmado 0-17 = idade dentro da faixa. O restante (fora da faixa ou
    vazio) é descartado - já validamos que a ausência aqui é MCAR/MAR e o
    volume (<5%) é seguro para exclusão (ver relatório, seção 5.1)."""
    if "IDADE" not in df.columns:
        df["SITUAÇAO IDADE"] = pd.NA
        return df

    idade_num = pd.to_numeric(df["IDADE"], errors="coerce")
    dentro_da_faixa = (idade_num >= 0) & (idade_num <= 17)

    df["SITUAÇAO IDADE"] = pd.NA
    df.loc[dentro_da_faixa, "SITUAÇAO IDADE"] = "Confirmado 0-17"

    return df[dentro_da_faixa].copy()


def classificar_faixa_etaria(df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    """Bases com FAIXA ETARIA categórica (ex: Estupro, Violência Doméstica).
    Confirmado 0-17 = faixa já dentro do recorte.
    Provável = faixa vazia + Natureza Jurídica com vínculo legal a menoridade.
    O restante (faixa vazia sem vínculo legal, ou faixa fora do recorte) é
    descartado."""
    if "FAIXA ETARIA" not in df.columns:
        df["SITUAÇAO IDADE"] = pd.NA
        return df

    faixas_ok = FAIXAS_INFANTO_JUVENIL.get(tipo, [])
    naturezas_prov = NATUREZAS_PROVAVEL.get(tipo, {})

    faixa_valor = df["FAIXA ETARIA"].astype(str).str.strip()
    vazio = _eh_vazio(df["FAIXA ETARIA"])
    confirmado = faixa_valor.isin(faixas_ok)

    # Para cada linha, olha se a Natureza Jurídica está no dicionário de
    # naturezas prováveis dessa base e, se estiver, pega o rótulo definido.
    rotulo_provavel = df["NATUREZA JURIDICA"].map(naturezas_prov)
    provavel = vazio & rotulo_provavel.notna() & ~confirmado

    df["SITUAÇAO IDADE"] = pd.NA
    df.loc[confirmado, "SITUAÇAO IDADE"] = "Confirmado 0-17"
    df.loc[provavel, "SITUAÇAO IDADE"] = rotulo_provavel[provavel]

    mantido = confirmado | provavel
    return df[mantido].copy()


# ----------------------------------------------------------------------
# 6) PROCESSAMENTO
# ----------------------------------------------------------------------
def main():
    bases_tratadas = []

    for item in ARQUIVOS:
        tipo = item["tipo"]
        print(f"Lendo: {item['caminho']} ({tipo})")
        df = carregar_base(item["caminho"], aba=item.get("aba"))

        # Renomeia colunas conforme o mapeamento definido para essa base
        df = df.rename(columns=MAPA_COLUNAS.get(tipo, {}))

        # Marca a origem do registro
        df["TIPO"] = tipo

        # Classifica e já filtra: só sobra Confirmado + Provável
        if item.get("tipo_idade") == "faixa":
            df = classificar_faixa_etaria(df, tipo)
        else:
            df = classificar_idade_numerica(df)

        # Mantém só as colunas finais que existirem nessa base
        colunas_existentes = [c for c in COLUNAS_FINAIS if c in df.columns]
        df = df[colunas_existentes]

        print(f"  -> {len(df)} registros mantidos após classificação")
        print(f"  {df['SITUAÇAO IDADE'].value_counts(dropna=False).to_string()}\n")
        bases_tratadas.append(df)

    # Empilha todas as bases (concatena linhas, mantendo as colunas em comum)
    base_final = pd.concat(bases_tratadas, ignore_index=True, sort=False)

    print(f"Total empilhado: {len(base_final)} registros")
    print(base_final["TIPO"].value_counts())

    base_final.to_excel("base_unificada.xlsx", index=False)
    base_final.to_csv("base_unificada.csv", index=False, sep=";", encoding="utf-8-sig")
    print("\nArquivos gerados: base_unificada.xlsx e base_unificada.csv")


if __name__ == "__main__":
    main()