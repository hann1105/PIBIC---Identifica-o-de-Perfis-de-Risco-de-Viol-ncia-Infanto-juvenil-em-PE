"""
K-Means EXPLORATÓRIO — violência infantojuvenil em Pernambuco (SDS-PE)
Unidade de análise: MUNICÍPIO (1 linha por município)

Passos:
  1. Carrega e limpa a base
  2. Agrega por município (proporções de cada tipo de crime, sexo, faixa etária)
  3. Normaliza (min-max, como no plano de trabalho)
  4. Cotovelo + silhueta para escolher k
  5. Roda o K-Means final e imprime o perfil de cada cluster

Uso:
  python kmeans_exploratorio.py base_unificada.xlsx
"""
import sys
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from sklearn.preprocessing import MinMaxScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.decomposition import PCA

ARQUIVO = sys.argv[1] if len(sys.argv) > 1 else "base_unificada.xlsx"
ANO_MIN = 2004          # antes de 2012 a base tem pouquíssimos registros (provável outra coleta)
MIN_REGISTROS = 30      # municípios com menos de 30 registros têm proporções instáveis
K_MIN, K_MAX = 2, 8
SEED = 42 

# ------------------------------------------------------------------
# 1. CARGA E LIMPEZA
# ------------------------------------------------------------------
df = pd.read_excel(ARQUIVO)
df = df[df["ANO"] >= ANO_MIN].copy()

# Padroniza texto
for c in ["MUNICIPIO", "REGIAO GEOGRAFICA", "NATUREZA JURIDICA", "SEXO", "FAIXA ETARIA"]:
    df[c] = df[c].astype(str).str.strip().str.upper()

# Recife aparece como "RECIFE" e "CAPITAL" -> unifica
df["REGIAO GEOGRAFICA"] = df["REGIAO GEOGRAFICA"].replace({"CAPITAL": "RECIFE"})

# Região de cada município = a mais frequente nos registros dele
regiao_mun = (df.groupby("MUNICIPIO")["REGIAO GEOGRAFICA"]
                .agg(lambda s: s.value_counts().index[0]))

# ------------------------------------------------------------------
# 2. CATEGORIAS DE CRIME (agrupa as ~25 naturezas em 5 grupos)
# ------------------------------------------------------------------
def grupo_crime(nat: str) -> str:
    if "ESTUPRO" in nat and "VULNER" in nat:
        return "estupro_vulneravel"
    if "ESTUPRO" in nat:
        return "estupro"
    if any(k in nat for k in ["HOMICIDIO", "LATROCINIO", "FEMINICIDIO",
                              "INTERVENCAO", "SEGUIDA DE MORTE"]):
        return "morte_violenta"
    if any(k in nat for k in ["LESÃO CORPORAL", "VIAS DE FATO", "MAUS TRATOS"]):
        return "dom_fisica"
    return "dom_outras"   # ameaça, injúria, difamação, dano, etc.

df["grupo_crime"] = df["NATUREZA JURIDICA"].apply(grupo_crime)
df["familiar"] = df["NATUREZA JURIDICA"].str.contains("DOMÉSTICA/FAMILIAR", regex=False)

# Pesa pelo nº de vítimas (uma ocorrência pode ter várias)
w = df["TOTAL DE VITIMAS"]

# ------------------------------------------------------------------
# 3. AGREGAÇÃO POR MUNICÍPIO
# ------------------------------------------------------------------
def proporcao(mask_col, valor=True):
    """proporção ponderada de linhas em que mask_col == valor, por município"""
    num = (w * (df[mask_col] == valor)).groupby(df["MUNICIPIO"]).sum()
    den = w.groupby(df["MUNICIPIO"]).sum()
    return num / den

feat = pd.DataFrame(index=sorted(df["MUNICIPIO"].unique()))
feat["n_registros"] = df.groupby("MUNICIPIO").size()
feat["n_vitimas"] = w.groupby(df["MUNICIPIO"]).sum()

# proporção de cada grupo de crime
for g in ["estupro_vulneravel", "estupro", "dom_fisica", "dom_outras", "morte_violenta"]:
    feat[f"p_{g}"] = proporcao("grupo_crime", g)

feat["p_familiar"] = proporcao("familiar", True)
feat["p_feminino"] = proporcao("SEXO", "FEMININO")

# faixa etária: ignora "(VAZIO)" no denominador
com_faixa = df[df["FAIXA ETARIA"] != "(VAZIO)"]
w2 = com_faixa["TOTAL DE VITIMAS"]
den2 = w2.groupby(com_faixa["MUNICIPIO"]).sum()
feat["p_0_11"] = (w2 * com_faixa["FAIXA ETARIA"].str.contains("00-11")).groupby(
    com_faixa["MUNICIPIO"]).sum() / den2
feat["p_0_11"] = feat["p_0_11"].fillna(feat["p_0_11"].median())  # raro; só se faltar tudo

feat["regiao"] = regiao_mun

# filtra municípios com poucos registros
n_antes = len(feat)
feat = feat[feat["n_registros"] >= MIN_REGISTROS].copy()
print(f"Municípios: {n_antes} -> {len(feat)} (após filtro de >= {MIN_REGISTROS} registros)\n")

COLS = ["p_estupro_vulneravel", "p_estupro", "p_dom_fisica", "p_dom_outras",
        "p_morte_violenta", "p_familiar", "p_feminino", "p_0_11"]

# ------------------------------------------------------------------
# 4. NORMALIZAÇÃO (min-max) E ESCOLHA DE k
# ------------------------------------------------------------------
X = MinMaxScaler().fit_transform(feat[COLS])

ks = list(range(K_MIN, K_MAX + 1))
inercia, silhueta = [], []
for k in ks:
    km = KMeans(n_clusters=k, init="k-means++", n_init=20, random_state=SEED).fit(X)
    inercia.append(km.inertia_)
    silhueta.append(silhouette_score(X, km.labels_))

fig, ax = plt.subplots(1, 2, figsize=(10, 4))
ax[0].plot(ks, inercia, "o-"); ax[0].set_title("Cotovelo (inércia)"); ax[0].set_xlabel("k")
ax[1].plot(ks, silhueta, "o-", color="tab:orange"); ax[1].set_title("Índice de silhueta"); ax[1].set_xlabel("k")
plt.tight_layout(); plt.savefig("kmeans_cotovelo_silhueta.png", dpi=150)

print("k  | inércia | silhueta")
for k, i, s in zip(ks, inercia, silhueta):
    print(f"{k}  | {i:7.2f} | {s:.3f}")

k_final = ks[int(np.argmax(silhueta))]
print(f"\nMelhor silhueta: k = {k_final}  (veja também o cotovelo no PNG; "
      f"vale testar k vizinhos e escolher o mais interpretável)\n")

# ------------------------------------------------------------------
# 5. K-MEANS FINAL + PERFIL DOS CLUSTERS
# ------------------------------------------------------------------
km = KMeans(n_clusters=k_final, init="k-means++", n_init=20, random_state=SEED).fit(X)
feat["cluster"] = km.labels_

pd.set_option("display.width", 200)
pd.set_option("display.max_columns", 30)

print("=== Perfil médio por cluster (proporções) ===")
perfil = feat.groupby("cluster")[COLS].mean().round(3)
perfil.insert(0, "n_municipios", feat.groupby("cluster").size())
print(perfil, "\n")

print("=== Região x cluster (nº de municípios) ===")
print(pd.crosstab(feat["cluster"], feat["regiao"]), "\n")

print("=== Maiores municípios de cada cluster (por nº de registros) ===")
for c, g in feat.groupby("cluster"):
    print(f"cluster {c}: ", ", ".join(g.sort_values("n_registros", ascending=False).head(6).index))

# visual 2D (PCA) só para enxergar
pca = PCA(n_components=2, random_state=SEED).fit_transform(X)
plt.figure(figsize=(7, 5))
plt.scatter(pca[:, 0], pca[:, 1], c=feat["cluster"], cmap="tab10", s=25)
plt.title(f"Municípios agrupados (k={k_final}) — projeção PCA"); plt.tight_layout()
plt.savefig("kmeans_pca.png", dpi=150)

feat.to_csv("municipios_clusters.csv", encoding="utf-8-sig")
print("\nArquivos gerados: municipios_clusters.csv, kmeans_cotovelo_silhueta.png, kmeans_pca.png")
