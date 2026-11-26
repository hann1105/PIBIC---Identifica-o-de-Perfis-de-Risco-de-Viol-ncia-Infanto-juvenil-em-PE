import re
import time
import unicodedata
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pathlib import Path

from config import FONTES, PALAVRAS_ARQUIVO, PASTA_BRUTOS

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/153.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
}

EXTENSOES = (".xlsx", ".xls", ".csv", ".zip")
MESES = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
    "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3,
    "ABRIL": 4, "MAIO": 5, "JUNHO": 6, "JULHO": 7,
    "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10,
    "NOVEMBRO": 11, "DEZEMBRO": 12,
}


def texto_normalizado(texto):
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()


def extrair_periodo_final(nome):
    """Retorna (ano, mes) do período final no nome do arquivo.

    Exemplos aceitos:
      ..._A_AGO_2026.xlsx
      ..._A_AGOSTO_2026.xlsx
      ..._A_SET_2026.xlsx
    """
    base = texto_normalizado(nome)
    padrao = (
        r"_A_"
        r"(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ|"
        r"JANEIRO|FEVEREIRO|MARCO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|"
        r"SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)_"
        r"(20\d{2})"
    )
    m = re.search(padrao, base, re.IGNORECASE)
    if not m:
        return None
    mes = MESES[m.group(1).upper()]
    ano = int(m.group(2))
    return ano, mes


def encontrar_links_planilhas(sessao, url):
    resposta = sessao.get(url, timeout=60)
    resposta.raise_for_status()

    soup = BeautifulSoup(resposta.text, "html.parser")
    links = []

    for a in soup.find_all("a", href=True):
        href = urljoin(url, a["href"])
        texto = a.get_text(" ", strip=True)
        alvo = f"{texto} {href}"
        extensao = href.lower().split("?")[0]

        if any(extensao.endswith(ext) for ext in EXTENSOES):
            links.append((texto, href, alvo))

    vistos = set()
    saida = []
    for item in links:
        if item[1] not in vistos:
            vistos.add(item[1])
            saida.append(item)
    return saida


def pontuar_link(texto, href, tipo):
    alvo = texto_normalizado(f"{texto} {href}")
    pontos = 0

    for palavra in PALAVRAS_ARQUIVO[tipo]:
        if texto_normalizado(palavra) in alvo:
            pontos += 4

    for palavra in ["dados", "microdados", "base", "planilha"]:
        if palavra in alvo:
            pontos += 1

    periodo = extrair_periodo_final(href)
    if periodo:
        pontos += periodo[0] * 100 + periodo[1]

    return pontos


def escolher_link(links, tipo):
    if not links:
        raise RuntimeError(f"Nenhuma planilha encontrada para {tipo}.")

    # Primeiro restringe pelos termos do indicador; depois usa o período final.
    candidatos = []
    palavras = [texto_normalizado(p) for p in PALAVRAS_ARQUIVO[tipo]]

    for item in links:
        alvo = texto_normalizado(f"{item[0]} {item[1]}")
        if any(p in alvo for p in palavras):
            candidatos.append(item)

    if not candidatos:
        candidatos = links

    com_periodo = [x for x in candidatos if extrair_periodo_final(x[1])]
    if com_periodo:
        escolhido = max(
            com_periodo,
            key=lambda x: (extrair_periodo_final(x[1])[0], extrair_periodo_final(x[1])[1])
        )
    else:
        escolhido = max(candidatos, key=lambda x: pontuar_link(x[0], x[1], tipo))

    print(f"[{tipo}] Links de planilha encontrados: {len(links)}")
    print(f"[{tipo}] Arquivo mais recente escolhido: {escolhido[1]}")
    return escolhido[1]


def nome_seguro(url, tipo):
    nome = url.split("/")[-1].split("?")[0]
    nome = re.sub(r"[^A-Za-z0-9._-]", "_", nome)
    if not Path(nome).suffix:
        nome = f"{tipo.lower().replace(' ', '_')}.xlsx"
    return nome


def baixar_arquivo(sessao, url, destino, pagina_origem):
    # Algumas hospedagens recusam o XLSX se o acesso não vier da página da SDS.
    headers = dict(HEADERS)
    headers["Referer"] = pagina_origem

    ultimo_erro = None
    for tentativa in range(1, 4):
        try:
            with sessao.get(url, headers=headers, timeout=180, stream=True) as r:
                r.raise_for_status()
                content_type = (r.headers.get("content-type") or "").lower()

                total = 0
                with open(destino, "wb") as f:
                    for bloco in r.iter_content(chunk_size=1024 * 1024):
                        if bloco:
                            f.write(bloco)
                            total += len(bloco)

            if total < 10_000:
                raise RuntimeError(
                    f"Resposta muito pequena ({total} bytes). A SDS pode ter devolvido uma página de erro."
                )

            print(f"[OK] Download concluído: {total / 1024 / 1024:.2f} MB")
            print(f"[OK] Content-Type: {content_type}")
            return
        except Exception as erro:
            ultimo_erro = erro
            print(f"[AVISO] Tentativa {tentativa}/3 falhou: {erro}")
            time.sleep(2 * tentativa)

    raise RuntimeError(f"Não foi possível baixar {url}. Último erro: {ultimo_erro}")


def atualizar_fontes():
    resultados = []
    sessao = requests.Session()
    sessao.headers.update(HEADERS)

    for tipo, cfg in FONTES.items():
        print(f"\n=== {tipo} ===")
        try:
            links = encontrar_links_planilhas(sessao, cfg["url"])
            arquivo_url = escolher_link(links, tipo)
            nome = nome_seguro(arquivo_url, tipo)
            destino = PASTA_BRUTOS / f"{tipo.lower().replace(' ', '_')}_{nome}"
            baixar_arquivo(sessao, arquivo_url, destino, cfg["url"])

            resultados.append({
                "tipo": tipo,
                "url_pagina": cfg["url"],
                "url_arquivo": arquivo_url,
                "arquivo": destino,
                "tipo_idade": cfg["tipo_idade"],
            })
        except Exception as erro:
            print(f"[ERRO] {tipo}: {erro}")
            print("       O pipeline continua com as outras fontes.")

    return resultados


if __name__ == "__main__":
    resultados = atualizar_fontes()
    print(f"\nFontes baixadas com sucesso: {len(resultados)}/{len(FONTES)}")
