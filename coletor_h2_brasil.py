#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Coletor de dados sobre Hidrogênio de Baixa Emissão de Carbono no Brasil
=========================================================================

Fontes cobertas por padrão (podem ser editadas na seção CONFIGURAÇÃO):
- dados.gov.br (novo e legado)  -> Portal Brasileiro de Dados Abertos (CKAN)
- dadosabertos.aneel.gov.br     -> Dados Abertos da ANEEL (CKAN)
- gisepeprd2.epe.gov.br         -> Camadas ArcGIS REST da EPE (descoberta automática)
- api.eia.gov                   -> EIA - U.S. Energy Information Administration (requer chave)
- energydata.info               -> Espelho da base de projetos de H2 da IEA (World Bank ESMAP)
- epe.gov.br, h2portal.com.br   -> Publicações e Portal Brasileiro de Hidrogênio (HTML)
- gov.br/mme, bndes.gov.br      -> MME (PNH2) e BNDES (HTML)
- gov.br/anp                    -> ANP - regulador oficial do H2 (Lei 14.948/2024) (HTML)
- repositorio.ipea.gov.br       -> IPEA - estudos técnicos sobre hidrogênio (HTML)
- dadosabertos.ccee.org.br      -> CCEE - dados abertos do mercado de energia (HTML)
- dados.ons.org.br              -> ONS - dados abertos do sistema elétrico (HTML)
- h2lac.org, h2brazil.com.br    -> H2LAC e H2 Brazil/GIZ (HTML)

Além da coleta, o script também executa (Seção 3 do TCC):
- Resolução de entidades (Algoritmo 1: blocking + similaridade de Levenshtein)
  sobre a base padronizada, para eliminar duplicatas entre fontes distintas.
  Gera: base_projetos_deduplicada.xlsx e pares_similares_resolucao_entidades.xlsx
- Cálculo do LCOH de referência (Equações 5-8, Tabela 2) por cenário
  (conservador/otimista) e tecnologia (ALK/PEM). Gera: lcoh_projecao_brasil.xlsx

Como usar:
    pip install -r requirements.txt
    python coletor_h2_brasil.py

Requisitos: requests, beautifulsoup4, pandas, openpyxl
"""

import os
import re
import sys
import json
import time
import logging
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
import pandas as pd
from bs4 import BeautifulSoup

# --------------------------------------------------------------------------
# CONFIGURAÇÃO
# --------------------------------------------------------------------------

TERMOS_BUSCA = [
    "hidrogênio",
    "hidrogenio",
    "hidrogênio de baixo carbono",
    "H2 baixa emissão",
    "hydrogen",
    "green hydrogen",
    "low-carbon hydrogen",
]

# OBS: dados.gov.br passou a exigir uma chave de API pessoal para consultas
# (gere a sua em https://dados.gov.br, na sua conta > "Tokens de acesso").
# Sem chave, essa fonte é pulada automaticamente (não é tratada como erro).
#
# "IEA / ENERGYDATA.INFO": espelho público (World Bank ESMAP, plataforma CKAN)
# da Hydrogen Production and Infrastructure Projects Database da IEA — citada
# na Tabela de Fontes do TCC. O site oficial da IEA exige login para baixar a
# versão mais atual; este espelho não exige login, mas pode conter uma versão
# desatualizada. Para a versão mais recente, veja PASTA_ENTRADA_MANUAL abaixo.
PORTAIS_CKAN = {
    "dados.gov.br": "https://dados.gov.br",
    "ANEEL - Dados Abertos": "https://dadosabertos.aneel.gov.br",
    "IEA / ENERGYDATA.INFO (espelho, World Bank ESMAP)": "https://energydata.info",
}
# OBS: "legado.dados.gov.br" foi testado e removido — o domínio não resolve
# mais (foi descontinuado pelo governo federal).

CHAVE_API_DADOS_GOV = os.environ.get("DADOS_GOV_API_KEY", "").strip()
CABECALHO_API_DADOS_GOV = "chave-api-dados-gov"  # nome do header exigido pelo portal

# A IEA disponibiliza a versão mais atual da Hydrogen Production and
# Infrastructure Projects Database só mediante login gratuito em
# https://www.iea.org/data-and-statistics/data-product/hydrogen-production-and-infrastructure-projects-database
# — não é possível automatizar esse download sem violar os termos de uso.
# Se você baixar esse arquivo manualmente (ou qualquer outra planilha do
# H2LAC, BNDES, MME etc. que não tenha link direto), basta soltar o arquivo
# nesta pasta que o script incorpora automaticamente ao índice final.
PASTA_ENTRADA_MANUAL = Path("dados_h2/entrada_manual")


# Camadas de dados geoespaciais da EPE (ArcGIS REST, público, sem necessidade de
# login) que alimentam o "Painel de Dados sobre Hidrogênio" do Portal Brasileiro
# de Hidrogênio. Em vez de fixar nomes de camada (que a EPE reedita periodicamente,
# ex.: H2_PROJETOS_202404, H2_PROJETOS_202410...), o script DESCOBRE automaticamente
# todos os serviços da pasta abaixo cujo nome contenha "H2", agrupa versões do mesmo
# tema e mantém apenas a versão mais recente de cada grupo.
ARCGIS_REST_BASE = "https://gisepeprd2.epe.gov.br/arcgis/rest/services"
ARCGIS_PASTA_H2 = "SEE"  # pasta onde a EPE publica as camadas de hidrogênio
MAX_CAMADAS_ARCGIS = 15  # limite de segurança para não varrer serviços demais

# Fallback manual: usado apenas se a descoberta automática acima falhar (ex.: a
# EPE reestruturar a pasta). Mantém pelo menos a base de projetos funcionando.
CAMADAS_ARCGIS_H2_FALLBACK = {
    "EPE - Projetos de Hidrogênio no Brasil": (
        "https://gisepeprd2.epe.gov.br/arcgis/rest/services/SEE/H2_PROJETOS_202410/MapServer/0"
    ),
}

# Nem todo serviço com "H2" no nome é realmente sobre hidrogênio: a EPE reaproveita
# o mesmo diretório para camadas de INFRAESTRUTURA GERAL do setor elétrico (usinas
# existentes/planejadas, linhas de transmissão, unidades de conservação, zonas
# marítimas etc.) usadas como insumo para escolher onde instalar projetos de H2 —
# mas o conteúdo em si não é uma "base de hidrogênio". Esses serviços ficam de fora
# por padrão. Mude para True se você quiser essa camada de suporte também.
INCLUIR_CAMADAS_DE_INFRAESTRUTURA_GERAL = False

# Nomes de serviço (a parte depois de "SEE/") que sabidamente contêm só
# infraestrutura genérica reaproveitada, não dados de hidrogênio propriamente
# ditos. Comparação ignora acentos/caixa.
SERVICOS_INFRAESTRUTURA_GERAL = {
    "h2_db", "db_see_h2", "dashboardh2", "h2_dbpoints", "dashboard_points",
}

# Quantos registros buscar por página nas consultas ArcGIS, e quantas páginas no
# máximo por camada (evita truncar silenciosamente em serviços com >1000 linhas,
# e evita ficar rodando para sempre em camadas gigantes).
TAMANHO_PAGINA_ARCGIS = 1000
MAX_PAGINAS_ARCGIS = 10  # até 10.000 registros por camada

# --- EIA (U.S. Energy Information Administration) ------------------------
# Fonte internacional complementar. A EIA não tem uma categoria "hidrogênio"
# dedicada, então em vez de fixar um código de produto (que pode não existir
# ou mudar), o script CONSULTA o catálogo de produtos da rota "international"
# e usa apenas os que mencionarem "hydrogen" no nome — se não houver nenhum,
# ele avisa e não força um resultado vazio ou errado.
# Gere sua chave gratuita em: https://www.eia.gov/opendata/register.php
EIA_API_KEY = os.environ.get("EIA_API_KEY", "").strip()
EIA_API_BASE = "https://api.eia.gov/v2"
EIA_PAIS_FOCO = "BRA"  # código de país da EIA para o Brasil (ISO 3166-1 alpha-3)

PAGINAS_HTML = {
    "EPE - Publicações e Dados Abertos": "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes",
    "Portal Brasileiro de Hidrogênio": "https://h2portal.com.br/",
    "MME - Programa Nacional do Hidrogênio (PNH2)": "https://www.gov.br/mme/pt-br/programa-nacional-do-hidrogenio-1",
    "BNDES - Hidrogênio de Baixo Carbono": "https://www.bndes.gov.br/wps/portal/site/home/onde_atuamos/verde/hidrogenio-baixo-carbono",
    "H2LAC (Hidrogênio América Latina e Caribe)": "https://h2lac.org/",
    # Novas fontes: ANP passou a regular oficialmente a cadeia do hidrogênio
    # pela Lei 14.948/2024; IPEA publica estudos técnicos sobre o setor; CCEE
    # e ONS são os dados abertos do mercado/sistema elétrico (relevantes para
    # o custo da eletricidade renovável usada na eletrólise); H2 Brazil é o
    # portfólio de projetos da cooperação GIZ-MME citado no TCC.
    "ANP - Dados Abertos": "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos",
    "IPEA - Repositório Institucional": "https://repositorio.ipea.gov.br/",
    "CCEE - Dados Abertos": "https://dadosabertos.ccee.org.br/",
    "ONS - Dados Abertos": "https://dados.ons.org.br/",
    "H2 Brazil (GIZ-MME)": "https://h2brazil.com.br/",
}

EXTENSOES_DADOS = (".xlsx", ".xls", ".csv", ".json")

EXTENSOES_EXTRA = (".pdf",)
BAIXAR_PDFS_TAMBEM = False

MAX_PROFUNDIDADE_CRAWLER = 2
MAX_PAGINAS_CRAWLER_POR_FONTE = 80
SEGUIR_LINKS_SEM_EXTENSAO = True

CAMINHOS_RELEVANTES_CRAWLER = (
    "hidrogen",
    "hydrogen",
    "h2",
    "projeto",
    "projetos",
    "dados",
    "dataset",
    "datasets",
    "publicacao",
    "publicacoes",
    "documento",
    "documentos",
    "noticia",
    "noticias",
    "pesquisa",
    "estudo",
    "energia",
    "renovavel",
    "eletrol",
    "amonia",
    "hub",
    "pnh2",
)

PALAVRAS_CHAVE_RELEVANCIA = [
    "hidrogenio",
    "h2 verde",
    "h2v",
    "hidrogenio verde",
    "hidrogenio azul",
    "hidrogenio cinza",
    "baixo carbono",
    "baixa emissao de carbono",
    "pnh2",
    "amonia verde",
    # termos em inglês, necessários para fontes internacionais (IEA, H2LAC etc.)
    "hydrogen",
    "green hydrogen",
    "blue hydrogen",
    "low-carbon hydrogen",
    "low carbon hydrogen",
]

PASTA_SAIDA = Path("dados_h2")
PASTA_BRUTOS = PASTA_SAIDA / "brutos"
PASTA_CONVERTIDOS = PASTA_SAIDA / "convertidos"
ARQUIVO_INDICE = PASTA_SAIDA / "indice_bases_hidrogenio.xlsx"
ARQUIVO_CHECKPOINT = PASTA_SAIDA / "checkpoint.json"
PASTA_HISTORICO = PASTA_SAIDA / "historico"

# Importante para execução AUTOMÁTICA/SEMANAL: os arquivos baixados via HTML
# usam um prefixo numérico (ex.: "003_titulo.xlsx") cuja ordem pode mudar de
# uma execução para outra — sem limpar a pasta antes, arquivos antigos ficam
# acumulados junto dos novos. Por padrão, cada execução arquiva o resultado
# da rodada anterior em dados_h2/historico/<data>/ (histórico útil pra
# acompanhar como a base evolui semana a semana) e começa a coleta do zero.
ARQUIVAR_EXECUCAO_ANTERIOR = True

# --- Fontes que são SPAs em JavaScript (candidatas ao Selenium opcional) --
FONTES_SPA_JAVASCRIPT = {
    "CCEE - Dados Abertos",
    "ONS - Dados Abertos",
    "H2LAC (Hidrogênio América Latina e Caribe)",
    "H2 Brazil (GIZ-MME)",
}
USAR_SELENIUM_PARA_SPA = os.environ.get("USAR_SELENIUM_PARA_SPA", "0").strip() == "1"
TEMPO_ESPERA_SELENIUM = 8  # segundos de espera pelo JS renderizar a página

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (compatible; ColetorH2Brasil/1.0; "
        "+https://www.gov.br/mme/pt-br/programa-nacional-do-hidrogenio-1)"
    )
}

TIMEOUT = 30  # segundos por requisição
PAUSA_ENTRE_REQUISICOES = 1.0  # segundos, para não sobrecarregar os servidores

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("coletor_h2")


# --------------------------------------------------------------------------
# ESTRUTURA DE DADOS
# --------------------------------------------------------------------------

@dataclass
class ItemEncontrado:
    fonte: str
    titulo: str
    formato: str
    url: str
    dataset_ou_pagina: str = ""
    caminho_local: str = ""
    status_download: str = ""


ITENS_ENCONTRADOS: list[ItemEncontrado] = []


# --------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# --------------------------------------------------------------------------

def normalizar_texto(texto: str) -> str:
    """Remove acentos e baixa a caixa, para comparação de relevância."""
    if not texto:
        return ""
    forma = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in forma if not unicodedata.combining(c))
    return sem_acento.lower()


def contexto_e_relevante(*textos: str) -> bool:
    """Verifica se algum dos textos fornecidos menciona hidrogênio de fato."""
    conjunto = normalizar_texto(" ".join(t for t in textos if t))
    return any(normalizar_texto(chave) in conjunto for chave in PALAVRAS_CHAVE_RELEVANCIA)


def nome_arquivo_seguro(texto: str, maximo: int = 120) -> str:
    """Remove caracteres problemáticos para usar como nome de arquivo."""
    texto = re.sub(r"[^\w\-.]+", "_", texto, flags=re.UNICODE)
    texto = re.sub(r"_{2,}", "_", texto).strip("_")
    return texto[:maximo] if texto else "arquivo"


def baixar_arquivo(url: str, destino: Path) -> bool:
    """Baixa um arquivo binário para o caminho de destino. Retorna True se ok."""
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        with requests.get(url, headers=HEADERS, timeout=TIMEOUT, stream=True) as resp:
            resp.raise_for_status()
            with open(destino, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
        return True
    except requests.RequestException as e:
        log.warning(f"  Falha ao baixar {url}: {e}")
        return False


def converter_csv_para_xlsx(caminho_csv: Path) -> Path | None:
    """Gera uma cópia .xlsx de um CSV baixado, para facilitar a análise."""
    try:
        try:
            df = pd.read_csv(caminho_csv, sep=None, engine="python", encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(caminho_csv, sep=None, engine="python", encoding="latin-1")

        PASTA_CONVERTIDOS.mkdir(parents=True, exist_ok=True)
        destino = PASTA_CONVERTIDOS / (caminho_csv.stem + ".xlsx")
        df.to_excel(destino, index=False)
        return destino
    except Exception as e:
        log.warning(f"  Não foi possível converter {caminho_csv.name} para xlsx: {e}")
        return None


# --------------------------------------------------------------------------
# COLETA EM PORTAIS CKAN (dados.gov.br, ANEEL, etc.)
# --------------------------------------------------------------------------

def buscar_em_portal_ckan(nome_fonte: str, url_base: str):
    log.info(f"== Buscando datasets em: {nome_fonte} ({url_base}) ==")
    endpoint = urljoin(url_base, "/api/3/action/package_search")

    cabecalhos = dict(HEADERS)
    dominio = urlparse(url_base).netloc.lower()
    if dominio == "dados.gov.br":  # a versão NOVA exige chave; o "legado." não
        if not CHAVE_API_DADOS_GOV:
            log.warning(
                "  Pulando dados.gov.br: este portal agora exige uma chave de API "
                "pessoal. Gere a sua em https://dados.gov.br (Minha conta > Tokens "
                "de acesso) e defina a variável de ambiente DADOS_GOV_API_KEY."
            )
            return
        cabecalhos[CABECALHO_API_DADOS_GOV] = CHAVE_API_DADOS_GOV

    ids_ja_vistos = set()

    for termo in TERMOS_BUSCA:
        try:
            resp = requests.get(
                endpoint,
                params={"q": termo, "rows": 50},
                headers=cabecalhos,
                timeout=TIMEOUT,
            )
            resp.raise_for_status()
            dados = resp.json()
        except (requests.RequestException, ValueError) as e:
            log.warning(f"  Falha ao consultar '{termo}' em {nome_fonte}: {e}")
            continue

        if not dados.get("success"):
            continue

        resultados = dados.get("result", {}).get("results", [])
        log.info(f"  Termo '{termo}': {len(resultados)} dataset(s) encontrado(s)")

        for pacote in resultados:
            pkg_id = pacote.get("id")
            if pkg_id in ids_ja_vistos:
                continue
            ids_ja_vistos.add(pkg_id)

            titulo_dataset = pacote.get("title", "sem_titulo")
            descricao_dataset = pacote.get("notes", "")
            tags_dataset = " ".join(
                t.get("name", "") for t in pacote.get("tags", []) if isinstance(t, dict)
            )

            if not contexto_e_relevante(titulo_dataset, descricao_dataset, tags_dataset):
                log.info(f"    (descartado por baixa relevância: '{titulo_dataset}')")
                continue

            recursos = pacote.get("resources", [])

            for recurso in recursos:
                fmt = (recurso.get("format") or "").lower().strip()
                rec_url = recurso.get("url", "")
                rec_titulo = recurso.get("name") or titulo_dataset

                if not rec_url:
                    continue

                ext = os.path.splitext(urlparse(rec_url).path)[1].lower()
                eh_dado = ext in EXTENSOES_DADOS or fmt in ("csv", "xlsx", "xls", "json")
                eh_extra = ext in EXTENSOES_EXTRA or fmt == "pdf"

                if not (eh_dado or eh_extra):
                    continue

                item = ItemEncontrado(
                    fonte=nome_fonte,
                    titulo=f"{titulo_dataset} — {rec_titulo}",
                    formato=ext.lstrip(".") or fmt,
                    url=rec_url,
                    dataset_ou_pagina=titulo_dataset,
                )
                ITENS_ENCONTRADOS.append(item)

        time.sleep(PAUSA_ENTRE_REQUISICOES)


# --------------------------------------------------------------------------
# DESCOBERTA AUTOMÁTICA DE CAMADAS ArcGIS DE HIDROGÊNIO (EPE)
# --------------------------------------------------------------------------

def _extrair_base_e_sufixo(nome_servico: str) -> tuple[str, int]:
    """Separa o nome de um serviço em (nome-base, sufixo numérico de versão).
    Ex.: 'H2_PROJETOS_202410' -> ('H2_PROJETOS', 202410)
         'H2_Capacitacao_1'   -> ('H2_Capacitacao', 1)
         'Educacao_H2'        -> ('Educacao_H2', -1)  (sem sufixo numérico)
    """
    m = re.search(r"^(.*?)_?(\d{1,8})$", nome_servico)
    if m and m.group(1):
        return m.group(1), int(m.group(2))
    return nome_servico, -1


def _listar_servicos_pasta_arcgis(pasta: str) -> list[dict]:
    endpoint = f"{ARCGIS_REST_BASE}/{pasta}"
    resp = requests.get(endpoint, params={"f": "json"}, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json().get("services", [])


def _listar_camadas_servico(url_servico: str) -> list[dict]:
    resp = requests.get(url_servico, params={"f": "json"}, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json().get("layers", [])


def descobrir_camadas_arcgis_h2() -> dict:
    """Varre a pasta de serviços da EPE, seleciona os que mencionam 'H2' e
    mantém apenas a versão mais recente de cada tema (por sufixo numérico de
    data/versão no nome). Retorna {nome_amigável: url_da_camada}."""
    log.info(f"== Descobrindo camadas de hidrogênio em ArcGIS REST (pasta '{ARCGIS_PASTA_H2}') ==")

    try:
        servicos = _listar_servicos_pasta_arcgis(ARCGIS_PASTA_H2)
    except (requests.RequestException, ValueError) as e:
        log.warning(f"  Falha ao listar serviços da pasta '{ARCGIS_PASTA_H2}': {e}")
        log.warning("  Usando lista de camadas de reserva (fallback manual).")
        return dict(CAMADAS_ARCGIS_H2_FALLBACK)

    candidatos = [
        s for s in servicos
        if "h2" in normalizar_texto(s.get("name", "").split("/")[-1]).replace(" ", "")
    ]
    log.info(f"  {len(candidatos)} serviço(s) com 'H2' no nome encontrados de {len(servicos)} total.")

    if not INCLUIR_CAMADAS_DE_INFRAESTRUTURA_GERAL:
        antes = len(candidatos)
        candidatos = [
            s for s in candidatos
            if normalizar_texto(s.get("name", "").split("/")[-1]) not in SERVICOS_INFRAESTRUTURA_GERAL
        ]
        removidos = antes - len(candidatos)
        if removidos:
            log.info(
                f"  {removidos} serviço(s) de infraestrutura geral (não específicos de H2) "
                f"ignorado(s) — mude INCLUIR_CAMADAS_DE_INFRAESTRUTURA_GERAL para True se quiser incluí-los."
            )

    # Agrupa por tema, mantendo só a versão de maior sufixo numérico (mais recente)
    grupos: dict[str, dict] = {}
    for s in candidatos:
        nome_completo = s.get("name", "")  # ex.: "SEE/H2_PROJETOS_202410"
        nome_curto = nome_completo.split("/")[-1]
        base, sufixo = _extrair_base_e_sufixo(nome_curto)
        chave = normalizar_texto(base)

        if chave not in grupos or sufixo > grupos[chave]["sufixo"]:
            grupos[chave] = {
                "nome_completo": nome_completo,
                "nome_curto": nome_curto,
                "sufixo": sufixo,
                "tipo": s.get("type", "MapServer"),
            }

    selecionados = list(grupos.values())[:MAX_CAMADAS_ARCGIS]
    log.info(f"  {len(selecionados)} tema(s) únicos selecionados (versão mais recente de cada).")

    camadas = {}
    for item in selecionados:
        url_servico = f"{ARCGIS_REST_BASE}/{item['nome_completo']}/{item['tipo']}"
        try:
            sub_camadas = _listar_camadas_servico(url_servico)
        except (requests.RequestException, ValueError) as e:
            log.warning(f"  Falha ao listar camadas de {item['nome_curto']}: {e}")
            continue

        for camada in sub_camadas:
            layer_id = camada.get("id")
            layer_nome = camada.get("name", f"camada{layer_id}")
            nome_amigavel = f"EPE - {item['nome_curto']} - {layer_nome}"
            camadas[nome_amigavel] = f"{url_servico}/{layer_id}"

    if not camadas:
        log.warning("  Nenhuma camada válida descoberta — usando lista de reserva (fallback manual).")
        return dict(CAMADAS_ARCGIS_H2_FALLBACK)

    return camadas


# --------------------------------------------------------------------------
# COLETA EM CAMADAS ArcGIS REST (base real de projetos de H2 da EPE)
# --------------------------------------------------------------------------

def buscar_em_camada_arcgis(nome_fonte: str, url_camada: str):
    """Consulta uma Feature Layer do ArcGIS REST e baixa todos os registros
    como tabela (sem precisar passar pela interface de mapa/StoryMap)."""
    log.info(f"== Consultando camada ArcGIS: {nome_fonte} ==")

    # 1) Verifica o tipo da camada antes de tentar consultar como tabela.
    #    Camadas Raster/Image (ex.: batimetria) não suportam query de atributos
    #    e só gerariam erro 400/500 — pulamos direto, economizando tempo.
    try:
        resp_meta = requests.get(
            url_camada, params={"f": "json"}, headers=HEADERS, timeout=TIMEOUT
        )
        resp_meta.raise_for_status()
        meta = resp_meta.json()
    except (requests.RequestException, ValueError) as e:
        log.warning(f"  Falha ao obter metadados de {nome_fonte}: {e}")
        return

    tipo_camada = meta.get("type", "")
    if tipo_camada and tipo_camada not in ("Feature Layer", "Table"):
        log.info(f"  Pulando ({tipo_camada} não é tabular/consultável como base de dados)")
        return

    # A maioria das camadas da EPE roda em uma versão de ArcGIS Server que NÃO
    # suporta paginação de verdade — mandar resultOffset/resultRecordCount para
    # elas faz o servidor responder com o erro "Pagination is not supported".
    # Só usamos esses parâmetros quando a própria camada afirma suportar.
    suporta_paginacao = bool(
        meta.get("advancedQueryCapabilities", {}).get("supportsPagination")
    )

    # 2) Consulta (paginada, se suportado; senão, uma única chamada)
    endpoint = url_camada.rstrip("/") + "/query"
    registros = []
    offset = 0

    for _pagina in range(MAX_PAGINAS_ARCGIS):
        geometria_desta_chamada = "true"
        dados = None

        # Até 2 tentativas por página: a 1ª pedindo geometria, a 2ª (só se a
        # 1ª der erro 500) sem geometria — algumas camadas de polígono
        # complexas (ex.: limites de municípios) travam ao devolver a forma
        # geográfica, mas respondem normalmente só com os atributos.
        for tentativa in range(2):
            params = {
                "where": "1=1",
                "outFields": "*",
                "f": "json",
                "returnGeometry": geometria_desta_chamada,
            }
            if suporta_paginacao:
                params["resultOffset"] = offset
                params["resultRecordCount"] = TAMANHO_PAGINA_ARCGIS

            try:
                resp = requests.get(endpoint, params=params, headers=HEADERS, timeout=TIMEOUT)
                resp.raise_for_status()
                dados = resp.json()
            except (requests.RequestException, ValueError) as e:
                log.warning(f"  Falha ao consultar camada {nome_fonte}: {e}")
                dados = None
                break

            if dados.get("error"):
                erro = dados["error"]
                if tentativa == 0 and erro.get("code") == 500 and geometria_desta_chamada == "true":
                    log.info(
                        f"  {nome_fonte}: erro 500 com geometria — tentando "
                        f"novamente só com os atributos (sem geometria)"
                    )
                    geometria_desta_chamada = "false"
                    continue
                log.warning(f"  Camada {nome_fonte} retornou erro: {erro}")
                dados = None
            break

        if dados is None:
            break

        features = dados.get("features", [])
        if not features:
            break

        for feat in features:
            linha = dict(feat.get("attributes", {}))
            geom = feat.get("geometry") or {}
            if "x" in geom and "y" in geom:
                linha["_longitude"] = geom.get("x")
                linha["_latitude"] = geom.get("y")
            registros.append(linha)

        if not suporta_paginacao:
            # Sem suporte a paginação: só dá pra pegar a primeira leva
            # (até o MaxRecordCount do serviço). Avisa se pode haver mais.
            if dados.get("exceededTransferLimit"):
                log.info(
                    f"  Aviso: {nome_fonte} pode ter mais registros do que o "
                    f"limite de {len(features)} por chamada, mas este servidor "
                    f"não suporta paginação — baixando só o primeiro lote."
                )
            break

        if not dados.get("exceededTransferLimit"):
            break  # já veio tudo, não precisa paginar mais
        offset += len(features)

    if not registros:
        log.warning(f"  Camada {nome_fonte} não retornou registros.")
        return

    df = pd.DataFrame(registros)

    PASTA_BRUTOS.mkdir(parents=True, exist_ok=True)
    nome_arquivo = nome_arquivo_seguro(nome_fonte) + ".xlsx"
    destino = PASTA_BRUTOS / nome_arquivo
    df.to_excel(destino, index=False)

    aviso_truncamento = ""
    if len(registros) >= TAMANHO_PAGINA_ARCGIS * MAX_PAGINAS_ARCGIS:
        aviso_truncamento = " (atingiu o limite de páginas — pode haver mais dados)"
    log.info(f"  {len(df)} registro(s) salvos em {destino.name}{aviso_truncamento}")

    item = ItemEncontrado(
        fonte=nome_fonte,
        titulo=nome_fonte,
        formato="xlsx",
        url=url_camada,
        dataset_ou_pagina="ArcGIS REST Feature Layer (EPE)",
        caminho_local=str(destino),
        status_download=f"baixado com sucesso ({len(df)} registros via ArcGIS REST)",
    )
    ITENS_ENCONTRADOS.append(item)
    time.sleep(PAUSA_ENTRE_REQUISICOES)


# --------------------------------------------------------------------------
# COLETA NA EIA (U.S. Energy Information Administration)
# --------------------------------------------------------------------------

def _eia_get(caminho: str, params: dict) -> dict:
    params = dict(params)
    params["api_key"] = EIA_API_KEY
    resp = requests.get(f"{EIA_API_BASE}{caminho}", params=params, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def _salvar_tabela_eia(nome_fonte: str, registros: list, url_referencia: str):
    if not registros:
        log.warning(f"  {nome_fonte}: nenhum registro retornado.")
        return

    df = pd.DataFrame(registros)
    PASTA_BRUTOS.mkdir(parents=True, exist_ok=True)
    destino = PASTA_BRUTOS / (nome_arquivo_seguro(nome_fonte) + ".xlsx")
    df.to_excel(destino, index=False)
    log.info(f"  {len(df)} registro(s) salvos em {destino.name}")

    ITENS_ENCONTRADOS.append(ItemEncontrado(
        fonte=nome_fonte,
        titulo=nome_fonte,
        formato="xlsx",
        url=url_referencia,
        dataset_ou_pagina="EIA API v2 (international)",
        caminho_local=str(destino),
        status_download=f"baixado com sucesso ({len(df)} registros via EIA API)",
    ))


def buscar_em_eia():
    nome_fonte_base = "EIA (U.S. Energy Information Administration)"
    log.info(f"== Buscando dados sobre hidrogênio na {nome_fonte_base} ==")

    if not EIA_API_KEY:
        log.warning(
            "  Pulando EIA: essa fonte exige uma chave de API pessoal e gratuita. "
            "Gere a sua em https://www.eia.gov/opendata/register.php e defina a "
            "variável de ambiente EIA_API_KEY."
        )
        return

    # 1) Descobre quais productId da rota "international" mencionam hidrogênio
    try:
        meta = _eia_get("/international/data/facet/productId", {})
    except (requests.RequestException, ValueError) as e:
        log.warning(f"  Falha ao consultar catálogo de produtos da EIA: {e}")
        return

    produtos = meta.get("response", {}).get("facets", [])
    produtos_h2 = [
        p for p in produtos
        if "hydrogen" in (p.get("name") or "").lower() or "hidrog" in normalizar_texto(p.get("name") or "")
    ]

    if not produtos_h2:
        log.warning(
            "  A EIA não tem, hoje, nenhum produto específico de 'hidrogênio' "
            "catalogado na rota internacional consultada — nada para baixar aqui."
        )
        return

    log.info(f"  {len(produtos_h2)} produto(s) relacionados a hidrogênio encontrados na EIA.")

    for produto in produtos_h2:
        produto_id = produto.get("id")
        produto_nome = produto.get("name", str(produto_id))
        nome_fonte = f"{nome_fonte_base} - {produto_nome} (Brasil)"

        params = {
            "frequency": "annual",
            "data[0]": "value",
            "facets[countryRegionId][]": EIA_PAIS_FOCO,
            "facets[productId][]": produto_id,
            "sort[0][column]": "period",
            "sort[0][direction]": "desc",
            "length": 5000,
        }
        try:
            dados = _eia_get("/international/data/", params)
        except (requests.RequestException, ValueError) as e:
            log.warning(f"  Falha ao consultar '{produto_nome}' para o Brasil na EIA: {e}")
            continue

        registros = dados.get("response", {}).get("data", [])

        if not registros:
            # Sem dado específico para o Brasil — tenta trazer o comparativo
            # internacional completo desse produto, que ainda é útil de contexto
            log.info(f"  Sem dados do Brasil para '{produto_nome}' — buscando comparativo internacional.")
            params_sem_pais = dict(params)
            del params_sem_pais["facets[countryRegionId][]"]
            try:
                dados_intl = _eia_get("/international/data/", params_sem_pais)
                registros = dados_intl.get("response", {}).get("data", [])
                nome_fonte = f"{nome_fonte_base} - {produto_nome} (Comparativo internacional)"
            except (requests.RequestException, ValueError) as e:
                log.warning(f"  Falha ao consultar comparativo internacional de '{produto_nome}': {e}")
                continue

        url_referencia = f"{EIA_API_BASE}/international/data/?facets[productId][]={produto_id}"
        _salvar_tabela_eia(nome_fonte, registros, url_referencia)
        time.sleep(PAUSA_ENTRE_REQUISICOES)


# --------------------------------------------------------------------------
# COLETA EM PÁGINAS HTML (EPE, Portal do Hidrogênio, etc.)
# --------------------------------------------------------------------------

def _url_html_candidata(url: str) -> bool:
    parsed = urlparse(url)
    caminho = parsed.path.lower()

    if not parsed.scheme.startswith("http"):
        return False

    if any(caminho.endswith(ext) for ext in EXTENSOES_DADOS + EXTENSOES_EXTRA):
        return False

    if any(p in caminho for p in ("/login", "/logout", "/entrar", "/cadastro", "/cookie")):
        return False

    return True


def _link_interno_relevante(url: str, texto: str) -> bool:
    conjunto = normalizar_texto(f"{texto} {url}")
    return any(normalizar_texto(palavra) in conjunto for palavra in CAMINHOS_RELEVANTES_CRAWLER)


def _adicionar_item_arquivo(nome_fonte: str, titulo: str, url: str, url_pagina: str):
    ext = os.path.splitext(urlparse(url).path)[1].lower()

    if not ext:
        return False

    eh_dado = ext in EXTENSOES_DADOS
    eh_extra = ext in EXTENSOES_EXTRA

    if not (eh_dado or eh_extra):
        return False

    if not contexto_e_relevante(titulo, url, url_pagina):
        return False

    chave = (normalizar_texto(nome_fonte), url.rstrip("/").lower())
    if any(
        normalizar_texto(item.fonte) == chave[0]
        and item.url.rstrip("/").lower() == chave[1]
        for item in ITENS_ENCONTRADOS
    ):
        return False

    item = ItemEncontrado(
        fonte=nome_fonte,
        titulo=titulo,
        formato=ext.lstrip("."),
        url=url,
        dataset_ou_pagina=url_pagina,
    )
    ITENS_ENCONTRADOS.append(item)
    return True


def buscar_em_pagina_html(nome_fonte: str, url_pagina: str):
    log.info(f"== Varrendo página HTML e subpáginas: {nome_fonte} ({url_pagina}) ==")

    dominio_base = urlparse(url_pagina).netloc.lower()
    fila = [(url_pagina, 0)]
    visitadas = set()
    encontrados = 0

    while fila and len(visitadas) < MAX_PAGINAS_CRAWLER_POR_FONTE:
        url_atual, profundidade = fila.pop(0)
        url_atual = url_atual.split("#")[0].rstrip("/")

        if not url_atual:
            continue

        if url_atual in visitadas:
            continue

        if profundidade > MAX_PROFUNDIDADE_CRAWLER:
            continue

        parsed_atual = urlparse(url_atual)
        if parsed_atual.netloc.lower() != dominio_base:
            continue

        visitadas.add(url_atual)

        log.info(
            f"  Página {len(visitadas)}/{MAX_PAGINAS_CRAWLER_POR_FONTE} "
            f"(profundidade {profundidade}): {url_atual}"
        )

        resp = None
        for tentativa in range(2):
            try:
                timeout_usado = TIMEOUT if tentativa == 0 else TIMEOUT * 2
                resp = requests.get(
                    url_atual,
                    headers=HEADERS,
                    timeout=timeout_usado,
                )
                resp.raise_for_status()
                break
            except requests.RequestException as e:
                if tentativa == 0:
                    log.info(f"  Tentativa 1 falhou ({e}); tentando novamente...")
                    continue
                log.warning(f"  Falha ao acessar {url_atual}: {e}")
                resp = None

        if resp is None:
            continue

        content_type = (resp.headers.get("Content-Type") or "").lower()

        if "text/html" not in content_type and "application/xhtml" not in content_type:
            ext = os.path.splitext(urlparse(url_atual).path)[1].lower()
            if ext in EXTENSOES_DADOS + EXTENSOES_EXTRA:
                titulo = os.path.basename(urlparse(url_atual).path) or url_atual
                if _adicionar_item_arquivo(nome_fonte, titulo, url_atual, url_pagina):
                    encontrados += 1
            continue

        soup = BeautifulSoup(resp.text, "html.parser")

        for link in soup.find_all("a", href=True):
            href = link["href"].strip()

            if not href or href.startswith(("#", "mailto:", "javascript:", "tel:")):
                continue

            url_completa = urljoin(url_atual, href).split("#")[0]
            parsed_link = urlparse(url_completa)

            if parsed_link.netloc.lower() != dominio_base:
                continue

            texto_link = (
                link.get_text(" ", strip=True)
                or os.path.basename(parsed_link.path)
                or url_completa
            )

            ext = os.path.splitext(parsed_link.path)[1].lower()

            if ext in EXTENSOES_DADOS + EXTENSOES_EXTRA:
                if _adicionar_item_arquivo(
                    nome_fonte,
                    texto_link,
                    url_completa,
                    url_atual,
                ):
                    encontrados += 1
                continue

            if (
                profundidade < MAX_PROFUNDIDADE_CRAWLER
                and SEGUIR_LINKS_SEM_EXTENSAO
                and _url_html_candidata(url_completa)
                and _link_interno_relevante(url_completa, texto_link)
            ):
                if url_completa.rstrip("/") not in visitadas:
                    fila.append((url_completa, profundidade + 1))

        time.sleep(PAUSA_ENTRE_REQUISICOES)

    log.info(
        f"  {len(visitadas)} página(s) HTML visitada(s) e "
        f"{encontrados} arquivo(s) relevante(s) descoberto(s)"
    )


# --------------------------------------------------------------------------
# DOWNLOAD DOS ITENS ENCONTRADOS
# --------------------------------------------------------------------------

def baixar_itens_encontrados():
    log.info(f"== Baixando {len(ITENS_ENCONTRADOS)} arquivo(s) encontrados ==")
    PASTA_BRUTOS.mkdir(parents=True, exist_ok=True)

    for i, item in enumerate(ITENS_ENCONTRADOS, start=1):
        if item.caminho_local:
            # Já foi obtido por outro meio (ex.: consulta direta ao ArcGIS REST / EIA)
            continue

        ext = item.formato if item.formato.startswith(".") else f".{item.formato}"
        if not ext or ext == ".":
            ext = os.path.splitext(urlparse(item.url).path)[1] or ".dat"

        if ext == ".pdf" and not BAIXAR_PDFS_TAMBEM:
            item.status_download = "não baixado (PDF ignorado por configuração)"
            continue

        # A IEA exige login para baixar seus arquivos — tentar programaticamente
        # sempre dá 403. Em vez de falhar feio, já orientamos o usuário.
        if "iea.org" in item.url.lower():
            item.status_download = (
                "não baixado automaticamente: a IEA exige login manual. "
                f"Baixe em {item.url} e coloque o arquivo em "
                f"{PASTA_ENTRADA_MANUAL}/ para ele ser incorporado na próxima execução."
            )
            continue

        nome_base = nome_arquivo_seguro(f"{i:03d}_{item.titulo}")
        destino = PASTA_BRUTOS / f"{nome_base}{ext}"

        log.info(f"  [{i}/{len(ITENS_ENCONTRADOS)}] {item.titulo[:70]}")
        ok = baixar_arquivo(item.url, destino)

        if ok:
            item.caminho_local = str(destino)
            item.status_download = "baixado com sucesso"

            if ext == ".csv":
                convertido = converter_csv_para_xlsx(destino)
                if convertido:
                    item.status_download += f" | convertido para {convertido.name}"
        else:
            item.status_download = "falha no download"

        time.sleep(PAUSA_ENTRE_REQUISICOES)


# --------------------------------------------------------------------------
# ENTRADA MANUAL (ex.: exportações da IEA que exigem login para baixar)
# --------------------------------------------------------------------------

def importar_entrada_manual():
    """Varre PASTA_ENTRADA_MANUAL em busca de planilhas que o usuário tenha
    baixado manualmente (ex.: a Hydrogen Projects Database da IEA, que exige
    login) e as incorpora ao índice final, sem precisar reprocessar."""
    if not PASTA_ENTRADA_MANUAL.exists():
        PASTA_ENTRADA_MANUAL.mkdir(parents=True, exist_ok=True)
        return

    arquivos = [
        f for f in PASTA_ENTRADA_MANUAL.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSOES_DADOS
    ]

    if not arquivos:
        return

    log.info(f"== Incorporando {len(arquivos)} arquivo(s) de dados adicionados manualmente ==")
    for f in arquivos:
        item = ItemEncontrado(
            fonte="Entrada manual (usuário)",
            titulo=f.stem,
            formato=f.suffix.lstrip("."),
            url="",
            dataset_ou_pagina=str(f),
            caminho_local=str(f),
            status_download="adicionado manualmente pelo usuário",
        )
        ITENS_ENCONTRADOS.append(item)
        log.info(f"  Incorporado: {f.name}")


# --------------------------------------------------------------------------
# PADRONIZAÇÃO — organiza os projetos no esquema de variáveis do TCC
# (Seção 3.5: Identificação, Localização, Técnicas, Econômicas, Maturidade,
# Aplicação final)
# --------------------------------------------------------------------------

# Mapeamento de {nome_da_camada: {campo_padrao: nome_da_coluna_na_fonte}}.
# A chave deve bater EXATAMENTE com o nome da camada dentro do serviço ArcGIS
# (a parte depois do último " - " em item.fonte, ex.: "EPE - H2_PROJETOS_BIO -
# Projetos" -> camada "Projetos"). Usar o nome do SERVIÇO aqui seria perigoso:
# "H2_PROJETOS_BIO" também tem uma camada "Agro_Bio" (potencial de resíduos,
# não são projetos!) que ficaria erroneamente incluída por um match parcial.
# Fontes não listadas aqui não entram na base padronizada (mas continuam
# disponíveis "cruas" em dados_h2/brutos/).
MAPEAMENTO_CAMPOS_PADRAO = {
    "projetos": {  # camada "Projetos" da EPE (H2_PROJETOS_202410, H2_PROJETOS_BIO etc.)
        "nome_projeto": "Nome",
        "capacidade": "Capacidade",
        "aplicacao_final": "Finalidade",
        "status_maturidade": "Estagio",
        "localizacao_bruta": "Local",
        "investimento": "valor",
        "latitude": "_latitude",
        "longitude": "_longitude",
    },
}


def _nome_da_camada(fonte: str) -> str:
    """Extrai o nome da camada a partir do nome completo da fonte, ex.:
    'EPE - H2_PROJETOS_BIO - Projetos' -> 'Projetos'."""
    partes = fonte.split(" - ")
    return partes[-1].strip() if partes else fonte


CAMPOS_PADRAO_ORDEM = [
    # --- Identificação ---
    "fonte_dado", "nome_projeto", "empresa_consorcio",
    # --- Localização ---
    "pais", "estado_regiao", "cidade_porto_hub", "localizacao_bruta", "latitude", "longitude",
    # --- Características técnicas ---
    "capacidade", "tecnologia_eletrolise", "fonte_renovavel", "produto_final",
    # --- Características econômicas ---
    "investimento", "tipo_parceria",
    # --- Maturidade ---
    "status_maturidade", "fase", "previsao_operacao",
    # --- Aplicação final ---
    "aplicacao_final",
]


def gerar_base_projetos_padronizada():
    """Lê os arquivos brutos já baixados cuja fonte tem um mapeamento
    conhecido (MAPEAMENTO_CAMPOS_PADRAO) e monta uma tabela única no esquema
    de variáveis definido na Seção 3.5 do plano de trabalho (identificação,
    localização, características técnicas, econômicas, maturidade e
    aplicação final) — pronta para alimentar o pipeline de ETL/Power BI."""
    log.info("== Gerando base padronizada de projetos (esquema da Seção 3.5) ==")

    linhas = []
    for item in ITENS_ENCONTRADOS:
        if not item.caminho_local or not item.caminho_local.endswith(".xlsx"):
            continue

        mapeamento = None
        nome_camada = normalizar_texto(_nome_da_camada(item.fonte))
        for chave, campos in MAPEAMENTO_CAMPOS_PADRAO.items():
            if normalizar_texto(chave) == nome_camada:
                mapeamento = campos
                break

        if not mapeamento:
            continue

        try:
            df_bruto = pd.read_excel(item.caminho_local)
        except Exception as e:
            log.warning(f"  Não foi possível ler {item.caminho_local}: {e}")
            continue

        for _, linha_bruta in df_bruto.iterrows():
            linha_padrao = {campo: None for campo in CAMPOS_PADRAO_ORDEM}
            linha_padrao["fonte_dado"] = item.fonte
            linha_padrao["pais"] = "Brasil"  # fontes mapeadas aqui são nacionais

            for campo_padrao, coluna_origem in mapeamento.items():
                if coluna_origem in df_bruto.columns:
                    linha_padrao[campo_padrao] = linha_bruta.get(coluna_origem)

            linhas.append(linha_padrao)

    if not linhas:
        log.warning(
            "  Nenhuma fonte com mapeamento conhecido foi baixada ainda — "
            "base padronizada não gerada. Adicione mapeamentos em "
            "MAPEAMENTO_CAMPOS_PADRAO conforme novas fontes forem incluídas."
        )
        return

    df_padrao = pd.DataFrame(linhas, columns=CAMPOS_PADRAO_ORDEM)

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    destino = PASTA_SAIDA / "base_projetos_padronizada.xlsx"
    df_padrao.to_excel(destino, index=False)
    log.info(f"  {len(df_padrao)} projeto(s) padronizado(s) salvos em {destino.name}")
    return df_padrao



# --------------------------------------------------------------------------
# PADRONIZAÇÃO AMPLIADA — projetos H2V de múltiplas fontes
# --------------------------------------------------------------------------

CAMPOS_H2V = [
    "id_projeto", "fonte_dado", "nome_projeto", "empresa_consorcio",
    "parceiros", "pais", "estado_regiao", "cidade_porto_hub",
    "localizacao_bruta", "latitude", "longitude", "capacidade_mw",
    "producao_t_ano", "producao_kg_dia", "tecnologia_eletrolise",
    "fonte_renovavel", "produto_final", "materia_prima", "investimento",
    "moeda_investimento", "tipo_parceria", "status_maturidade", "fase",
    "data_anuncio", "previsao_operacao", "aplicacao_final", "descricao",
    "url_fonte", "arquivo_origem"
]

ALIASES_H2V = {
    "nome_projeto": ["nome", "projeto", "nome projeto", "project", "project name", "titulo", "title", "empreendimento"],
    "empresa_consorcio": ["empresa", "company", "developer", "owner", "desenvolvedor", "responsavel", "promotor"],
    "parceiros": ["parceiros", "partners", "socios", "sócios", "consorcio"],
    "pais": ["pais", "país", "country"],
    "estado_regiao": ["estado", "uf", "state", "regiao", "região", "region"],
    "cidade_porto_hub": ["municipio", "município", "cidade", "city", "municipality", "porto", "hub"],
    "localizacao_bruta": ["local", "localizacao", "localização", "location", "ubicacion", "endereco", "endereço"],
    "latitude": ["latitude", "lat", "_latitude"],
    "longitude": ["longitude", "long", "lon", "_longitude"],
    "capacidade_mw": ["capacidade", "capacity", "capacity mw", "potencia", "potência", "installed capacity", "potencia instalada", "capacidade de eletrolise", "capacidade de eletrólise", "electrolysis capacity", "electrolyzer capacity", "potencia eletrolisador", "potência eletrolisador"],
    "producao_t_ano": ["producao", "produção", "production", "production capacity", "capacidade de producao", "capacidade de produção", "capacidad de produccion", "t/ano", "t ano", "toneladas por ano", "tonnes per year"],
    "producao_kg_dia": ["kg/dia", "kg dia", "kg/day", "kg day", "production kg/day"],
    "tecnologia_eletrolise": ["tecnologia", "technology", "tipo de eletrolisador", "tipo de electrolizador", "electrolyzer type", "electrolyser type", "eletrolisador"],
    "fonte_renovavel": ["fonte renovavel", "fonte renovável", "fonte de energia", "fonte de alimentação", "fuente de alimentacion", "power source", "energy source", "renewable source"],
    "produto_final": ["produto", "produto final", "product", "output", "derivado"],
    "materia_prima": ["materia prima", "matéria prima", "feedstock", "raw material"],
    "investimento": ["investimento", "investment", "inversion", "valor", "investment value"],
    "moeda_investimento": ["moeda", "currency"],
    "tipo_parceria": ["tipo parceria", "partnership", "partnership type", "modelo de negocio", "modelo de negócio"],
    "status_maturidade": ["status", "situacao", "situação", "project status", "maturity", "maturidade"],
    "fase": ["fase", "stage", "phase", "estagio", "estágio"],
    "data_anuncio": ["data anuncio", "data de anuncio", "announcement date", "announcement"],
    "previsao_operacao": ["previsao operacao", "previsão operação", "data operacao", "data de operacao", "operation date", "inauguracao", "inauguración"],
    "aplicacao_final": ["finalidade", "aplicacao", "aplicação", "uso", "use", "end use", "application"],
    "descricao": ["descricao", "descrição", "description", "descripcion", "detalhes", "details", "project description"]
}

def _coluna_h2v(colunas, aliases):
    norm={c: normalizar_texto(str(c).replace('_',' ').replace('-',' ')).strip() for c in colunas}
    als=[normalizar_texto(a).strip() for a in aliases]
    for a in als:
        for c,n in norm.items():
            if n==a:
                return c
    for a in als:
        for c,n in norm.items():
            if a in n or n in a:
                return c
    return None

def _numero_h2v(valor):
    if valor is None or (isinstance(valor,float) and pd.isna(valor)):
        return None
    m=re.search(r'-?\d+(?:[.,]\d+)?', str(valor))
    if not m:
        return None
    x=m.group(0)
    if ',' in x and '.' in x:
        x=x.replace('.','').replace(',','.') if x.rfind(',')>x.rfind('.') else x.replace(',','')
    else:
        x=x.replace(',','.')
    try: return float(x)
    except ValueError: return None

def _producao_t_ano(valor):
    if valor is None: return None
    x=_numero_h2v(valor)
    if x is None: return None
    t=normalizar_texto(str(valor))
    if 'kg/dia' in t or 'kg dia' in t or 'kg/day' in t: return x*365/1000
    if 'kg/ano' in t or 'kg ano' in t: return x/1000
    if 't/dia' in t or 'ton dia' in t: return x*365
    return x

def gerar_base_h2v_multifonte():
    log.info("== Padronizando projetos H2V de todas as fontes coletadas ==")
    frames=[]
    for item in ITENS_ENCONTRADOS:
        if not item.caminho_local: continue
        caminho=Path(item.caminho_local)
        if not caminho.exists() or caminho.suffix.lower() not in ('.csv','.xlsx','.xls','.json'): continue
        try:
            if caminho.suffix.lower()=='.csv':
                try: tabelas=[pd.read_csv(caminho,sep=None,engine='python',encoding='utf-8')]
                except UnicodeDecodeError: tabelas=[pd.read_csv(caminho,sep=None,engine='python',encoding='latin-1')]
            elif caminho.suffix.lower() in ('.xlsx','.xls'):
                xl=pd.ExcelFile(caminho); tabelas=[xl.parse(a) for a in xl.sheet_names]
            else: tabelas=[pd.read_json(caminho)]
        except Exception as e:
            log.warning(f"  Falha ao ler {caminho.name}: {e}"); continue
        for df in tabelas:
            if df.empty: continue
            out=pd.DataFrame(index=df.index)
            for campo in CAMPOS_H2V: out[campo]=None
            for campo,aliases in ALIASES_H2V.items():
                c=_coluna_h2v(df.columns,aliases)
                if c is not None: out[campo]=df[c]
            out['fonte_dado']=item.fonte; out['url_fonte']=item.url; out['arquivo_origem']=str(caminho)
            out['pais']=out['pais'].fillna('Brasil')
            texto=df.fillna('').astype(str).agg(' '.join,axis=1).map(normalizar_texto)
            if out['status_maturidade'].isna().all():
                out['status_maturidade']=texto.map(lambda x: 'Em operação' if ('operacao' in x or 'operational' in x) else 'Em construção' if 'construcao' in x or 'construction' in x else 'Em desenvolvimento' if 'desenvolvimento' in x or 'development' in x else 'Planejado' if 'planejado' in x or 'planned' in x else None)
            out['capacidade_mw']=out['capacidade_mw'].map(_numero_h2v)
            out['producao_t_ano']=out['producao_t_ano'].map(_producao_t_ano)
            out['producao_kg_dia']=out['producao_kg_dia'].map(_numero_h2v)
            out['investimento']=out['investimento'].map(_numero_h2v)
            out['latitude']=out['latitude'].map(_numero_h2v); out['longitude']=out['longitude'].map(_numero_h2v)
            out=out[out[['nome_projeto','empresa_consorcio','localizacao_bruta','capacidade_mw','producao_t_ano']].notna().any(axis=1)]
            if not out.empty: frames.append(out)
    if not frames:
        log.warning('  Nenhum projeto estruturado encontrado nas fontes coletadas.'); return None
    df=pd.concat(frames,ignore_index=True)
    df['nome_projeto']=df['nome_projeto'].fillna('Projeto não identificado')
    df=df.drop_duplicates(subset=['nome_projeto','empresa_consorcio','estado_regiao','cidade_porto_hub'],keep='first').reset_index(drop=True)
    df['id_projeto']=[f'H2V-{i:05d}' for i in range(1,len(df)+1)]
    destino=PASTA_SAIDA/'base_projetos_h2v_multifonte.xlsx'; csv=PASTA_SAIDA/'base_projetos_h2v_multifonte.csv'
    df[CAMPOS_H2V].to_excel(destino,index=False); df[CAMPOS_H2V].to_csv(csv,index=False,encoding='utf-8-sig')
    log.info(f'  {len(df)} projeto(s) H2V padronizado(s): {destino.name}')
    return df[CAMPOS_H2V]

def gerar_modelo_estrela_h2v(df):
    if df is None or df.empty: return
    pasta=PASTA_SAIDA/'modelo_estrela'; pasta.mkdir(parents=True,exist_ok=True)
    def dim(cols,nome):
        d=df[cols].drop_duplicates().reset_index(drop=True); d.insert(0,'id_'+nome,range(1,len(d)+1)); d.to_csv(pasta/f'dim_{nome}.csv',index=False,encoding='utf-8-sig'); return d
    fontes=dim(['fonte_dado'],'fontes'); estados=dim(['pais','estado_regiao'],'estados'); empresas=dim(['empresa_consorcio'],'empresas'); status=dim(['status_maturidade','fase'],'status'); tecnologias=dim(['tecnologia_eletrolise'],'tecnologias'); aplicacoes=dim(['aplicacao_final','produto_final'],'aplicacoes'); localizacoes=dim(['pais','estado_regiao','cidade_porto_hub','localizacao_bruta','latitude','longitude'],'localizacao')
    fato=df[['id_projeto','fonte_dado','pais','estado_regiao','empresa_consorcio','status_maturidade','fase','tecnologia_eletrolise','aplicacao_final','produto_final','capacidade_mw','producao_t_ano','producao_kg_dia','investimento','moeda_investimento','data_anuncio','previsao_operacao','fonte_renovavel','materia_prima','tipo_parceria','url_fonte']].copy()
    fato=fato.merge(fontes,on=['fonte_dado']).merge(estados,on=['pais','estado_regiao']).merge(empresas,on=['empresa_consorcio']).merge(status,on=['status_maturidade','fase']).merge(tecnologias,on=['tecnologia_eletrolise']).merge(aplicacoes,on=['aplicacao_final','produto_final'])
    fato.to_csv(pasta/'fato_projetos_h2v.csv',index=False,encoding='utf-8-sig')
    with pd.ExcelWriter(pasta/'modelo_estrela_h2v.xlsx',engine='openpyxl') as w:
        fato.to_excel(w,sheet_name='Fato_Projetos_H2V',index=False); fontes.to_excel(w,sheet_name='Dim_Fontes',index=False); estados.to_excel(w,sheet_name='Dim_Estados',index=False); empresas.to_excel(w,sheet_name='Dim_Empresas',index=False); status.to_excel(w,sheet_name='Dim_Status',index=False); tecnologias.to_excel(w,sheet_name='Dim_Tecnologias',index=False); aplicacoes.to_excel(w,sheet_name='Dim_Aplicacoes',index=False); localizacoes.to_excel(w,sheet_name='Dim_Localizacao',index=False)
    log.info(f'  Modelo estrela gerado em {pasta}/')

# --------------------------------------------------------------------------
# RESOLUÇÃO DE ENTIDADES — Algoritmo 1 do TCC (blocking + similaridade de
# Levenshtein) para eliminar duplicatas do mesmo projeto vindo de fontes
# distintas (ex.: EPE + IEA + H2LAC citando o mesmo empreendimento).
# --------------------------------------------------------------------------

# Limiar de similaridade (tau) para considerar dois registros como o mesmo
# projeto — Seção 2.2.2 / Equação 3 do TCC. Pares com similaridade entre
# LIMIAR_RELATORIO e LIMIAR_SIMILARIDADE_ENTIDADES não são fundidos
# automaticamente, mas aparecem no relatório de auditoria para você revisar
# manualmente (é a base do "conjunto-ouro rotulado manualmente" da Seção 3.7).
LIMIAR_SIMILARIDADE_ENTIDADES = 0.85
LIMIAR_RELATORIO = 0.65

# Campo usado para "blocking" (Seção 2.2.2, Equação 2): só comparamos pares de
# registros que caem no mesmo bloco, para não fazer O(n²) comparações à toa.
CAMPO_BLOCKING = "localizacao_bruta"


def _distancia_levenshtein(a: str, b: str) -> int:
    """Distância de edição clássica (programação dinâmica), sem depender de
    bibliotecas externas — usada na Equação 3 (similaridade normalizada)."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)

    linha_anterior = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        linha_atual = [i] + [0] * len(b)
        for j, cb in enumerate(b, start=1):
            custo = 0 if ca == cb else 1
            linha_atual[j] = min(
                linha_anterior[j] + 1,       # remoção
                linha_atual[j - 1] + 1,      # inserção
                linha_anterior[j - 1] + custo,  # substituição
            )
        linha_anterior = linha_atual
    return linha_anterior[-1]


def similaridade_levenshtein(a: str, b: str) -> float:
    """sim(a,b) = 1 - lev(a,b) / max(|a|,|b|) — Equação 3 do TCC."""
    a, b = normalizar_texto(a or ""), normalizar_texto(b or "")
    if not a and not b:
        return 1.0
    maior = max(len(a), len(b))
    if maior == 0:
        return 1.0
    return 1 - (_distancia_levenshtein(a, b) / maior)


def _bloco_de(valor) -> str:
    """Chave de bloco simplificada: normaliza e usa os 4 primeiros caracteres
    do campo de localização (ex.: 'Pecém/CE' e 'Pecem - CE' caem no mesmo
    bloco). Registros sem esse campo vão para um bloco único 'sem_bloco'."""
    texto = normalizar_texto(str(valor)) if valor and str(valor).strip() else ""
    return texto[:4] if texto else "sem_bloco"


class _UniaoBusca:
    """Estrutura union-find simples para agrupar pares equivalentes (M) em
    clusters de entidades — usada para materializar o resultado do
    Algoritmo 1 (que devolve pares) em grupos finais."""

    def __init__(self, n: int):
        self.pai = list(range(n))

    def encontrar(self, x: int) -> int:
        while self.pai[x] != x:
            self.pai[x] = self.pai[self.pai[x]]
            x = self.pai[x]
        return x

    def unir(self, x: int, y: int):
        rx, ry = self.encontrar(x), self.encontrar(y)
        if rx != ry:
            self.pai[rx] = ry


def resolver_entidades(df_padrao: pd.DataFrame):
    """Implementa o Algoritmo 1 do TCC: agrupa registros de df_padrao por
    bloco, compara pares dentro de cada bloco pela similaridade de nome
    (Equação 3) e funde os que ultrapassam LIMIAR_SIMILARIDADE_ENTIDADES.

    Gera dois arquivos:
    - base_projetos_deduplicada.xlsx: uma linha por entidade única, com as
      fontes de origem concatenadas em 'fontes_conciliadas'.
    - pares_similares_resolucao_entidades.xlsx: relatório de auditoria com
      todos os pares comparados acima de LIMIAR_RELATORIO (para revisão
      manual / construção do conjunto-ouro da Seção 3.7).
    """
    log.info("== Executando resolução de entidades (Algoritmo 1: blocking + Levenshtein) ==")

    if df_padrao is None or df_padrao.empty:
        log.warning("  Base padronizada vazia — resolução de entidades não executada.")
        return

    df = df_padrao.reset_index(drop=True).copy()
    n = len(df)

    # 1) Blocking (Equação 2): agrupa índices por chave de bloco
    blocos: dict[str, list[int]] = {}
    for idx, valor in enumerate(df[CAMPO_BLOCKING]):
        chave = _bloco_de(valor)
        blocos.setdefault(chave, []).append(idx)

    uniao = _UniaoBusca(n)
    pares_relatorio = []
    total_comparacoes = 0

    # 2) Dentro de cada bloco, compara todos os pares (C_bloco da Equação 2)
    for chave_bloco, indices in blocos.items():
        for a in range(len(indices)):
            for b in range(a + 1, len(indices)):
                i, j = indices[a], indices[b]
                total_comparacoes += 1
                nome_i = df.at[i, "nome_projeto"] or ""
                nome_j = df.at[j, "nome_projeto"] or ""
                sim = similaridade_levenshtein(nome_i, nome_j)

                if sim >= LIMIAR_RELATORIO:
                    pares_relatorio.append({
                        "bloco": chave_bloco,
                        "indice_a": i, "nome_a": nome_i, "fonte_a": df.at[i, "fonte_dado"],
                        "indice_b": j, "nome_b": nome_j, "fonte_b": df.at[j, "fonte_dado"],
                        "similaridade": round(sim, 4),
                        "fundido": sim >= LIMIAR_SIMILARIDADE_ENTIDADES,
                    })

                if sim >= LIMIAR_SIMILARIDADE_ENTIDADES:
                    uniao.unir(i, j)

    log.info(
        f"  {total_comparacoes} comparação(ões) de pares dentro dos blocos "
        f"({len(blocos)} bloco(s) de '{CAMPO_BLOCKING}')."
    )

    # 3) Materializa os clusters (grupos de entidades) a partir do union-find
    grupos: dict[int, list[int]] = {}
    for idx in range(n):
        raiz = uniao.encontrar(idx)
        grupos.setdefault(raiz, []).append(idx)

    linhas_dedup = []
    for grupo_id, indices in enumerate(grupos.values(), start=1):
        sub = df.loc[indices]
        linha = {}
        for coluna in df.columns:
            valores_nao_nulos = sub[coluna].dropna()
            linha[coluna] = valores_nao_nulos.iloc[0] if not valores_nao_nulos.empty else None
        linha["grupo_entidade"] = grupo_id
        linha["qtd_registros_conciliados"] = len(indices)
        linha["fontes_conciliadas"] = "; ".join(sorted(set(sub["fonte_dado"].dropna().astype(str))))
        linhas_dedup.append(linha)

    colunas_finais = list(df.columns) + [
        "grupo_entidade", "qtd_registros_conciliados", "fontes_conciliadas"
    ]
    df_dedup = pd.DataFrame(linhas_dedup, columns=colunas_finais)

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    destino_dedup = PASTA_SAIDA / "base_projetos_deduplicada.xlsx"
    df_dedup.to_excel(destino_dedup, index=False)

    duplicatas_fundidas = n - len(df_dedup)
    log.info(
        f"  {n} registro(s) originais -> {len(df_dedup)} entidade(s) únicas "
        f"({duplicatas_fundidas} duplicata(s) fundida(s)). Salvo em {destino_dedup.name}"
    )

    if pares_relatorio:
        destino_relatorio = PASTA_SAIDA / "pares_similares_resolucao_entidades.xlsx"
        pd.DataFrame(pares_relatorio).sort_values(
            "similaridade", ascending=False
        ).to_excel(destino_relatorio, index=False)
        log.info(
            f"  {len(pares_relatorio)} par(es) com similaridade >= {LIMIAR_RELATORIO} "
            f"registrados em {destino_relatorio.name} (auditoria/conjunto-ouro)."
        )
    else:
        log.info(f"  Nenhum par com similaridade >= {LIMIAR_RELATORIO} encontrado.")

    return df_dedup


# --------------------------------------------------------------------------
# CÁLCULO DO LCOH — Equações 5-8 e Tabela 2 do TCC (Custo Nivelado do
# Hidrogênio por cenário e tecnologia de eletrólise)
# --------------------------------------------------------------------------

# Parâmetros de referência da Tabela 2 do TCC (faixas da literatura,
# IRENA 2020 / CCEE 2024). Uso o PONTO MÉDIO de cada faixa como valor de
# cálculo — ajuste aqui se quiser usar os extremos (mínimo/máximo) em vez
# da média, ou se publicar novos parâmetros nacionais específicos.
PARAMETROS_LCOH = {
    "ALK": {
        "capex_kw_2020": (650, 1000),
        "capex_kw_2030": (250, 450),
        "eta_el": (51.2, 51.2),       # kWh/kg (PCI)
        "f_om": (0.02, 0.04),         # fração do CAPEX/ano
        "r": 0.08,                    # taxa de desconto
        "n": 20,                      # vida útil (anos)
    },
    "PEM": {
        "capex_kw_2020": (700, 1400),
        "capex_kw_2030": (300, 700),
        "eta_el": (51.2, 55.0),
        "f_om": (0.02, 0.04),
        "r": 0.08,
        "n": 20,
    },
}

# Preço da eletricidade renovável por cenário (USD/MWh) — Tabela 2.
PRECO_ENERGIA_CENARIOS = {
    "conservador": (60, 70),
    "otimista": (25, 35),
}

# Fator de capacidade da planta (FC) — não especificado explicitamente na
# Tabela 2 do TCC; assumido aqui como 50% (valor típico de referência para
# eletrolisadores acoplados a fontes eólica+solar complementares). AJUSTE
# este valor se você tiver uma referência mais específica para o Brasil.
FATOR_CAPACIDADE_PADRAO = 0.50

ANOS_LCOH = (2020, 2030)


def _media(faixa) -> float:
    if isinstance(faixa, tuple):
        return (faixa[0] + faixa[1]) / 2
    return faixa


def calcular_lcoh_referencia(fator_capacidade: float = FATOR_CAPACIDADE_PADRAO):
    """Calcula o LCOH de referência (Equações 5-8) para cada combinação de
    tecnologia (ALK/PEM), cenário (conservador/otimista) e ano (2020/2030),
    usando os parâmetros da Tabela 2. Gera 'lcoh_projecao_brasil.xlsx',
    pronta para a medida DAX 'LCOH Médio ($/kg)' citada no TCC."""
    log.info("== Calculando LCOH de referência (Equações 5-8, Tabela 2) ==")

    linhas = []
    for tecnologia, params in PARAMETROS_LCOH.items():
        r = params["r"]
        n = params["n"]
        crf = (r * (1 + r) ** n) / ((1 + r) ** n - 1)  # Equação 8

        for ano in ANOS_LCOH:
            capex_kw = _media(params[f"capex_kw_{ano}"])
            eta_el = _media(params["eta_el"])
            f_om = _media(params["f_om"])

            for cenario, preco_faixa in PRECO_ENERGIA_CENARIOS.items():
                p_en = _media(preco_faixa)

                # Equação 7: componente de energia (USD/kg)
                componente_energia = eta_el * p_en * 1e-3

                # Equação 8: componente de capital (USD/kg)
                componente_capex = (
                    (crf + f_om) * capex_kw * eta_el
                ) / (8760 * fator_capacidade)

                # Equação 6: LCOH total
                lcoh = componente_capex + componente_energia

                linhas.append({
                    "tecnologia": tecnologia,
                    "cenario": cenario,
                    "ano": ano,
                    "capex_kw_usd": round(capex_kw, 2),
                    "eta_el_kwh_kg": round(eta_el, 2),
                    "f_om_fracao": round(f_om, 4),
                    "r_taxa_desconto": r,
                    "n_anos": n,
                    "crf": round(crf, 5),
                    "p_en_usd_mwh": round(p_en, 2),
                    "fator_capacidade": fator_capacidade,
                    "componente_capex_usd_kg": round(componente_capex, 4),
                    "componente_energia_usd_kg": round(componente_energia, 4),
                    "lcoh_usd_kg": round(lcoh, 4),
                })

    df_lcoh = pd.DataFrame(linhas)

    # Redução percentual e taxa anual composta entre o primeiro e o último
    # ano calculado, por tecnologia e cenário (fórmula da Seção 3.6 do TCC).
    resumo = []
    for (tecnologia, cenario), grupo in df_lcoh.groupby(["tecnologia", "cenario"]):
        grupo_ordenado = grupo.sort_values("ano")
        lcoh_inicial = grupo_ordenado.iloc[0]["lcoh_usd_kg"]
        lcoh_final = grupo_ordenado.iloc[-1]["lcoh_usd_kg"]
        ano_inicial = grupo_ordenado.iloc[0]["ano"]
        ano_final = grupo_ordenado.iloc[-1]["ano"]

        reducao_pct = (lcoh_inicial - lcoh_final) / lcoh_inicial * 100 if lcoh_inicial else None
        if lcoh_inicial and lcoh_final and ano_final != ano_inicial:
            taxa_anual = 1 - (lcoh_final / lcoh_inicial) ** (1 / (ano_final - ano_inicial))
        else:
            taxa_anual = None

        resumo.append({
            "tecnologia": tecnologia,
            "cenario": cenario,
            f"lcoh_{ano_inicial}_usd_kg": round(lcoh_inicial, 4),
            f"lcoh_{ano_final}_usd_kg": round(lcoh_final, 4),
            "reducao_percentual": round(reducao_pct, 2) if reducao_pct is not None else None,
            "taxa_reducao_anual_composta": round(taxa_anual, 4) if taxa_anual is not None else None,
        })

    df_resumo = pd.DataFrame(resumo)

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    destino = PASTA_SAIDA / "lcoh_projecao_brasil.xlsx"
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        df_lcoh.to_excel(writer, sheet_name="LCOH detalhado", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo por tecnologia", index=False)

    log.info(f"  {len(df_lcoh)} linha(s) de LCOH calculadas e salvas em {destino.name}")
    for _, linha in df_resumo.iterrows():
        log.info(
            f"    {linha['tecnologia']}/{linha['cenario']}: "
            f"LCOH {ANOS_LCOH[0]}={linha[f'lcoh_{ANOS_LCOH[0]}_usd_kg']:.2f} USD/kg -> "
            f"{ANOS_LCOH[-1]}={linha[f'lcoh_{ANOS_LCOH[-1]}_usd_kg']:.2f} USD/kg "
            f"(-{linha['reducao_percentual']:.1f}%)"
        )


# --------------------------------------------------------------------------
# VALIDAÇÃO COMPARATIVA ENTRE FONTES (Seção 3.7 / 4.3 do TCC)
# --------------------------------------------------------------------------

def gerar_relatorio_validacao_comparativa(df_dedup: pd.DataFrame):
    """A partir da base já deduplicada (com a coluna 'fontes_conciliadas'
    gerada por resolver_entidades), monta a comparação exigida na Seção 3.7:
    quantos projetos cada fonte contribui com exclusividade, e quantos
    aparecem em mais de uma fonte (confirmados de forma independente) —
    a evidência quantitativa da Seção 4.3 (Validação)."""
    log.info("== Gerando relatório de validação comparativa entre fontes (Seção 3.7/4.3) ==")

    if df_dedup is None or df_dedup.empty or "fontes_conciliadas" not in df_dedup.columns:
        log.warning("  Base deduplicada indisponível — validação comparativa não executada.")
        return

    contagem_exclusivos: dict[str, int] = {}
    contagem_total: dict[str, int] = {}
    total_compartilhados = 0

    for fontes_str in df_dedup["fontes_conciliadas"].dropna():
        fontes = [f.strip() for f in str(fontes_str).split(";") if f.strip()]
        if not fontes:
            continue

        for f in fontes:
            contagem_total[f] = contagem_total.get(f, 0) + 1

        if len(fontes) == 1:
            contagem_exclusivos[fontes[0]] = contagem_exclusivos.get(fontes[0], 0) + 1
        else:
            total_compartilhados += 1

    linhas = []
    for fonte in sorted(contagem_total):
        total = contagem_total[fonte]
        exclusivos = contagem_exclusivos.get(fonte, 0)
        linhas.append({
            "fonte": fonte,
            "total_de_projetos_nesta_fonte": total,
            "projetos_exclusivos_desta_fonte": exclusivos,
            "projetos_confirmados_por_outra_fonte": total - exclusivos,
            "percentual_exclusivo": round(100 * exclusivos / total, 1) if total else None,
        })

    df_validacao = pd.DataFrame(linhas)

    resumo_geral = pd.DataFrame([{
        "total_de_entidades_unicas": len(df_dedup),
        "projetos_confirmados_por_2_ou_mais_fontes": total_compartilhados,
        "projetos_vistos_em_apenas_1_fonte": len(df_dedup) - total_compartilhados,
    }])

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    destino = PASTA_SAIDA / "validacao_comparativa_fontes.xlsx"
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        df_validacao.to_excel(writer, sheet_name="Por fonte", index=False)
        resumo_geral.to_excel(writer, sheet_name="Resumo geral", index=False)

    log.info(f"  Relatório de validação comparativa salvo em {destino.name}")
    for _, linha in df_validacao.iterrows():
        log.info(
            f"    {linha['fonte']}: {linha['total_de_projetos_nesta_fonte']} projeto(s), "
            f"{linha['projetos_exclusivos_desta_fonte']} exclusivo(s) "
            f"({linha['percentual_exclusivo']}%)"
        )


# --------------------------------------------------------------------------
# VALIDAÇÃO DE QUALIDADE / SCHEMA DOS DADOS
# --------------------------------------------------------------------------

ANO_MINIMO_PLAUSIVEL = 2015
ANO_MAXIMO_PLAUSIVEL = 2050

# Bounding box aproximado do território brasileiro
BRASIL_LAT_MIN, BRASIL_LAT_MAX = -34.0, 6.0
BRASIL_LON_MIN, BRASIL_LON_MAX = -74.5, -28.5

CAMPOS_OBRIGATORIOS_QUALIDADE = ["nome_projeto"]


def validar_qualidade_dados(df: pd.DataFrame):
    """Percorre a base final e sinaliza inconsistências de tipo/faixa/campo
    obrigatório, sem alterar os dados — gera relatorio_qualidade_dados.xlsx.
    Funciona tanto com o esquema estreito (CAMPOS_PADRAO_ORDEM, campo
    'capacidade') quanto com o esquema amplo H2V (CAMPOS_H2V, campo
    'capacidade_mw') — detecta qual dos dois existe no DataFrame recebido."""
    log.info("== Validando qualidade/schema da base final ==")

    if df is None or df.empty:
        log.warning("  Base vazia — validação de qualidade não executada.")
        return

    campo_capacidade = (
        "capacidade_mw" if "capacidade_mw" in df.columns
        else "capacidade" if "capacidade" in df.columns
        else None
    )

    problemas = []

    def registrar(idx, campo, valor, motivo):
        problemas.append({
            "indice": idx,
            "nome_projeto": df.at[idx, "nome_projeto"] if "nome_projeto" in df.columns else None,
            "campo": campo,
            "valor": valor,
            "problema": motivo,
        })

    for idx, linha in df.iterrows():
        for campo in CAMPOS_OBRIGATORIOS_QUALIDADE:
            if campo in df.columns and (pd.isna(linha[campo]) or str(linha[campo]).strip() == ""):
                registrar(idx, campo, linha.get(campo), "campo obrigatório ausente")

        if campo_capacidade and pd.notna(linha[campo_capacidade]):
            valor = pd.to_numeric(linha[campo_capacidade], errors="coerce")
            if pd.isna(valor):
                registrar(idx, campo_capacidade, linha[campo_capacidade], "não é numérico")
            elif valor <= 0:
                registrar(idx, campo_capacidade, valor, "capacidade <= 0 (deveria ser positiva)")

        if "investimento" in df.columns and pd.notna(linha["investimento"]):
            valor = pd.to_numeric(linha["investimento"], errors="coerce")
            if pd.isna(valor):
                registrar(idx, "investimento", linha["investimento"], "não é numérico")
            elif valor < 0:
                registrar(idx, "investimento", valor, "investimento negativo")

        if "previsao_operacao" in df.columns and pd.notna(linha["previsao_operacao"]):
            valor = pd.to_numeric(linha["previsao_operacao"], errors="coerce")
            if pd.isna(valor):
                registrar(idx, "previsao_operacao", linha["previsao_operacao"], "não é um ano numérico")
            elif not (ANO_MINIMO_PLAUSIVEL <= valor <= ANO_MAXIMO_PLAUSIVEL):
                registrar(
                    idx, "previsao_operacao", valor,
                    f"fora da faixa plausível ({ANO_MINIMO_PLAUSIVEL}-{ANO_MAXIMO_PLAUSIVEL})"
                )

        lat = pd.to_numeric(linha.get("latitude"), errors="coerce") if "latitude" in df.columns else None
        lon = pd.to_numeric(linha.get("longitude"), errors="coerce") if "longitude" in df.columns else None
        if pd.notna(lat) and pd.notna(lon):
            if not (BRASIL_LAT_MIN <= lat <= BRASIL_LAT_MAX and BRASIL_LON_MIN <= lon <= BRASIL_LON_MAX):
                registrar(
                    idx, "latitude/longitude", f"({lat}, {lon})",
                    "coordenada fora do território brasileiro (verificar geocodificação)"
                )

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    destino = PASTA_SAIDA / "relatorio_qualidade_dados.xlsx"

    if not problemas:
        pd.DataFrame(columns=["indice", "nome_projeto", "campo", "valor", "problema"]).to_excel(
            destino, index=False
        )
        log.info("  Nenhum problema de qualidade encontrado. ✔")
        return

    df_problemas = pd.DataFrame(problemas)
    df_problemas.to_excel(destino, index=False)
    log.info(f"  {len(df_problemas)} problema(s) de qualidade encontrados — ver {destino.name}")


# --------------------------------------------------------------------------
# CHECKPOINT INCREMENTAL — evita perder o progresso se a execução cair
# --------------------------------------------------------------------------

def salvar_checkpoint():
    """Salva o estado atual de ITENS_ENCONTRADOS em disco, ao final de cada
    estágio principal do main(), para uma queda no meio da execução não
    jogar fora o que já foi achado (os arquivos já baixados continuam
    intactos em dados_h2/brutos/ de qualquer forma)."""
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    dados = [item.__dict__ for item in ITENS_ENCONTRADOS]
    try:
        with open(ARQUIVO_CHECKPOINT, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except OSError as e:
        log.warning(f"  Não foi possível salvar o checkpoint: {e}")


def carregar_checkpoint() -> list[ItemEncontrado]:
    """Carrega itens de uma execução anterior interrompida, se existir."""
    if not ARQUIVO_CHECKPOINT.exists():
        return []
    try:
        with open(ARQUIVO_CHECKPOINT, "r", encoding="utf-8") as f:
            dados = json.load(f)
        return [ItemEncontrado(**d) for d in dados]
    except (OSError, ValueError, TypeError) as e:
        log.warning(f"  Checkpoint encontrado mas ilegível ({e}) — ignorando.")
        return []


def deduplicar_itens_encontrados():
    """Remove entradas repetidas em ITENS_ENCONTRADOS (mesma fonte + mesma
    URL) — pode acontecer ao mesclar um checkpoint de execução anterior com
    os itens desta execução."""
    vistos = set()
    unicos = []
    for item in ITENS_ENCONTRADOS:
        chave = (normalizar_texto(item.fonte), (item.url or item.caminho_local or "").rstrip("/").lower())
        if chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(item)

    removidos = len(ITENS_ENCONTRADOS) - len(unicos)
    ITENS_ENCONTRADOS.clear()
    ITENS_ENCONTRADOS.extend(unicos)
    if removidos:
        log.info(f"  {removidos} item(ns) duplicado(s) removido(s) (mesma fonte + URL).")


# --------------------------------------------------------------------------
# SELENIUM OPCIONAL — para as fontes que são SPAs em JavaScript e por isso
# não aparecem no scraping HTML simples (CCEE, ONS, H2LAC, H2 Brazil).
# Só roda se USAR_SELENIUM_PARA_SPA=1 (variável de ambiente) E o pacote
# selenium + um navegador Chrome/Chromium estiverem instalados.
# --------------------------------------------------------------------------

def buscar_em_pagina_spa_com_selenium(nome_fonte: str, url_pagina: str):
    log.info(f"== [Selenium] Varrendo página em JavaScript: {nome_fonte} ({url_pagina}) ==")

    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
    except ImportError:
        log.warning(
            "  Selenium não está instalado — pulando esta fonte. Rode "
            "'pip install selenium' e garanta que o Chrome/Chromium esteja "
            "instalado para habilitar essa coleta."
        )
        return

    opcoes = Options()
    opcoes.add_argument("--headless=new")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument(f"user-agent={HEADERS['User-Agent']}")

    driver = None
    encontrados = 0
    try:
        driver = webdriver.Chrome(options=opcoes)
        driver.set_page_load_timeout(TIMEOUT)
        driver.get(url_pagina)
        time.sleep(TEMPO_ESPERA_SELENIUM)  # dá tempo do JS renderizar os links

        dominio_base = urlparse(url_pagina).netloc.lower()
        elementos = driver.find_elements("tag name", "a")

        for el in elementos:
            try:
                href = el.get_attribute("href")
                texto_link = el.text.strip()
            except Exception:
                continue

            if not href:
                continue

            url_completa = href.split("#")[0]
            if urlparse(url_completa).netloc.lower() != dominio_base:
                continue

            if _adicionar_item_arquivo(nome_fonte, texto_link or url_completa, url_completa, url_pagina):
                encontrados += 1

    except Exception as e:
        log.warning(f"  Falha ao executar Selenium em {nome_fonte}: {e}")
    finally:
        if driver is not None:
            driver.quit()

    log.info(f"  [Selenium] {encontrados} arquivo(s) relevante(s) descoberto(s) em {nome_fonte}")


# --------------------------------------------------------------------------
# ARQUIVAMENTO / LIMPEZA — necessário para execução automática semanal, para
# não acumular arquivos velhos misturados com os novos a cada rodada.
# --------------------------------------------------------------------------

def arquivar_e_limpar_execucao_anterior():
    """Se houver resultado de uma execução anterior (índice existente), move
    os artefatos principais para dados_h2/historico/<data-e-hora>/ e limpa a
    pasta de trabalho para a nova coleta começar do zero — evita acumular
    arquivos numerados de execuções antigas junto com os novos.

    NÃO mexe em:
    - dados_h2/entrada_manual/ (arquivos que você adicionou manualmente)
    - dados_h2/historico/ (as próprias rodadas já arquivadas)
    """
    if not ARQUIVAR_EXECUCAO_ANTERIOR:
        return

    if not PASTA_SAIDA.exists() or not ARQUIVO_INDICE.exists():
        # Não há execução anterior completa para arquivar — segue normalmente.
        PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
        return

    import shutil
    from datetime import datetime

    carimbo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    destino_historico = PASTA_HISTORICO / carimbo
    destino_historico.mkdir(parents=True, exist_ok=True)

    log.info(f"== Arquivando resultado da execução anterior em historico/{carimbo}/ ==")

    # Move os principais artefatos de nível raiz (não a pasta inteira, para
    # preservar entrada_manual/ e historico/ no lugar).
    padroes_para_arquivar = ["*.xlsx", "*.csv"]
    for padrao in padroes_para_arquivar:
        for arquivo in PASTA_SAIDA.glob(padrao):
            try:
                shutil.move(str(arquivo), str(destino_historico / arquivo.name))
            except OSError as e:
                log.warning(f"  Não foi possível arquivar {arquivo.name}: {e}")

    for subpasta in ("brutos", "convertidos", "modelo_estrela"):
        origem = PASTA_SAIDA / subpasta
        if origem.exists():
            try:
                shutil.move(str(origem), str(destino_historico / subpasta))
            except OSError as e:
                log.warning(f"  Não foi possível arquivar a pasta {subpasta}/: {e}")

    log.info(f"  Execução anterior arquivada. Coleta desta semana começa limpa.")


# --------------------------------------------------------------------------
# GERAÇÃO DO ÍNDICE EM EXCEL
# --------------------------------------------------------------------------

def gerar_indice_excel():
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

    if not ITENS_ENCONTRADOS:
        log.warning("Nenhum item encontrado — gerando índice vazio.")
        df = pd.DataFrame(
            columns=[
                "Fonte", "Título", "Formato", "URL",
                "Dataset/Página de origem", "Arquivo local", "Status do download",
            ]
        )
    else:
        df = pd.DataFrame(
            [
                {
                    "Fonte": it.fonte,
                    "Título": it.titulo,
                    "Formato": it.formato,
                    "URL": it.url,
                    "Dataset/Página de origem": it.dataset_ou_pagina,
                    "Arquivo local": it.caminho_local,
                    "Status do download": it.status_download,
                }
                for it in ITENS_ENCONTRADOS
            ]
        )

    with pd.ExcelWriter(ARQUIVO_INDICE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Índice geral", index=False)

        for fonte in df["Fonte"].unique() if not df.empty else []:
            aba = re.sub(r"[\\/*?:\[\]]", "_", fonte)[:31]
            df[df["Fonte"] == fonte].to_excel(writer, sheet_name=aba, index=False)

        for sheet in writer.sheets.values():
            for col_cells in sheet.columns:
                comprimento = max((len(str(c.value)) for c in col_cells if c.value), default=10)
                sheet.column_dimensions[col_cells[0].column_letter].width = min(comprimento + 2, 80)

    log.info(f"Índice consolidado salvo em: {ARQUIVO_INDICE.resolve()}")


# --------------------------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# --------------------------------------------------------------------------

def main():
    log.info("Iniciando coleta de dados sobre hidrogênio de baixa emissão de carbono no Brasil")

    arquivar_e_limpar_execucao_anterior()

    # Recupera itens de uma execução anterior interrompida, se existir
    itens_checkpoint = carregar_checkpoint()
    if itens_checkpoint:
        log.info(
            f"  Checkpoint de execução anterior encontrado: {len(itens_checkpoint)} "
            f"item(ns) recuperados (arquivos já baixados não serão baixados de novo)."
        )
        ITENS_ENCONTRADOS.extend(itens_checkpoint)

    for nome, url in PORTAIS_CKAN.items():
        buscar_em_portal_ckan(nome, url)
    salvar_checkpoint()

    camadas_h2 = descobrir_camadas_arcgis_h2()
    for nome, url in camadas_h2.items():
        buscar_em_camada_arcgis(nome, url)
    salvar_checkpoint()

    buscar_em_eia()
    salvar_checkpoint()

    for nome, url in PAGINAS_HTML.items():
        if nome in FONTES_SPA_JAVASCRIPT:
            if USAR_SELENIUM_PARA_SPA:
                buscar_em_pagina_spa_com_selenium(nome, url)
            else:
                log.info(
                    f"== Pulando varredura profunda de '{nome}' (é uma SPA em "
                    f"JavaScript) — defina USAR_SELENIUM_PARA_SPA=1 para tentar "
                    f"via Selenium. Tentando o crawler HTML simples mesmo assim... =="
                )
                buscar_em_pagina_html(nome, url)
        else:
            buscar_em_pagina_html(nome, url)
    salvar_checkpoint()

    importar_entrada_manual()
    deduplicar_itens_encontrados()
    salvar_checkpoint()

    log.info(f"Total de itens (dados + relatórios) encontrados: {len(ITENS_ENCONTRADOS)}")

    baixar_itens_encontrados()
    salvar_checkpoint()

    # Base estreita (mapeamento manual, só a camada "Projetos" da EPE) —
    # mantida por compatibilidade / conferência cruzada.
    df_padrao = gerar_base_projetos_padronizada()

    # Base ampla (heurística de alias de colunas, varre TODOS os arquivos
    # baixados de TODAS as fontes) — é essa que alimenta a resolução de
    # entidades e validações a partir de agora.
    df_h2v = gerar_base_h2v_multifonte()
    gerar_modelo_estrela_h2v(df_h2v)

    df_dedup = resolver_entidades(df_h2v)

    if df_dedup is not None and not df_dedup.empty:
        df_dedup.to_excel(PASTA_SAIDA / "base_projetos_deduplicada.xlsx", index=False)

        gerar_relatorio_validacao_comparativa(df_dedup)
        validar_qualidade_dados(df_dedup)

    calcular_lcoh_referencia()
    gerar_indice_excel()

    # Execução concluída com sucesso -- o checkpoint só existe para recuperar
    # de uma queda no meio do caminho; numa execução completa, removemos para
    # a próxima rodada começar limpa (evita acumular itens obsoletos).
    if ARQUIVO_CHECKPOINT.exists():
        try:
            ARQUIVO_CHECKPOINT.unlink()
        except OSError:
            pass

    log.info("Concluído. Verifique a pasta 'dados_h2/' para os arquivos e o índice em Excel.")


if __name__ == "__main__":
    main()