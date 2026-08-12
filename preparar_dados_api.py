#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Prepara os dados do coletor_h2_brasil.py para a API pública.

Copia só o essencial de dados_h2/ para dados_api/ (pasta enxuta, sem os
arquivos brutos de cada fonte) — é essa pasta menor que vai pro repositório
Git e é servida pela API no Render.

Além das duas bases principais (projetos e LCOH), também prepara em
dados_api/tabelas/ o modelo estrela (fato + dimensões) e os relatórios de
validação comparativa e qualidade de dados — cada um vira automaticamente um
endpoint /tabelas/{nome} na API, sem precisar mexer no código dela.

Rode isso DEPOIS do coletor_h2_brasil.py, sempre:
    python coletor_h2_brasil.py
    python preparar_dados_api.py
"""

import shutil
import logging
from pathlib import Path

import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("preparar_dados_api")

PASTA_COLETOR = Path("dados_h2")
PASTA_API = Path("dados_api")
PASTA_TABELAS = PASTA_API / "tabelas"

# Arquivos principais, copiados tal como estão (a API lê o .xlsx direto)
ARQUIVOS_NECESSARIOS = [
    "base_projetos_deduplicada.xlsx",
    "lcoh_projecao_brasil.xlsx",
]

# Tabelas do modelo estrela — já vêm em .csv, só copiar (renomeando o nome
# do arquivo final para o nome que vai aparecer em /tabelas/{nome})
TABELAS_MODELO_ESTRELA = {
    "fato_projetos_h2v.csv": "fato_projetos_h2v",
    "dim_fontes.csv": "dim_fontes",
    "dim_estados.csv": "dim_estados",
    "dim_empresas.csv": "dim_empresas",
    "dim_status.csv": "dim_status",
    "dim_tecnologias.csv": "dim_tecnologias",
    "dim_aplicacoes.csv": "dim_aplicacoes",
    "dim_localizacao.csv": "dim_localizacao",
}

# Tabelas que já ficam prontas em .csv direto na raiz de dados_h2/ (não
# dentro de uma subpasta) — só copiar pra dados_api/tabelas/.
TABELAS_RAIZ = {
    "potencial_tecnico_h2_por_usina.csv": "potencial_tecnico_h2_por_usina",
}


# Relatórios em .xlsx que precisam ser convertidos para .csv (um arquivo de
# saída por aba relevante)
RELATORIOS_PARA_CONVERTER = {
    "validacao_comparativa_fontes.xlsx": {
        "Por fonte": "validacao_por_fonte",
        "Resumo geral": "validacao_resumo_geral",
    },
    "relatorio_qualidade_dados.xlsx": {
        None: "qualidade_dados",  # None = primeira/única aba
    },
    "pares_similares_resolucao_entidades.xlsx": {
        None: "pares_similares_entidades",
    },
}


def copiar_arquivos_principais():
    faltando = []
    for nome_arquivo in ARQUIVOS_NECESSARIOS:
        origem = PASTA_COLETOR / nome_arquivo
        destino = PASTA_API / nome_arquivo

        if not origem.exists():
            faltando.append(nome_arquivo)
            continue

        shutil.copy2(origem, destino)
        log.info(f"Copiado: {nome_arquivo} ({origem.stat().st_size / 1024:.1f} KB)")

    if faltando:
        log.warning(
            f"Não encontrados em {PASTA_COLETOR}/: {', '.join(faltando)} — "
            f"rode coletor_h2_brasil.py primeiro."
        )


def copiar_modelo_estrela():
    pasta_origem = PASTA_COLETOR / "modelo_estrela"
    if not pasta_origem.exists():
        log.warning(f"  Pasta '{pasta_origem}/' não encontrada — modelo estrela não preparado.")
        return

    for nome_arquivo, nome_tabela in TABELAS_MODELO_ESTRELA.items():
        origem = pasta_origem / nome_arquivo
        if not origem.exists():
            continue
        destino = PASTA_TABELAS / f"{nome_tabela}.csv"
        shutil.copy2(origem, destino)
        log.info(f"Copiado (modelo estrela): {nome_tabela}.csv")


def converter_relatorios():
    for nome_arquivo, abas in RELATORIOS_PARA_CONVERTER.items():
        origem = PASTA_COLETOR / nome_arquivo
        if not origem.exists():
            continue

        for nome_aba, nome_tabela in abas.items():
            try:
                if nome_aba is None:
                    df = pd.read_excel(origem)
                else:
                    df = pd.read_excel(origem, sheet_name=nome_aba)
            except Exception as e:
                log.warning(f"  Falha ao converter '{nome_arquivo}' (aba {nome_aba}): {e}")
                continue

            destino = PASTA_TABELAS / f"{nome_tabela}.csv"
            df.to_csv(destino, index=False, encoding="utf-8-sig")
            log.info(f"Convertido: {nome_tabela}.csv ({len(df)} linha(s))")


def copiar_tabelas_raiz():
    for nome_arquivo, nome_tabela in TABELAS_RAIZ.items():
        origem = PASTA_COLETOR / nome_arquivo
        if not origem.exists():
            continue
        destino = PASTA_TABELAS / f"{nome_tabela}.csv"
        shutil.copy2(origem, destino)
        log.info(f"Copiado: {nome_tabela}.csv")


def main():
    PASTA_API.mkdir(parents=True, exist_ok=True)
    PASTA_TABELAS.mkdir(parents=True, exist_ok=True)

    copiar_arquivos_principais()
    copiar_modelo_estrela()
    copiar_tabelas_raiz()
    converter_relatorios()

    log.info(f"Pronto. Pastas '{PASTA_API}/' e '{PASTA_TABELAS}/' atualizadas — é isso que deve ir pro git push.")


if __name__ == "__main__":
    main()
