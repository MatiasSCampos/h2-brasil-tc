#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Prepara os dados do coletor_h2_brasil.py para a API pública.

Copia só o essencial de dados_h2/ para dados_api/ (pasta enxuta, sem os
arquivos brutos de cada fonte) — é essa pasta menor que vai pro repositório
Git e é servida pela API no Render.

Rode isso DEPOIS do coletor_h2_brasil.py, sempre:
    python coletor_h2_brasil.py
    python preparar_dados_api.py
"""

import shutil
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("preparar_dados_api")

PASTA_COLETOR = Path("dados_h2")
PASTA_API = Path("dados_api")

ARQUIVOS_NECESSARIOS = [
    "base_projetos_deduplicada.xlsx",
    "lcoh_projecao_brasil.xlsx",
]


def main():
    PASTA_API.mkdir(parents=True, exist_ok=True)

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

    log.info(f"Pronto. Pasta '{PASTA_API}/' atualizada — é essa que deve ir pro git push.")


if __name__ == "__main__":
    main()
