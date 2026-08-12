#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import math
from pathlib import Path
from datetime import datetime, timezone
from typing import Optional

import pandas as pd
from fastapi import FastAPI, HTTPException, Query, Header, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware


DADOS_API_DIR = Path(os.environ.get("DADOS_API_DIR", "dados_api"))
ARQUIVO_PROJETOS = DADOS_API_DIR / "base_projetos_deduplicada.xlsx"
ARQUIVO_LCOH = DADOS_API_DIR / "lcoh_projecao_brasil.xlsx"

# Pasta genérica de tabelas extras (modelo estrela, validação, qualidade):
# qualquer .csv colocado aqui vira automaticamente um endpoint /tabelas/{nome}
# — não precisa mexer no código da API para adicionar uma tabela nova, só
# atualizar o preparar_dados_api.py para copiar o arquivo pra cá.
PASTA_TABELAS_EXTRAS = DADOS_API_DIR / "tabelas"

CHAVE_ADMIN = os.environ.get("API_ADMIN_KEY", "")


app = FastAPI(
    title="API - Hidrogênio de Baixa Emissão de Carbono no Brasil",
    description=(
        "Dados de projetos de hidrogênio de baixa emissão de carbono no Brasil, "
        "coletados e consolidados a partir de fontes públicas (EPE, IEA, ANEEL, "
        "EIA, MME, BNDES, ANP e outras). Projeto de TCC — Engenharia da "
        "Computação, UNIFEI."
    ),
    version="1.0.0",
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


_estado = {
    "df_projetos": pd.DataFrame(),
    "df_lcoh": pd.DataFrame(),
    "carregado_em": None,
}


def _valor_para_json(valor):
    if valor is None:
        return None

    if isinstance(valor, dict):
        return {k: _valor_para_json(v) for k, v in valor.items()}

    if isinstance(valor, list):
        return [_valor_para_json(v) for v in valor]

    if isinstance(valor, tuple):
        return [_valor_para_json(v) for v in valor]

    if pd.isna(valor):
        return None

    if hasattr(valor, "item"):
        try:
            valor = valor.item()
        except (ValueError, TypeError):
            pass

    if isinstance(valor, (datetime, pd.Timestamp)):
        return valor.isoformat()

    return valor


def _df_para_json(df: pd.DataFrame):
    registros = df.to_dict(orient="records")
    return _valor_para_json(registros)


def carregar_dados():
    df_projetos = pd.DataFrame()
    df_lcoh = pd.DataFrame()

    if ARQUIVO_PROJETOS.exists():
        df_projetos = pd.read_excel(ARQUIVO_PROJETOS)

    if ARQUIVO_LCOH.exists():
        df_lcoh = pd.read_excel(
            ARQUIVO_LCOH,
            sheet_name="LCOH detalhado"
        )

    _estado["df_projetos"] = df_projetos
    _estado["df_lcoh"] = df_lcoh
    _estado["carregado_em"] = datetime.now(timezone.utc).isoformat()


@app.on_event("startup")
def ao_iniciar():
    DADOS_API_DIR.mkdir(parents=True, exist_ok=True)
    PASTA_TABELAS_EXTRAS.mkdir(parents=True, exist_ok=True)
    carregar_dados()


@app.get("/", tags=["Info"])
def raiz():
    return {
        "nome": "API - Hidrogênio de Baixa Emissão de Carbono no Brasil",
        "documentacao": "/docs",
        "endpoints": [
            "/health",
            "/projetos",
            "/projetos/{id_projeto}",
            "/lcoh",
            "/estatisticas",
            "/tabelas",
            "/tabelas/{nome}",
        ],
    }


@app.get("/tabelas", tags=["Tabelas extras"])
def listar_tabelas():
    """Lista as tabelas extras disponíveis (modelo estrela, validação,
    qualidade dos dados) — cada uma é um arquivo .csv em dados_api/tabelas/,
    copiado para lá pelo preparar_dados_api.py."""
    if not PASTA_TABELAS_EXTRAS.exists():
        return {"tabelas": []}

    nomes = sorted(p.stem for p in PASTA_TABELAS_EXTRAS.glob("*.csv"))
    return {"total": len(nomes), "tabelas": nomes}


@app.get("/tabelas/{nome}", tags=["Tabelas extras"])
def obter_tabela(nome: str):
    """Retorna o conteúdo de uma tabela extra (ver /tabelas para a lista de
    nomes disponíveis) — ex.: /tabelas/fato_projetos_h2v, /tabelas/dim_estados."""
    caminho = PASTA_TABELAS_EXTRAS / f"{nome}.csv"
    if not caminho.exists():
        raise HTTPException(
            status_code=404,
            detail=f"Tabela '{nome}' não encontrada. Veja /tabelas para a lista de nomes disponíveis.",
        )

    try:
        df = pd.read_csv(caminho)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Falha ao ler a tabela '{nome}': {e}")

    return {"total": len(df), "resultados": _df_para_json(df)}


@app.get("/health", tags=["Info"])
def health():
    tabelas_extras = sorted(p.stem for p in PASTA_TABELAS_EXTRAS.glob("*.csv")) if PASTA_TABELAS_EXTRAS.exists() else []
    return {
        "status": "ok",
        "projetos_carregados": len(_estado["df_projetos"]),
        "linhas_lcoh_carregadas": len(_estado["df_lcoh"]),
        "tabelas_extras_disponiveis": tabelas_extras,
        "dados_atualizados_em": _estado["carregado_em"],
    }


@app.get("/projetos", tags=["Projetos"])
def listar_projetos(
    estado: Optional[str] = Query(
        None,
        description="Filtra por UF/estado (ex.: 'CE')"
    ),
    tecnologia: Optional[str] = Query(
        None,
        description="Filtra por tecnologia de eletrólise (ex.: 'PEM')"
    ),
    status: Optional[str] = Query(
        None,
        description="Filtra por status/maturidade (ex.: 'Planejado')"
    ),
    limite: int = Query(
        100,
        ge=1,
        le=1000,
        description="Quantos registros retornar (máx. 1000)"
    ),
    pagina: int = Query(
        1,
        ge=1,
        description="Número da página (começa em 1)"
    ),
):
    df = _estado["df_projetos"]

    if df.empty:
        return {
            "total": 0,
            "pagina": pagina,
            "limite": limite,
            "resultados": []
        }

    if estado and "estado_regiao" in df.columns:
        df = df[
            df["estado_regiao"]
            .astype(str)
            .str.contains(estado, case=False, na=False)
        ]

    if tecnologia and "tecnologia_eletrolise" in df.columns:
        df = df[
            df["tecnologia_eletrolise"]
            .astype(str)
            .str.contains(tecnologia, case=False, na=False)
        ]

    if status and "status_maturidade" in df.columns:
        df = df[
            df["status_maturidade"]
            .astype(str)
            .str.contains(status, case=False, na=False)
        ]

    total = len(df)

    inicio = (pagina - 1) * limite
    fim = inicio + limite

    pagina_df = df.iloc[inicio:fim]

    return {
        "total": total,
        "pagina": pagina,
        "limite": limite,
        "total_paginas": math.ceil(total / limite) if total else 0,
        "resultados": _df_para_json(pagina_df),
    }


@app.get("/projetos/{id_projeto}", tags=["Projetos"])
def obter_projeto(id_projeto: str):
    df = _estado["df_projetos"]

    if df.empty or "id_projeto" not in df.columns:
        raise HTTPException(
            status_code=404,
            detail="Base de projetos vazia ou sem coluna id_projeto"
        )

    encontrado = df[df["id_projeto"].astype(str) == str(id_projeto)]

    if encontrado.empty:
        raise HTTPException(
            status_code=404,
            detail=f"Projeto '{id_projeto}' não encontrado"
        )

    return _valor_para_json(
        encontrado.iloc[0].to_dict()
    )


@app.get("/lcoh", tags=["LCOH"])
def listar_lcoh(
    tecnologia: Optional[str] = Query(
        None,
        description="Filtra por tecnologia (ALK/PEM)"
    ),
    cenario: Optional[str] = Query(
        None,
        description="Filtra por cenário (conservador/otimista)"
    ),
):
    df = _estado["df_lcoh"]

    if df.empty:
        return {
            "total": 0,
            "resultados": []
        }

    if tecnologia and "tecnologia" in df.columns:
        df = df[
            df["tecnologia"]
            .astype(str)
            .str.lower() == tecnologia.lower()
        ]

    if cenario and "cenario" in df.columns:
        df = df[
            df["cenario"]
            .astype(str)
            .str.lower() == cenario.lower()
        ]

    return {
        "total": len(df),
        "resultados": _df_para_json(df)
    }


@app.get("/estatisticas", tags=["Info"])
def estatisticas():
    df = _estado["df_projetos"]

    if df.empty:
        return {
            "total_projetos": 0
        }

    def contagem_serializavel(serie: pd.Series) -> dict:
        return {
            str(k): int(v)
            for k, v in serie
            .dropna()
            .astype(str)
            .value_counts()
            .items()
        }

    resposta = {
        "total_projetos": len(df)
    }

    if "estado_regiao" in df.columns:
        resposta["por_estado"] = contagem_serializavel(
            df["estado_regiao"]
        )

    if "tecnologia_eletrolise" in df.columns:
        resposta["por_tecnologia"] = contagem_serializavel(
            df["tecnologia_eletrolise"]
        )

    if "status_maturidade" in df.columns:
        resposta["por_status"] = contagem_serializavel(
            df["status_maturidade"]
        )

    if "capacidade_mw" in df.columns:
        soma = pd.to_numeric(
            df["capacidade_mw"],
            errors="coerce"
        ).sum(skipna=True)

        resposta["capacidade_total_mw"] = float(
            round(soma, 2)
        )

    return resposta


@app.post("/admin/atualizar", tags=["Admin"])
async def atualizar_dados(
    x_api_key: str = Header(
        ...,
        description="Chave de administração (API_ADMIN_KEY)"
    ),
    arquivo_projetos: Optional[UploadFile] = File(None),
    arquivo_lcoh: Optional[UploadFile] = File(None),
):
    if not CHAVE_ADMIN:
        raise HTTPException(
            status_code=503,
            detail=(
                "Endpoint de atualização desabilitado: "
                "variável API_ADMIN_KEY não configurada no servidor."
            ),
        )

    if x_api_key != CHAVE_ADMIN:
        raise HTTPException(
            status_code=401,
            detail="Chave de administração inválida."
        )

    DADOS_API_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    atualizados = []

    if arquivo_projetos is not None:
        conteudo = await arquivo_projetos.read()
        ARQUIVO_PROJETOS.write_bytes(conteudo)
        atualizados.append("projetos")

    if arquivo_lcoh is not None:
        conteudo = await arquivo_lcoh.read()
        ARQUIVO_LCOH.write_bytes(conteudo)
        atualizados.append("lcoh")

    if not atualizados:
        raise HTTPException(
            status_code=400,
            detail="Nenhum arquivo enviado."
        )

    carregar_dados()

    return {
        "status": "ok",
        "arquivos_atualizados": atualizados,
        "projetos_carregados": len(
            _estado["df_projetos"]
        ),
        "linhas_lcoh_carregadas": len(
            _estado["df_lcoh"]
        ),
    }