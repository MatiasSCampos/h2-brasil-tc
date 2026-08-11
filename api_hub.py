import uvicorn
from fastapi import FastAPI
from pydantic import BaseModel
from typing import Optional, List

app = FastAPI(
    title="Hub de Dados - Hidrogenio Verde (TFG UNIFEI)",
    description="API central para projecoes de LCOH, projetos e consumo.",
    version="1.2.0",
)

# Armazenamento em memoria, uma lista por tabela
ARMAZENAMENTO_LCOH = []
ARMAZENAMENTO_PROJETOS = []
ARMAZENAMENTO_CONSUMO = []


# --- Modelos de validacao (Pydantic) ---
class RegistroLCOH(BaseModel):
    Ano: int
    Tecnologia: str
    Cenario: str
    CAPEX_USD_kW: float
    Eficiencia_kWh_kg: float
    Preco_Energia_USD_MWh: float
    Componente_CAPEX: float
    Componente_Energia: float
    LCOH_Final_USD_kg: float


class RegistroProjeto(BaseModel):
    Id: int
    Nome: str
    Capacidade_T_ano: Optional[float] = None
    Finalidade: Optional[str] = None
    Estagio: Optional[str] = None
    Local: Optional[str] = None
    Valor_R_Milhoes: Optional[float] = None


class RegistroConsumo(BaseModel):
    Data: int
    Regiao: str
    Classe: str
    Consumo: float
    Consumidores: Optional[float] = None


# --- Rotas POST: o pipeline alimenta a base ---
@app.post("/inserir-lcoh")
def inserir_lcoh(dados: List[RegistroLCOH]):
    global ARMAZENAMENTO_LCOH
    ARMAZENAMENTO_LCOH = [d.model_dump() for d in dados]
    return {"status": "ok", "registros": len(ARMAZENAMENTO_LCOH)}


@app.post("/inserir-projetos")
def inserir_projetos(dados: List[RegistroProjeto]):
    global ARMAZENAMENTO_PROJETOS
    ARMAZENAMENTO_PROJETOS = [d.model_dump() for d in dados]
    return {"status": "ok", "registros": len(ARMAZENAMENTO_PROJETOS)}


@app.post("/inserir-consumo")
def inserir_consumo(dados: List[RegistroConsumo]):
    global ARMAZENAMENTO_CONSUMO
    ARMAZENAMENTO_CONSUMO = [d.model_dump() for d in dados]
    return {"status": "ok", "registros": len(ARMAZENAMENTO_CONSUMO)}


# --- Rotas GET: o Power BI consome os dados tratados ---
@app.get("/lcoh")
def obter_lcoh():
    return ARMAZENAMENTO_LCOH


@app.get("/projetos-epe")
def obter_projetos():
    return ARMAZENAMENTO_PROJETOS


@app.get("/consumo-eletrobras")
def obter_consumo():
    return ARMAZENAMENTO_CONSUMO


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
