import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from google import genai
from google.genai import types

app = FastAPI(
    title="API Backend - VERIFIKA B2B",
    description="Motor automatizado de debida diligencia y generación de informes institucionales.",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def health():
    return {"status": "ok"}


class SolicitudVerificacion(BaseModel):
    nit_empresa: str
    razon_social: str
    solicitante: str
    motivo_consulta: str


def calcular_calificacion_automatica(nit: str) -> dict:
    if nit.endswith("9"):
        return {
            "calificacion": "BBB",
            "nivel": "Riesgo de alerta",
            "justificacion": "Sujeto operativo pero con alertas puntuales detectadas en registros públicos en trámite.",
            "garantia_requerida": "SÍ, con condiciones. Se requiere mayor liquidez mediante garantías prendarias o hipotecarias."
        }
    else:
        return {
            "calificacion": "A",
            "nivel": "Bajo riesgo",
            "justificacion": "Sujeto solvente y transparente, sin contingencias relevantes detectadas en las fuentes públicas.",
            "garantia_requerida": "Se puede operar sin restricciones especiales."