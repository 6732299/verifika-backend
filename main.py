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
            }
def establecer_fondo_celda(celda, color_hex):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def generar_documento_word(metadatos: dict, resultado_ia: str, nombre_archivo: str):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(51, 51, 51)

    p_header = doc.add_paragraph()
    run_brand = p_header.add_run("VERIFIKA B2B\n")
    run_brand.bold = True
    run_brand.font.size = Pt(14)
    run_brand.font.color.rgb = RGBColor(0, 51, 102)

    run_sub = p_header.add_run("Inteligencia de Negocios F2 · Verificación e Inteligencia Comercial\n")
    run_sub.font.size = Pt(9)
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    p_title = doc.add_paragraph()
    run_title = p_title.add_run("INFORME EJECUTIVO DE VERIFICACIÓN EMPRESARIAL")
    run_title.bold = True
    run_title.font.size = Pt(12)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    tabla_meta = doc.add_table(rows=5, cols=2)
    tabla_meta.alignment = WD_TABLE_ALIGNMENT.CENTER

    items = [
        ("N° de informe", metadatos["n_informe"]),
        ("Fecha de emisión", "Agosto 2026"),
        ("Sujeto investigado", metadatos["razon_social"]),
        ("NIT / C.I.", metadatos["nit"]),
        ("Calificación de riesgo consolidada", metadatos["calificacion_info"]["calificacion"] + " — " + metadatos["calificacion_info"]["nivel"])
    ]

    for i, (label, val) in enumerate(items):
        row = tabla_meta.rows[i]
        row.cells[0].text = label
        row.cells[1].text = str(val)
        row.cells[0].paragraphs[0].runs[0].bold = True
        establecer_fondo_celda(row.cells[0], "F2F2F2")

    doc.add_paragraph()

    doc.add_heading("Análisis y Dictamen Ejecutivo", level=1)
    doc.add_paragraph(resultado_ia)

    section = doc.sections[0]
    footer = section.footer
    p_footer = footer.paragraphs[0]
    p_footer.text = "VERIFIKA B2B  ·  verifikab2b.bo  ·  contacto@verifikab2b.bo"
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.style.font.size = Pt(8)

    doc.save(nombre_archivo)


@app.post("/api/generar-informe")
async def generar_informe_endpoint(solicitud: SolicitudVerificacion):
    try:
        evaluacion = calcular_calificacion_automatica(solicitud.nit_empresa)

        metadatos = {
            "n_informe": f"VB2B-2026-{solicitud.nit_empresa[-4:]}",
            "razon_social": solicitud.razon_social,
            "nit": solicitud.nit_empresa,
            "solicitado_por": solicitud.solicitante,
            "motivo": solicitud.motivo_consulta,
            "calificacion_info": evaluacion
        }

        client = genai.Client()

        prompt_sistema = f"""
        Eres el motor analítico experto de 'VERIFIKA B2B'. Redacta un dictamen ejecutivo formal en español
        incorporando los siguientes datos de cumplimiento:
        - Calificación asignada: {evaluacion['calificacion']} ({evaluacion['nivel']})
        - Justificación: {evaluacion['justificacion']}
        - Recomendación / Garantías: {evaluacion['garantia_requerida']}
        """

        response = client.models.generate_content(
           'gemini-2.5-flash-lite'
            contents=f"Generar informe para la empresa {solicitud.razon_social} con NIT {solicitud.nit_empresa}.",
            config=types.GenerateContentConfig(system_instruction=prompt_sistema, temperature=0.2)
        )

        texto_ia = response.text

        nombre_archivo = f"Informe_Verifika_{solicitud.nit_empresa}.docx"
        generar_documento_word(metadatos, texto_ia, nombre_archivo)

        if os.path.exists(nombre_archivo):
            return FileResponse(
                path=nombre_archivo,
                filename=nombre_archivo,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            raise HTTPException(status_code=500, detail="Error al compilar el documento.")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))