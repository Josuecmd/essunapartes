import io
import os
from typing import List, Optional
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(
    title="Sistema de Parte General - Brigada de Guardiamarinas",
    description="Backend para el conteo, control nominal y exportación del parte militar.",
    version="1.0.0"
)

# Detectar la ruta base del proyecto de forma dinámica para Vercel
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

@app.get("/", response_class=HTMLResponse)
async def leer_tablero(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# --- CONFIGURACIÓN DE CONSTANTES MILITARES ---
SECCIONES_VALIDAS = [
    "Primer Año Alfa", "Primer Año Bravo", "Primer Año Charlie", "Primer Año Delta", "Primer Año Echo",
    "Segundo Año Alfa", "Segundo Año Bravo",
    "Tercer Año Alfa", "Tercer Año Bravo", "Tercer Año Abastecimientos",
    "Cuarto Año Alfa", "Cuarto Año Bravo", "Cuarto Año Abastecimientos"
]

# ... Deja todo el resto de tu código de lógica de negocio y endpoints (POST) exactamente igual abajo ...

ESTADOS_VALIDOS = ["FILA", "COMISIÓN", "EXC. LITERA", "EXC. FORMACIÓN", "EXC. POR REVALIDAR", "DESCANSO DOMICILIO"]

# --- MODELOS DE DATOS (PYDANTIC) ---
class Guardiamarina(BaseModel):
    id: int
    nombre: str = Field(..., example="Guardiamarina Juan Pérez")
    seccion: str = Field(..., example="Primer Año Alfa")
    estado: str = Field(default="FILA", example="FILA")

class ParteRequest(BaseModel):
    nombre_archivo: str = Field(default="Parte_General_Brigada", example="Parte_Diario_07_Junio")
    brigada: List[Guardiamarina]

# --- FUNCIONES DE PROCESAMIENTO Y FORMATEO ---
def generar_excel_profesional(df_nominal: pd.DataFrame, df_resumen: pd.DataFrame) -> io.BytesIO:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumen.to_excel(writer, sheet_name='Resumen Numérico', index=True)
        df_nominal.to_excel(writer, sheet_name='Listado Nominal', index=False)
        workbook = writer.book
        
        ws_resumen = workbook['Resumen Numérico']
        ws_resumen.views.sheetView[0].showGridLines = True
        
        font_header = Font(name='Arial', size=11, bold=True, color='FFFFFF')
        font_body = Font(name='Arial', size=10)
        font_total = Font(name='Arial', size=10, bold=True)
        fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid') 
        fill_total_row = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid') 
        border_thin = Side(border_style="thin", color="D9D9D9")
        border_double = Side(border_style="double", color="000000")
        box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
        total_border = Border(top=border_thin, bottom=border_double)

        ws_resumen['A1'] = "Secciones / Cursos"
        for col in range(1, ws_resumen.max_column + 1):
            cell = ws_resumen.cell(row=1, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for row in range(2, ws_resumen.max_row + 1):
            is_total_row = (ws_resumen.cell(row=row, column=1).value == "TOTAL GENERAL")
            for col in range(1, ws_resumen.max_column + 1):
                cell = ws_resumen.cell(row=row, column=col)
                cell.font = font_total if (is_total_row or col == ws_resumen.max_column) else font_body
                cell.border = total_border if is_total_row else box_border
                if is_total_row:
                    cell.fill = fill_total_row
                cell.alignment = Alignment(horizontal='center' if col > 1 else 'left')

        for col in ws_resumen.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_resumen.column_dimensions[col_letter].width = max(max_len + 3, 12)

        ws_nominal = workbook['Listado Nominal']
        ws_nominal.views.sheetView[0].showGridLines = True
        for col in range(1, ws_nominal.max_column + 1):
            cell = ws_nominal.cell(row=1, column=col)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal='left', vertical='center')
        for row in range(2, ws_nominal.max_row + 1):
            for col in range(1, ws_nominal.max_column + 1):
                cell = ws_nominal.cell(row=row, column=col)
                cell.font = font_body
                cell.border = box_border
        for col in ws_nominal.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_nominal.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output.seek(0)
    return output

@app.post("/api/parte/procesar-y-exportar")
async def procesar_y_exportar_parte(request: ParteRequest):
    if not request.brigada:
        raise HTTPException(status_code=400, detail="La lista de la brigada está vacía.")
    datos_nominales = [g.model_dump() for g in request.brigada]
    df_nominal = pd.DataFrame(datos_nominales)

    secciones_invalidas = set(df_nominal['seccion']) - set(SECCIONES_VALIDAS)
    if secciones_invalidas:
        raise HTTPException(status_code=400, detail=f"Secciones inválidas: {list(secciones_invalidas)}")
    estados_invalidos = set(df_nominal['estado']) - set(ESTADOS_VALIDOS)
    if estados_invalidos:
        raise HTTPException(status_code=400, detail=f"Estados inválidos: {list(estados_invalidos)}")

    conteo_real = df_nominal.groupby(['seccion', 'estado']).size().unstack(fill_value=0)
    matriz_final = conteo_real.reindex(index=SECCIONES_VALIDAS, columns=ESTADOS_VALIDOS, fill_value=0)
    matriz_final['TOTAL ALTA'] = matriz_final.sum(axis=1)
    matriz_final.loc['TOTAL GENERAL'] = matriz_final.sum(axis=0)

    excel_binario = generar_excel_profesional(df_nominal, matriz_final)
    nombre_descarga = f"{request.nombre_archivo.replace(' ', '_')}.xlsx"
    return StreamingResponse(
        excel_binario,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nombre_descarga}"}
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)