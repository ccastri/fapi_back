from fastapi import FastAPI, APIRouter, HTTPException
from pydantic import BaseModel
from openpyxl import load_workbook
from fastapi.responses import FileResponse
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

router = APIRouter(prefix="/hdv", tags=["add"])


class FormData(BaseModel):
    departamento: str
    municipio: str
    entidad: str
    correo: str
    direccion: str
    telefono: str


# Route for filling out the excel template with the user's inputs
@router.post("/fill_excel")
async def fill_excel(data: FormData):
    try:
        # Cargar el archivo Excel hoja_De_vida.xlsx
        workbook = load_workbook("hdv.xlsx")
        sheet = workbook.active

        # Llenar los campos del archivo Excel con los datos recibidos tipo AllFormData
        # provenientes de Next JS:
        # I. Ubicacion geográfica
        sheet["C6"] = data.departamento
        sheet["C7"] = data.municipio
        sheet["C8"] = data.entidad
        sheet["C9"] = data.correo
        sheet["C10"] = data.direccion
        sheet["C11"] = data.telefono
        # Llenar los otros campos del formulario
        # II. ...
        # Guardar el archivo Excel modificado en un directorio temporal
        temp_excel_file_path = "hdv_temp.xlsx"
        workbook.save(temp_excel_file_path)
        return FileResponse(temp_excel_file_path, filename="hdv.xlsx")
    except Exception as e:
        # Handle the error and provide a meaningful error message
        error_message = f"Error al procesar los datos: {str(e)}"
        raise HTTPException(status_code=500, detail=error_message)
