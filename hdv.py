from fastapi import FastAPI, APIRouter, HTTPException
from pydantic import BaseModel
from openpyxl import load_workbook
from fastapi.responses import FileResponse
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

router = APIRouter(prefix="/hdv", tags=["add"])


@router.get("/numbers")
def add_numbers():
    return {"message": "we're in numbers"}


class FormData(BaseModel):
    departamento: str
    municipio: str
    entidad: str
    correo: str
    direccion: str
    telefono: str


def create_pdf_from_excel(excel_file_path, pdf_file_path):
    # Read the Excel file
    workbook = load_workbook(excel_file_path)
    sheet = workbook.active
    data = []

    # Extract data from the Excel and add it to the list
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    # Create the PDF
    doc = SimpleDocTemplate(pdf_file_path, pagesize=letter)
    table = Table(data)

    # Add borders to the table cells
    style = TableStyle([("GRID", (0, 0), (-1, -1), 1, (0.5, 0.5, 0.5))])
    table.setStyle(style)

    # Add the table to the PDF
    elements = [table]
    doc.build(elements)


@router.post("/api/fill_excel")
async def fill_excel(data: FormData):
    try:
        # Cargar el archivo Excel hoja_De_vida.xlsx
        workbook = load_workbook("hdv.xlsx")
        sheet = workbook.active

        # Llenar los campos del archivo Excel con los datos recibidos
        sheet["C6"] = data.departamento
        sheet["C7"] = data.municipio
        sheet["C8"] = data.entidad
        sheet["C9"] = data.correo
        sheet["C10"] = data.direccion
        sheet["C11"] = data.telefono
        # Llenar los otros campos del formulario

        # Guardar el archivo Excel modificado en un directorio temporal
        temp_excel_file_path = "hdv_temp.xlsx"
        workbook.save(temp_excel_file_path)

        # Create the PDF from the Excel file
        temp_pdf_file_path = "hdv_temp.pdf"
        create_pdf_from_excel(temp_excel_file_path, temp_pdf_file_path)

        # Respond with both the Excel and PDF files
        return {
            "message": "Data processed successfully",
            "excel_file": FileResponse(temp_excel_file_path, filename="hdv.xlsx"),
            "pdf_file": FileResponse(temp_pdf_file_path, filename="hdv.pdf"),
        }
    except Exception as e:
        # Handle the error and provide a meaningful error message
        error_message = f"Error al procesar los datos: {str(e)}"
        raise HTTPException(status_code=500, detail=error_message)
