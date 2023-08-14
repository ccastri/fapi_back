from fastapi import FastAPI, APIRouter, HTTPException, File, UploadFile, Form, Depends
from pydantic import BaseModel, Field
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from fastapi.responses import FileResponse
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from datetime import date, datetime
import logging
import os
import tempfile
from typing import List 
import json
# from dataclasses import dataclass
import time
from PIL import Image as PILImage
# import io

router = APIRouter(prefix="/api/hdv", tags=["add"])


# Got to fix this model to be on the hdv_sb_models.py
# @dataclass
class FormData(BaseModel):
    class Config:
        arbitrary_types_allowed = True
    # departamento: str
    # img: UploadFile
    # data: dict
    # departamento: str
    #  = Form(...)
    # municipio: str
    #  = Form(...)
    # entidad: str
    #  = Form(...)
    # correo: str
    #  = Form(...)
    # direccion: str
    #  = Form(...)
    # telefono: str
    #  = Form(...)


    # equipo: str
    
    # marca: str
    # modelo: str
    # serie: str
    # activoFijo: str
    # registroSanitario: str
    # ubicacion: str
    # proveedor: str
    

    # AdquisitionWay: str
    # yearOfFabrication: date
    # boughtDate: date
    # installationDate: date
    # warrantyEnd: date
    # fabricante: str

    # tension: str
    # potencia: str
    # presion: str
    # corriente: str
    # frecuencia: str
    # rangoTemperatura: str
    # peso: str
    # velocidad: str
    # tecnologiaPredominante: str

    # rangoVoltaje: str
    # rangoPresion: str
    # rangoHumedad: str
    # rangoCorriente: str
    # rangoFrecuencia: str

    # diagnostico: str
    # rehabilitacion: str
    # prevencion: str
    # tratamientoVida: str
    # analisisLaboratorio: str
    # filteredInputs: List[str] = Field(default_factory=list)
    # clasificacion: str = None
    # periodicidad: str = None
    # calibracion: str = None
    # copiaRegistroSanitario: str = None
    # copiaRegistroImportacion: str = None
    # copiaFactura: str = None
    # copiaIngresoAlmacen: str = None
    # copiaActaReciboSatisfaccion: str = None


# Route for filling out the excel template with the user's inputs
@router.post("/fill_excel")
async def fill_excel(
    img: UploadFile = File(...),
    data: str = Form(...)
    ):
    # json_data = data.data
    # departamento = json_data.get("departamento")
    # municipio = json_data.get("municipio")
#     {
#   "departamento": "un valor",
#   "municipio": "un valor",
#   "entidad": "un valor",
#   "correo": "un valor",
#   "direccion": "un valor",
#   "telefono": "un valor"
# }

    try:
        start_time = time.time()
        # data_dict = json.loads(data)
        # data_dict = json.loads(form_data.data)
        data_dict = json.loads(data)
        print(data_dict['departamento'])
        print(img)
        # departamento = data["departamento"]
        # municipio = data_dict["municipio"]
        # entidad = data_dict["entidad"]
        # correo = data_dict["correo"]
        # direccion = data_dict["direccion"]
        # telefono = data_dict["telefono"]
        # print(telefono)
        # print(data.img)
        module_directory = os.path.dirname(__file__)
        excel_template_path = os.path.join(module_directory, "hdv.xlsx")
        # Cargar el archivo Excel hoja_De_vida.xlsx
        workbook = load_workbook(excel_template_path)
        sheet = workbook.active

        # Llenar los campos del archivo Excel con los datos recibidos tipo AllFormData
        # provenientes de Next JS:
        # I. Ubicacion geográfica
        sheet["C6"] =  data_dict['departamento']
        sheet["C7"] =  data_dict['municipio']
        sheet["C8"] =  data_dict['entidad']
        sheet["C9"] =  data_dict['correo']
        sheet["C10"] = data_dict['direccion']
        sheet["C11"] = data_dict['telefono']
        sheet["C13"] = data_dict['equipo']
        sheet["C14"] = data_dict['marca']
        sheet["C15"] = data_dict['modelo']
        sheet["C16"] = data_dict['serie']
        sheet["C17"] = data_dict['activoFijo']
        sheet["C18"] = data_dict['registroSanitario']
        sheet["C19"] = data_dict['ubicacion']
        sheet["C20"] = data_dict['proveedor']
        temp_dir = tempfile.mkdtemp()
        
        # Generate a temporary file path
        temp_image_path = os.path.join(temp_dir, img.filename)
        with open(temp_image_path, "wb") as img_file:
           img_file.write(img.file.read())
        
                # Open and resize the image using PIL
        pil_image = PILImage.open(temp_image_path)
        pil_image.thumbnail((800, 800))  # Resize the image
        
        # Save the compressed image back to the temporary path
        compressed_image_path = os.path.join(temp_dir, "compressed_image.jpg")
        pil_image.save(compressed_image_path, format="JPEG", quality=85)  # Adjust quality as needed
        
        # Add the compressed image to the Excel
        img = Image(compressed_image_path)
        img.width = 250  # Set the width of the image
        img.height = 155  # Set the height of the image
        img_anchor = "H5"  # The top-left cell for the image
        sheet.add_image(img, img_anchor)
        
        
        sheet["G13"] = data_dict['AdquisitionWay']
        data_dict['yearOfFabrication'] = datetime.strptime(data_dict['yearOfFabrication'], '%Y-%m-%d')
        data_dict['boughtDate'] = datetime.strptime(data_dict['boughtDate'], '%Y-%m-%d')
        data_dict['installationDate'] = datetime.strptime(data_dict['installationDate'], '%Y-%m-%d')
        data_dict['warrantyEnd'] = datetime.strptime(data_dict['warrantyEnd'], '%Y-%m-%d')
        # Extract year, month, and day
        year = data_dict['yearOfFabrication'].year
        month = data_dict['yearOfFabrication'].month
        day = data_dict['yearOfFabrication'].day
        sheet["H16"] = day
        sheet["I16"] = month
        sheet["J16"] = year
        year_boughtDate = data_dict['boughtDate'].year
        month_boughtDate = data_dict['boughtDate'].month
        day_boughtDate = data_dict['boughtDate'].day
        sheet["H17"] = day_boughtDate
        sheet["I17"] = month_boughtDate
        sheet["J17"] = year_boughtDate
        year_installationDate = data_dict['installationDate'].year
        month_installationDate = data_dict['installationDate'].month
        day_installationDate = data_dict['installationDate'].day
        sheet["H18"] = day_installationDate
        sheet["I18"] = month_installationDate
        sheet["J18"] = year_installationDate
        year_warrantyEnd = data_dict['warrantyEnd'].year
        month_warrantyEnd = data_dict['warrantyEnd'].month
        day_warrantyEnd = data_dict['warrantyEnd'].day
        sheet["H19"] = day_warrantyEnd
        sheet["I19"] = month_warrantyEnd
        sheet["J19"] = year_warrantyEnd

        # # sheet["G17"] = data_dictdata_dictdata_dict.warrantyEnd
        # sheet["F20"] = data_dictdata_dict.fabricante

        sheet["C23"] = data_dict['tension']
        sheet["C24"] = data_dict['potencia']
        sheet["C25"] = data_dict['presion']
        sheet["G23"] = data_dict['corriente']
        sheet["G24"] = data_dict['frecuencia']
        sheet["G25"] = data_dict['rangoTemperatura']
        sheet["J23"] = data_dict['peso']
        sheet["J24"] = data_dict['velocidad']
        sheet["J25"] = data_dict['tecnologiaPredominante']

        sheet["C27"] = data_dict['rangoVoltaje']
        sheet["C28"] = data_dict['rangoPresion']
        sheet["C29"] = data_dict['rangoHumedad']
        sheet["C30"] = data_dict['rangoCorriente']
        sheet["C31"] = data_dict['rangoFrecuencia']

        sheet["H27"] = data_dict['diagnostico']
        sheet["H28"] = data_dict['rehabilitacion']
        sheet["H29"] = data_dict['prevencion']
        sheet["H30"] = data_dict['tratamientoVida']
        sheet["H31"] = data_dict['analisisLaboratorio']

        sheet["A34"] = data_dict['clasificacion']
        sheet["E34"] = data_dict['periodicidad']
        sheet["H34"] = data_dict['calibracion']
    #   Define the starting row for columns A and E
        start_row_A = 36
        start_row_E = 36
        # print(data_dict['filteredInputs'])
        filtered_inputs = data_dict.get('filteredInputs', [])
        # Loop through the filteredInputs list and assign them to cells in the specified pattern
        for idx, input_value in enumerate(filtered_inputs):
            if idx < 3:
                cell_row_A = start_row_A + idx
                cell = sheet[f"A{cell_row_A}"]
                cell.value = input_value
            elif idx < 6:
                cell_row_E = start_row_E + (idx - 3)
                cell = sheet[f"E{cell_row_E}"]
                cell.value = input_value

         
      
      
        # field_value_mappings = {
        #     "copiaRegistroSanitario": data_dict.copiaRegistroSanitario,
        #     "copiaRegistroImportacion": data_dict.copiaRegistroImportacion,
        #     "copiaFactura": data_dict.copiaFactura,
        #     "copiaIngresoAlmacen": data_dict.copiaIngresoAlmacen,
        #     "copiaActaReciboSatisfaccion": data_dict.copiaActaReciboSatisfaccion,
        #     # Add other fields and values here
        # }
        # start_row = 53
        # end_row = 64
        # columns = ["H", "I", "J", "K", "L", "M"]
        
        # for field_name, field_value in field_value_mappings.items():
        #     if hasattr(data, field_name):
        #         for row_idx in range(start_row, end_row + 1):
        #             for col_letter in columns:
        #                 cell_value = sheet[f"{col_letter}{row_idx}"].value
        #                 if cell_value == field_value:
        #                     sheet[f"{col_letter}{row_idx}"] = "X"
        # Llenar los otros campos del formulario
        # II. ...
        # Guardar el archivo Excel modificado en un directorio temporal
        temp_excel_file_path = os.path.join(module_directory, "hdv_temp.xlsx")
        workbook.save(temp_excel_file_path)
        end_time = time.time()
        processing_time = end_time - start_time
        print(processing_time) 
        return FileResponse(temp_excel_file_path, filename="hdv.xlsx")
    
    except FileNotFoundError as file_error:
        error_message = f"Error al abrir el archivo de plantilla: {str(file_error)}"
        logging.error(error_message)
        raise HTTPException(status_code=500, detail=error_message)
    except Exception as e:
        # Handle other exceptions and provide a generic error message
        error_message = f"Error al procesar los datos: {str(e)}"
        logging.error(error_message)
        raise HTTPException(status_code=500, detail="Ocurrió un error en el servidor.")
