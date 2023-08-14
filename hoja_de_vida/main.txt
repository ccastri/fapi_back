
from typing import Union
from fastapi import FastAPI
from pydantic import BaseModel
from .hoja_de_vida import hdv
from .middlewares import upload_image
from . import users
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(docs_url="/api/docs", openapi_url="/api/openapi.json")
# encoded_url = os.environ.get("NEXT_PUBLIC_VERCEL_URL")
# decoded_url = urllib.parse.unquote(encoded_url)

# Cors enabled to receive http request from nextjs frontend
# origins = [
#     "http://localhost:3000",
# ]

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Trial fastapi model
class Item(BaseModel):
    name: str
    price: float
    is_offer: Union[bool, None] = None


app.include_router(hdv.router)
app.include_router(upload_image.router)
app.include_router(users.router)


@app.get("/api/hola")
def read_root():
    return {"Hello": "World"}

 # at last, the bottom of the file/module
# if __name__ == "__main__":
#     uvicorn.run(app, host="127.0.0.1", port=8000)

# import os
# import databases
# import sqlalchemy
# from typing import List
# from fastapi import FastAPI
# from pydantic import BaseModel
# from dotenv import load_dotenv
# from fastapi.middleware.cors import CORSMiddleware

# load_dotenv()

# DATABASE_URL = os.getenv("DATABASE_URL")

# database = databases.Database(DATABASE_URL)

# metadata = sqlalchemy.MetaData()

# notes = sqlalchemy.Table(
#     "notes",
#     metadata,
#     sqlalchemy.Column("id", sqlalchemy.Integer, primary_key=True),
#     sqlalchemy.Column("text", sqlalchemy.un valor),
#     sqlalchemy.Column("completed", sqlalchemy.Boolean),
# )


# engine = sqlalchemy.create_engine(
#     DATABASE_URL
# )
# # metadata.create_all(engine)


# class NoteIn(BaseModel):
#     text: str
#     completed: bool


# class Note(BaseModel):
#     id: int
#     text: str
#     completed: bool


# app = FastAPI()

# origins = [
#     "http://localhost",
#     "http://localhost:8080",
#     "http://localhost:3000"
# ]

# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=origins,
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )


# @app.on_event("startup")
# async def startup():
#     await database.connect()


# @app.on_event("shutdown")
# async def shutdown():
#     await database.disconnect()


# @app.get("/notes/", response_model=List[Note])
# async def read_notes():
#     query = notes.select()
#     return await database.fetch_all(query)


# @app.post("/notes/", response_model=Note)
# async def create_note(note: NoteIn):
#     print(note)
#     query = notes.insert().values(text=note.text, completed=note.completed)
#     last_record_id = await database.execute(query)
#     return {**note.dict(), "id": last_record_id}



#   "departamento": "scssd",
#   "municipio": "un valor",
#   "entidad": "un valor",
#   "correo": "un valor",
#   "direccion": "un valor",
#   "telefono": "un valor",
#   "equipo": "un valor",
#   "marca": "un valor",
#   "modelo": "un valor",
#   "serie": "un valor",
#   "activoFijo": "un valor",
#   "registroSanitario": "un valor",
#   "ubicacion": "un valor",
#   "proveedor": "un valor",
#   "img": "un valor",
#   "AdquisitionWay": "un valor",
#   "yearOfFabrication": "2023-08-10",
#   "boughtDate": "2023-08-10",
#   "installationDate": "2023-08-10",
#   "warrantyEnd": "2023-08-10",
#   "fabricante": "un valor",
#   "tension": "un valor",
#   "potencia": "un valor",
#   "presion": "un valor",
#   "corriente": "un valor",
#   "frecuencia": "un valor",
#   "rangoTemperatura": "un valor",
#   "peso": "un valor",
#   "velocidad": "un valor",
#   "tecnologiaPredominante": "un valor",
#   "rangoVoltaje": "un valor",
#   "rangoPresion": "un valor",
#   "rangoHumedad": "un valor",
#   "rangoCorriente": "un valor",
#   "rangoFrecuencia": "un valor",
#   "diagnostico": "un valor",
#   "rehabilitacion": "un valor",
#   "prevencion": "un valor",
#   "tratamientoVida": "un valor",
#   "analisisLaboratorio": "un valor",
#   "clasificacion": "un valor",
#   "periodicidad": "un valor",
#   "calibracion": "un valor"