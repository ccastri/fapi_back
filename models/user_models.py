from dataclasses import dataclass


@dataclass
class User:
    empresa: str
    nit: str
    ciudad: str
    departamento: str
    correo: str
    contraseña: str
    confirmarContraseña: str
    role: str
    tos: bool
