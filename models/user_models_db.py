from sqlalchemy import Column, String, Boolean, Integer
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()


class UserDB(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, autoincrement=True)
    empresa = Column(String)
    nit = Column(String)
    ciudad = Column(String)
    departamento = Column(String)
    correo = Column(String)
    contraseña = Column(String)
    confirmarContraseña = Column(String)
    role = Column(String)
    tos = Column(Boolean)

    def __repr__(self):
        return f"<User(id={self.id}, correo='{self.correo}')>"
