from datetime import datetime

from pydantic import BaseModel, ConfigDict


class CategoriaCreate(BaseModel):
    nome: str

class CategoriaUpdate(BaseModel):
    nome: str

class CategoriaResponse(BaseModel):
    id: int
    nome: str
    ativa: bool
    created_at: datetime
    updated_at: datetime

    model_config = ConfigDict(from_attributes=True)