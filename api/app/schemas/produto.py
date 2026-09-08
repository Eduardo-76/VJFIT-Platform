from datetime import datetime

from pydantic import BaseModel, ConfigDict


class ProdutoCreate(BaseModel):
    categoria_id: int
    nome: str
    descricao: str | None = None


class ProdutoUpdate(BaseModel):
    categoria_id: int
    nome: str
    descricao: str | None = None


class ProdutoResponse(BaseModel):
    id: int
    categoria_id: int
    nome: str
    descricao: str | None
    ativo: bool
    created_at: datetime
    updated_at: datetime

    model_config = ConfigDict(from_attributes=True)