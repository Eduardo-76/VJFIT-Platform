from datetime import datetime
from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field


class SKUCreate(BaseModel):
    produto_id: int

    codigo: str = Field(
        min_length=1,
        max_length=50
    )

    cor: str = Field(
        min_length=1,
        max_length=50
    )

    tamanho: str = Field(
        min_length=1,
        max_length=50
    )

    publico: str = Field(
        min_length=1,
        max_length=20
    )

    quantidade: int = Field(
        default=0,
        ge=0
    )

    preco_custo: Decimal = Field(
        ge=0
    )

    preco_venda: Decimal = Field(
        ge=0
    )

    em_promocao: bool = False

    preco_promocional: Decimal | None = Field(
        default=None,
        ge=0
    )


class SKUUpdate(BaseModel):
    produto_id: int

    codigo: str = Field(
        min_length=1,
        max_length=50
    )

    cor: str = Field(
        min_length=1,
        max_length=50
    )

    tamanho: str = Field(
        min_length=1,
        max_length=50
    )

    publico: str = Field(
        min_length=1,
        max_length=20
    )

    quantidade: int = Field(
        ge=0
    )

    preco_custo: Decimal = Field(
        ge=0
    )

    preco_venda: Decimal = Field(
        ge=0
    )

    em_promocao: bool

    preco_promocional: Decimal | None = Field(
        default=None,
        ge=0
    )


class SKUResponse(BaseModel):
    id: int
    produto_id: int
    codigo: str
    cor: str
    tamanho: str
    publico: str
    quantidade: int
    preco_custo: Decimal
    preco_venda: Decimal
    em_promocao: bool
    preco_promocional: Decimal | None
    created_at: datetime
    updated_at: datetime

    model_config = ConfigDict(
        from_attributes=True
    )