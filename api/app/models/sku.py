from datetime import datetime
from decimal import Decimal

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    DateTime,
    ForeignKey,
    Integer,
    Numeric,
    String,
    Text,
    func,
)
from sqlalchemy.orm import Mapped, mapped_column

from app.database import Base


class SKU(Base):
    __tablename__ = "sku"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True
    )

    produto_id: Mapped[int] = mapped_column(
        ForeignKey("produto.id"),
        nullable=False
    )

    codigo: Mapped[str] = mapped_column(
        String(50),
        unique=True,
        nullable=False
    )

    cor: Mapped[str] = mapped_column(
        String(50),
        nullable=False
    )

    tamanho: Mapped[str] = mapped_column(
        String(50),
        nullable=False
    )

    publico: Mapped[str] = mapped_column(
        String(20),
        nullable=False
    )

    quantidade: Mapped[int] = mapped_column(
        Integer,
        default=0,
        nullable=False
    )

    preco_custo: Mapped[Decimal] = mapped_column(
        Numeric(10, 2),
        nullable=False
    )

    preco_venda: Mapped[Decimal] = mapped_column(
        Numeric(10, 2),
        nullable=False
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        server_default=func.now(),
        nullable=False
    )

    updated_at: Mapped[datetime] = mapped_column(
        DateTime,
        server_default=func.now(),
        nullable=False
    )

    em_promocao: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
        nullable=False
    )

    preco_promocional: Mapped[Decimal | None] = mapped_column(
        Numeric(10, 2),
        nullable=True
    )

    __table_args__ = (
        CheckConstraint(
            "preco_custo >= 0",
            name="sku_preco_custo_check"
        ),
        CheckConstraint(
            "preco_venda >= 0",
            name="sku_preco_venda_check"
        ),
        CheckConstraint(
            "quantidade >= 0",
            name="sku_quantidade_check"
        ),
        CheckConstraint(
            "em_promocao = false "
            "OR preco_promocional IS NOT NULL "
            "AND preco_promocional <= preco_venda",
            name="sku_promocao_check"
        ),
    )