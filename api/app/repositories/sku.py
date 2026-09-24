from sqlalchemy import select
from sqlalchemy.orm import Session

from app.models.sku import SKU


def buscar_por_codigo(
    db: Session,
    codigo: str
) -> SKU | None:
    query = select(SKU).where(
        SKU.codigo == codigo
    )

    return db.execute(query).scalar_one_or_none()


def buscar_por_id(
    db: Session,
    sku_id: int
) -> SKU | None:
    query = select(SKU).where(
        SKU.id == sku_id
    )

    return db.execute(query).scalar_one_or_none()


def listar_skus(
    db: Session
) -> list[SKU]:
    query = select(SKU).order_by(SKU.id)

    return list(db.execute(query).scalars().all())


def criar_sku(
    db: Session,
    sku: SKU
) -> SKU:
    db.add(sku)
    db.commit()
    db.refresh(sku)

    return sku


def atualizar_sku(
    db: Session,
    sku: SKU
) -> SKU:
    db.commit()
    db.refresh(sku)

    return sku


def desativar_sku(
    db: Session,
    sku: SKU
) -> SKU:
    db.commit()
    db.refresh(sku)

    return sku

def buscar_por_variacao(
    db: Session,
    produto_id: int,
    cor: str,
    tamanho: str,
    publico: str,
    ignorar_sku_id: int | None = None
) -> SKU | None:

    query = select(SKU).where(
        SKU.produto_id == produto_id,
        SKU.cor == cor,
        SKU.tamanho == tamanho,
        SKU.publico == publico
    )

    if ignorar_sku_id is not None:
        query = query.where(
            SKU.id != ignorar_sku_id
        )

    return db.execute(query).scalar_one_or_none()