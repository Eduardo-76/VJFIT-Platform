from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.database import SessionLocal

from app.schemas.sku import (
    SKUCreate,
    SKUUpdate,
    SKUResponse,
)

from app.services.sku import (
    criar_novo_sku,
    buscar_sku_por_id,
    listar_todos_skus,
    atualizar_sku_service,
    SKUNotFoundError,
    SKUConflictError,
    SKUValidationError,
)


router = APIRouter(
    prefix="/skus",
    tags=["SKUs"]
)


def get_db():
    db = SessionLocal()

    try:
        yield db
    finally:
        db.close()


@router.post(
    "/",
    response_model=SKUResponse
)
def criar_sku(
    sku: SKUCreate,
    db: Session = Depends(get_db)
):
    try:
        return criar_novo_sku(
            db=db,
            produto_id=sku.produto_id,
            codigo=sku.codigo,
            cor=sku.cor,
            tamanho=sku.tamanho,
            publico=sku.publico,
            quantidade=sku.quantidade,
            preco_custo=sku.preco_custo,
            preco_venda=sku.preco_venda,
            em_promocao=sku.em_promocao,
            preco_promocional=sku.preco_promocional,
        )

    except SKUNotFoundError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )

    except SKUConflictError as erro:
        raise HTTPException(
            status_code=409,
            detail=str(erro)
        )

    except SKUValidationError as erro:
        raise HTTPException(
            status_code=422,
            detail=str(erro)
        )


@router.get(
    "/",
    response_model=list[SKUResponse]
)
def listar_skus(
    db: Session = Depends(get_db)
):
    return listar_todos_skus(
        db=db
    )


@router.get(
    "/{sku_id}",
    response_model=SKUResponse
)
def buscar_sku(
    sku_id: int,
    db: Session = Depends(get_db)
):
    try:
        return buscar_sku_por_id(
            db=db,
            sku_id=sku_id
        )

    except SKUNotFoundError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )


@router.put(
    "/{sku_id}",
    response_model=SKUResponse
)
def atualizar_sku(
    sku_id: int,
    sku: SKUUpdate,
    db: Session = Depends(get_db)
):
    try:
        return atualizar_sku_service(
            db=db,
            sku_id=sku_id,
            produto_id=sku.produto_id,
            codigo=sku.codigo,
            cor=sku.cor,
            tamanho=sku.tamanho,
            publico=sku.publico,
            quantidade=sku.quantidade,
            preco_custo=sku.preco_custo,
            preco_venda=sku.preco_venda,
            em_promocao=sku.em_promocao,
            preco_promocional=sku.preco_promocional,
        )

    except SKUNotFoundError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )

    except SKUConflictError as erro:
        raise HTTPException(
            status_code=409,
            detail=str(erro)
        )

    except SKUValidationError as erro:
        raise HTTPException(
            status_code=422,
            detail=str(erro)
        )