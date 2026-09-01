from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.database import SessionLocal
from app.schemas.categoria import (
    CategoriaCreate,
    CategoriaUpdate,
    CategoriaResponse
)

from app.services.categoria import (
    criar_nova_categoria,
    listar_categorias as listar_categorias_service,
    buscar_categoria_por_id,
    atualizar_categoria_service,
    desativar_categoria_service,
    ativar_categoria_service,
    CategoriaNaoEncontradaError,
    CategoriaJaCadastradaError,
)

router = APIRouter(
    prefix="/categorias",
    tags=["Categorias"]
)


def get_db():
    db = SessionLocal()

    try:
        yield db
    finally:
        db.close()


@router.post(
    "/",
    response_model=CategoriaResponse
)
def criar_categoria(
    categoria: CategoriaCreate,
    db: Session = Depends(get_db)
):
    try:
        return criar_nova_categoria(
            db=db,
            nome=categoria.nome
        )

    except CategoriaJaCadastradaError as erro:
        raise HTTPException(
            status_code=409,
            detail=str(erro)
        )


@router.get("/", response_model=list[CategoriaResponse])
def listar_categorias(db: Session = Depends(get_db)):
    return listar_categorias_service(db=db)


@router.get("/{categoria_id}", response_model=CategoriaResponse)
def buscar_categoria(
    categoria_id: int,
    db: Session = Depends(get_db)
):
    try:
        return buscar_categoria_por_id(
            db=db,
            categoria_id=categoria_id
        )

    except CategoriaNaoEncontradaError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )


@router.put(
    "/{categoria_id}",
    response_model=CategoriaResponse
)
def atualizar_categoria(
    categoria_id: int,
    categoria: CategoriaUpdate,
    db: Session = Depends(get_db)
):
    try:
        return atualizar_categoria_service(
            db=db,
            categoria_id=categoria_id,
            nome=categoria.nome
        )

    except CategoriaNaoEncontradaError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )

    except CategoriaJaCadastradaError as erro:
        raise HTTPException(
            status_code=409,
            detail=str(erro)
        )

@router.delete(
    "/{categoria_id}",
    response_model=CategoriaResponse
)
def desativar_categoria(
    categoria_id: int,
    db: Session = Depends(get_db)
):
    try:
        return desativar_categoria_service(
            db=db,
            categoria_id=categoria_id
        )

    except ValueError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )

@router.patch(
    "/{categoria_id}/ativar",
    response_model=CategoriaResponse
)
def ativar_categoria(
    categoria_id: int,
    db: Session = Depends(get_db)
):
    try:
        return ativar_categoria_service(
            db=db,
            categoria_id=categoria_id
        )

    except ValueError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro)
        )