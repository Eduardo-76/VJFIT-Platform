from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from app.database import SessionLocal

from app.schemas.produto import (
    ProdutoCreate,
    ProdutoUpdate,
    ProdutoResponse,
)

from app.services.produto import (
    criar_novo_produto,
    listar_produtos_service,
    buscar_produto_por_id,
    atualizar_produto_service,
    desativar_produto_service,
    ativar_produto_service,
)


router = APIRouter(
    prefix="/produtos",
    tags=["Produtos"],
)


def get_db():
    db = SessionLocal()

    try:
        yield db
    finally:
        db.close()


@router.post(
    "/",
    response_model=ProdutoResponse,
)
def criar_produto(
    produto: ProdutoCreate,
    db: Session = Depends(get_db),
):
    try:
        return criar_novo_produto(
            db=db,
            categoria_id=produto.categoria_id,
            nome=produto.nome,
            descricao=produto.descricao,
        )

    except ValueError as erro:
        raise HTTPException(
            status_code=409,
            detail=str(erro),
        )


@router.get(
    "/",
    response_model=list[ProdutoResponse],
)
def listar_produtos(
    apenas_ativos: bool = True,
    db: Session = Depends(get_db),
):
    return listar_produtos_service(
        db=db,
        apenas_ativos=apenas_ativos,
    )


@router.get(
    "/{produto_id}",
    response_model=ProdutoResponse,
)
def buscar_produto(
    produto_id: int,
    db: Session = Depends(get_db),
):
    try:
        return buscar_produto_por_id(
            db=db,
            produto_id=produto_id,
        )

    except ValueError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro),
        )


@router.put(
    "/{produto_id}",
    response_model=ProdutoResponse,
)
def atualizar_produto(
    produto_id: int,
    produto: ProdutoUpdate,
    db: Session = Depends(get_db),
):
    try:
        return atualizar_produto_service(
            db=db,
            produto_id=produto_id,
            categoria_id=produto.categoria_id,
            nome=produto.nome,
            descricao=produto.descricao,
        )

    except ValueError as erro:
        mensagem = str(erro)

        if "não encontrado" in mensagem.lower():
            status_code = 404
        else:
            status_code = 409

        raise HTTPException(
            status_code=status_code,
            detail=mensagem,
        )


@router.delete(
    "/{produto_id}",
    response_model=ProdutoResponse,
)
def desativar_produto(
    produto_id: int,
    db: Session = Depends(get_db),
):
    try:
        return desativar_produto_service(
            db=db,
            produto_id=produto_id,
        )

    except ValueError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro),
        )


@router.patch(
    "/{produto_id}/ativar",
    response_model=ProdutoResponse,
)
def ativar_produto(
    produto_id: int,
    db: Session = Depends(get_db),
):
    try:
        return ativar_produto_service(
            db=db,
            produto_id=produto_id,
        )

    except ValueError as erro:
        raise HTTPException(
            status_code=404,
            detail=str(erro),
        )