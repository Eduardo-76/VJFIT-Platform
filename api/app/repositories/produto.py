from sqlalchemy import select
from sqlalchemy.orm import Session

from app.models.produto import Produto


def buscar_por_nome(
    db: Session,
    nome: str,
    categoria_id: int,
    produto_id: int | None = None,
):
    query = select(Produto).where(
        Produto.nome == nome,
        Produto.categoria_id == categoria_id,
    )

    if produto_id is not None:
        query = query.where(Produto.id != produto_id)

    return db.execute(query).scalar_one_or_none()


def criar_produto(
    db: Session,
    categoria_id: int,
    nome: str,
    descricao: str | None = None,
):
    produto = Produto(
        categoria_id=categoria_id,
        nome=nome,
        descricao=descricao,
    )

    db.add(produto)
    db.commit()
    db.refresh(produto)

    return produto


def buscar_por_id(
    db: Session,
    produto_id: int,
):
    query = select(Produto).where(
        Produto.id == produto_id
    )

    return db.execute(query).scalar_one_or_none()


def listar_produtos(
    db: Session,
    apenas_ativos: bool = True,
):
    query = select(Produto)

    if apenas_ativos:
        query = query.where(
            Produto.ativo.is_(True)
        )

    query = query.order_by(Produto.id)

    return db.execute(query).scalars().all()


def atualizar_produto(
    db: Session,
    produto: Produto,
    nome: str,
    categoria_id: int,
    descricao: str | None,
):
    produto.nome = nome
    produto.categoria_id = categoria_id
    produto.descricao = descricao

    db.commit()
    db.refresh(produto)

    return produto


def desativar_produto(
    db: Session,
    produto: Produto,
):
    produto.ativo = False

    db.commit()
    db.refresh(produto)

    return produto


def ativar_produto(
    db: Session,
    produto: Produto,
):
    produto.ativo = True

    db.commit()
    db.refresh(produto)

    return produto