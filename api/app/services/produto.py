from sqlalchemy.orm import Session

from app.models.categoria import Categoria
from app.models.produto import Produto

from app.repositories.produto import (
    buscar_por_nome,
    criar_produto,
    buscar_por_id,
    listar_produtos,
    atualizar_produto,
    desativar_produto,
    ativar_produto,
)


def criar_novo_produto(
    db: Session,
    categoria_id: int,
    nome: str,
    descricao: str | None = None,
):
    nome = nome.strip().upper()

    if not nome:
        raise ValueError("O nome do produto é obrigatório.")

    categoria = db.get(Categoria, categoria_id)

    if categoria is None:
        raise ValueError("Categoria não encontrada.")

    if not categoria.ativa:
        raise ValueError("Não é possível criar produto em uma categoria inativa.")

    produto_existente = buscar_por_nome(
        db=db,
        nome=nome,
        categoria_id=categoria_id,
    )

    if produto_existente:
        raise ValueError("Produto já cadastrado nesta categoria.")

    return criar_produto(
        db=db,
        categoria_id=categoria_id,
        nome=nome,
        descricao=descricao,
    )


def listar_produtos_service(
    db: Session,
    apenas_ativos: bool = True,
):
    return listar_produtos(
        db=db,
        apenas_ativos=apenas_ativos,
    )


def buscar_produto_por_id(
    db: Session,
    produto_id: int,
):
    produto = buscar_por_id(
        db=db,
        produto_id=produto_id,
    )

    if produto is None:
        raise ValueError("Produto não encontrado.")

    return produto


def atualizar_produto_service(
    db: Session,
    produto_id: int,
    categoria_id: int,
    nome: str,
    descricao: str | None = None,
):
    nome = nome.strip().upper()

    if not nome:
        raise ValueError("O nome do produto é obrigatório.")

    produto = buscar_por_id(
        db=db,
        produto_id=produto_id,
    )

    if produto is None:
        raise ValueError("Produto não encontrado.")

    categoria = db.get(Categoria, categoria_id)

    if categoria is None:
        raise ValueError("Categoria não encontrada.")

    if not categoria.ativa:
        raise ValueError("Não é possível vincular o produto a uma categoria inativa.")

    produto_existente = buscar_por_nome(
        db=db,
        nome=nome,
        categoria_id=categoria_id,
        produto_id=produto_id,
    )

    if produto_existente:
        raise ValueError("Já existe outro produto com este nome nesta categoria.")

    return atualizar_produto(
        db=db,
        produto=produto,
        nome=nome,
        categoria_id=categoria_id,
        descricao=descricao,
    )


def desativar_produto_service(
    db: Session,
    produto_id: int,
):
    produto = buscar_por_id(
        db=db,
        produto_id=produto_id,
    )

    if produto is None:
        raise ValueError("Produto não encontrado.")

    if not produto.ativo:
        raise ValueError("Produto já está inativo.")

    return desativar_produto(
        db=db,
        produto=produto,
    )


def ativar_produto_service(
    db: Session,
    produto_id: int,
):
    produto = buscar_por_id(
        db=db,
        produto_id=produto_id,
    )

    if produto is None:
        raise ValueError("Produto não encontrado.")

    if produto.ativo:
        raise ValueError("Produto já está ativo.")

    categoria = db.get(Categoria, produto.categoria_id)

    if categoria is None:
        raise ValueError("Categoria do produto não encontrada.")

    if not categoria.ativa:
        raise ValueError(
            "Não é possível reativar o produto enquanto a categoria estiver inativa."
        )

    return ativar_produto(
        db=db,
        produto=produto,
    )