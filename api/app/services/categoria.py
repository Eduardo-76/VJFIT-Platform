from sqlalchemy.orm import Session

from app.models.categoria import Categoria
from app.repositories.categoria import (
    buscar_por_nome,
    criar_categoria,
    buscar_por_id,
    listar_categorias as listar_categorias_repository,
    atualizar_categoria,
    ativar_categoria,
    desativar_categoria
)


class CategoriaNaoEncontradaError(ValueError):
    pass


class CategoriaJaCadastradaError(ValueError):
    pass


def criar_nova_categoria(
    db: Session,
    nome: str
) -> Categoria:

    nome_normalizado = nome.strip().upper()

    categoria_existente = buscar_por_nome(
        db=db,
        nome=nome_normalizado
    )

    if categoria_existente:
        raise CategoriaJaCadastradaError(
            "Categoria já cadastrada."
        )

    categoria = Categoria(
        nome=nome_normalizado
    )

    return criar_categoria(
        db=db,
        categoria=categoria
    )


def listar_categorias(db: Session) -> list[Categoria]:
    return listar_categorias_repository(db=db)


def buscar_categoria_por_id(
    db: Session,
    categoria_id: int
) -> Categoria:

    categoria = buscar_por_id(
        db=db,
        categoria_id=categoria_id
    )

    if categoria is None:
        raise CategoriaNaoEncontradaError(
            "Categoria não encontrada."
        )

    return categoria


def atualizar_categoria_service(
    db: Session,
    categoria_id: int,
    nome: str
) -> Categoria:

    nome_normalizado = nome.strip().upper()

    categoria = buscar_por_id(
        db=db,
        categoria_id=categoria_id
    )

    if categoria is None:
        raise CategoriaNaoEncontradaError(
            "Categoria não encontrada."
        )

    categoria_existente = buscar_por_nome(
        db=db,
        nome=nome_normalizado
    )

    if (
        categoria_existente
        and categoria_existente.id != categoria_id
    ):
        raise CategoriaJaCadastradaError(
            "Categoria já cadastrada."
        )

    return atualizar_categoria(
        db=db,
        categoria=categoria,
        nome=nome_normalizado
    )

def ativar_categoria_service(
    db: Session,
    categoria_id: int
) -> Categoria:

    categoria = buscar_por_id(
        db=db,
        categoria_id=categoria_id
    )

    if categoria is None:
        raise ValueError("Categoria não encontrada.")

    if categoria.ativa:
        raise ValueError("Categoria já está ativa.")

    return ativar_categoria(
        db=db,
        categoria=categoria
    )


def desativar_categoria_service(
    db: Session,
    categoria_id: int
) -> Categoria:

    categoria = buscar_por_id(
        db=db,
        categoria_id=categoria_id
    )

    if categoria is None:
        raise ValueError("Categoria não encontrada.")

    if not categoria.ativa:
        raise ValueError("Categoria já está inativa.")

    return desativar_categoria(
        db=db,
        categoria=categoria
    )