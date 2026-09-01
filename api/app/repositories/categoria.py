from sqlalchemy.orm import Session

from app.models.categoria import Categoria


def buscar_por_nome(
    db: Session,
    nome: str
) -> Categoria | None:
    return (
        db.query(Categoria)
        .filter(Categoria.nome == nome)
        .first()
    )

def criar_categoria(
    db: Session,
    categoria: Categoria
) -> Categoria:
    db.add(categoria)
    db.commit()
    db.refresh(categoria)

    return categoria

def listar_categorias(db: Session) -> list[Categoria]:
    return db.query(Categoria).all()

def buscar_por_id(
    db: Session,
    categoria_id: int
) -> Categoria | None:
    return (
        db.query(Categoria)
        .filter(Categoria.id == categoria_id)
        .first()
    )

def atualizar_categoria(
    db: Session,
    categoria: Categoria,
    nome: str
) -> Categoria:

    categoria.nome = nome

    db.commit()
    db.refresh(categoria)

    return categoria

def desativar_categoria(
    db: Session,
    categoria: Categoria
) -> Categoria:

    categoria.ativa = False

    db.commit()
    db.refresh(categoria)

    return categoria

def ativar_categoria(
    db: Session,
    categoria: Categoria
) -> Categoria:

    categoria.ativa = True

    db.commit()
    db.refresh(categoria)

    return categoria