from decimal import Decimal

from sqlalchemy.orm import Session

from app.models.sku import SKU
from app.repositories.sku import (
    buscar_por_codigo,
    buscar_por_id,
    buscar_por_variacao,
    listar_skus,
    criar_sku,
    atualizar_sku,
)

from app.repositories.produto import buscar_por_id as buscar_produto_por_id


PUBLICOS_VALIDOS = {
    "MASCULINO",
    "FEMININO",
    "UNISSEX",
    "INFANTIL",
}


class SKUNotFoundError(ValueError):
    pass


class SKUConflictError(ValueError):
    pass


class SKUValidationError(ValueError):
    pass


def validar_publico(publico: str) -> str:
    publico = publico.strip().upper()

    if publico not in PUBLICOS_VALIDOS:
        raise SKUValidationError(
            "Público inválido. Use: "
            "MASCULINO, FEMININO, UNISSEX ou INFANTIL."
        )

    return publico


def validar_promocao(
    em_promocao: bool,
    preco_promocional: Decimal | None,
    preco_venda: Decimal,
) -> None:

    if em_promocao and preco_promocional is None:
        raise SKUValidationError(
            "Preço promocional é obrigatório quando o SKU está em promoção."
        )

    if (
        preco_promocional is not None
        and preco_promocional > preco_venda
    ):
        raise SKUValidationError(
            "O preço promocional não pode ser maior que o preço de venda."
        )


def criar_novo_sku(
    db: Session,
    produto_id: int,
    codigo: str,
    cor: str,
    tamanho: str,
    publico: str,
    quantidade: int,
    preco_custo: Decimal,
    preco_venda: Decimal,
    em_promocao: bool = False,
    preco_promocional: Decimal | None = None,
) -> SKU:

    codigo = codigo.strip().upper()
    cor = cor.strip().upper()
    tamanho = tamanho.strip().upper()
    publico = validar_publico(publico)

    # ---------------------------------------------------------
    # 1. Verificar produto
    # ---------------------------------------------------------

    produto = buscar_produto_por_id(
        db=db,
        produto_id=produto_id
    )

    if not produto:
        raise SKUNotFoundError(
            "Produto não encontrado."
        )

    if not produto.ativo:
        raise SKUValidationError(
            "Não é possível criar SKU para um produto inativo."
        )

    # ---------------------------------------------------------
    # 2. Validar promoção
    # ---------------------------------------------------------

    validar_promocao(
        em_promocao=em_promocao,
        preco_promocional=preco_promocional,
        preco_venda=preco_venda,
    )

    # ---------------------------------------------------------
    # 3. Verificar código
    # ---------------------------------------------------------

    sku_existente = buscar_por_codigo(
        db=db,
        codigo=codigo
    )

    if sku_existente:
        raise SKUConflictError(
            "SKU já cadastrado com este código."
        )

    # ---------------------------------------------------------
    # 4. Verificar duplicidade da variação
    # ---------------------------------------------------------

    variacao_existente = buscar_por_variacao(
        db=db,
        produto_id=produto_id,
        cor=cor,
        tamanho=tamanho,
        publico=publico,
    )

    if variacao_existente:
        raise SKUConflictError(
            "Já existe um SKU com esta combinação "
            "de produto, cor, tamanho e público."
        )

    # ---------------------------------------------------------
    # 5. Criar SKU
    # ---------------------------------------------------------

    sku = SKU(
        produto_id=produto_id,
        codigo=codigo,
        cor=cor,
        tamanho=tamanho,
        publico=publico,
        quantidade=quantidade,
        preco_custo=preco_custo,
        preco_venda=preco_venda,
        em_promocao=em_promocao,
        preco_promocional=preco_promocional,
    )

    return criar_sku(
        db=db,
        sku=sku
    )


def buscar_sku_por_id(
    db: Session,
    sku_id: int
) -> SKU:

    sku = buscar_por_id(
        db=db,
        sku_id=sku_id
    )

    if not sku:
        raise SKUNotFoundError(
            "SKU não encontrado."
        )

    return sku


def listar_todos_skus(
    db: Session
) -> list[SKU]:

    return listar_skus(db=db)


def atualizar_sku_service(
    db: Session,
    sku_id: int,
    produto_id: int,
    codigo: str,
    cor: str,
    tamanho: str,
    publico: str,
    quantidade: int,
    preco_custo: Decimal,
    preco_venda: Decimal,
    em_promocao: bool,
    preco_promocional: Decimal | None,
) -> SKU:

    # ---------------------------------------------------------
    # 1. Buscar SKU
    # ---------------------------------------------------------

    sku = buscar_por_id(
        db=db,
        sku_id=sku_id
    )

    if not sku:
        raise SKUNotFoundError(
            "SKU não encontrado."
        )

    # ---------------------------------------------------------
    # 2. Normalizar dados
    # ---------------------------------------------------------

    codigo = codigo.strip().upper()
    cor = cor.strip().upper()
    tamanho = tamanho.strip().upper()
    publico = validar_publico(publico)

    # ---------------------------------------------------------
    # 3. Verificar produto
    # ---------------------------------------------------------

    produto = buscar_produto_por_id(
        db=db,
        produto_id=produto_id
    )

    if not produto:
        raise SKUNotFoundError(
            "Produto não encontrado."
        )

    if not produto.ativo:
        raise SKUValidationError(
            "Não é possível associar SKU a um produto inativo."
        )

    # ---------------------------------------------------------
    # 4. Validar promoção
    # ---------------------------------------------------------

    validar_promocao(
        em_promocao=em_promocao,
        preco_promocional=preco_promocional,
        preco_venda=preco_venda,
    )

    # ---------------------------------------------------------
    # 5. Verificar código
    # ---------------------------------------------------------

    sku_com_mesmo_codigo = buscar_por_codigo(
        db=db,
        codigo=codigo
    )

    if (
        sku_com_mesmo_codigo
        and sku_com_mesmo_codigo.id != sku_id
    ):
        raise SKUConflictError(
            "Já existe outro SKU com este código."
        )

    # ---------------------------------------------------------
    # 6. Verificar duplicidade da variação
    # ---------------------------------------------------------

    variacao_existente = buscar_por_variacao(
        db=db,
        produto_id=produto_id,
        cor=cor,
        tamanho=tamanho,
        publico=publico,
        ignorar_sku_id=sku_id,
    )

    if variacao_existente:
        raise SKUConflictError(
            "Já existe outro SKU com esta combinação "
            "de produto, cor, tamanho e público."
        )

    # ---------------------------------------------------------
    # 7. Atualizar
    # ---------------------------------------------------------

    sku.produto_id = produto_id
    sku.codigo = codigo
    sku.cor = cor
    sku.tamanho = tamanho
    sku.publico = publico
    sku.quantidade = quantidade
    sku.preco_custo = preco_custo
    sku.preco_venda = preco_venda
    sku.em_promocao = em_promocao
    sku.preco_promocional = preco_promocional

    return atualizar_sku(
        db=db,
        sku=sku
    )