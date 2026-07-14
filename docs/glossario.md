# Glossário da VJFIT Platform

## Produto

Representa um modelo comercial.

Exemplo:
Camisa Nike Dry Fit.

Não possui estoque próprio.

---

## SKU

Representa uma variação específica do Produto.

Possui:

- Cor
- Tamanho
- Quantidade
- Preço

É a unidade efetivamente vendida.

---

## Pedido

Representa a intenção de compra do cliente.

Ainda pode ser cancelado.

---

## Venda

Representa um pedido confirmado.

Após sua conclusão ocorre a baixa do estoque.

---

## Catálogo

Interface utilizada pelos clientes para visualizar os produtos disponíveis.

Não altera informações.

Apenas consulta a API.

---

## Desktop

Sistema administrativo utilizado pelos funcionários.

Responsável por cadastrar produtos e concluir vendas.

---

## API

Única responsável pelas regras de negócio e pela alteração dos dados persistentes.

---

## Estoque

Quantidade disponível de um SKU.