# Toda regra de negócio deve existir em apenas um lugar.

## Baixa de estoque

A criação de um pedido não altera o estoque.

O estoque será alterado somente durante a finalização da venda.

Antes de concluir uma venda, a API deverá validar novamente a
disponibilidade de todos os SKUs presentes no pedido.

Caso qualquer SKU possua estoque insuficiente, a venda não deverá
ser concluída.

# Se houver mais de uma imagem, como transforma uma de 10 imagens da principal 