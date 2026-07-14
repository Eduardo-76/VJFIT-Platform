# Entidades
Produto

Representa um modelo de produto disponível para venda.
Não possui estoque próprio.
Agrupa todas as suas variações.
Exemplo: Nike Dry Fit Masculina


# SKU

Representa uma variação específica do produto.
Possui código único.
Possui estoque.
É a unidade efetivamente vendida ao cliente.
Exemplo: Nike Dry Fit Masculina | Azul | M


Produto

↓

1

↓

N

↓

SKU

---

# Categoria - O que uma Categoria precisa saber sobre si mesma?
Categoria Precisa saber seu proprio nome.
Precisa saber se está ativa
Precisa de descrição?

# Produto 
Talvez aja a necessidade do produto conhecer a categoria, para saber de onde se origina, mas pode ser desnecessario a principio.


Categoria
──────────────────────────

id - Indetificar a categoria
nome - Saber o nome de cada categoria (O nome da categoria deve ser único.)
ativa - Saber se essa categoria ainda esta ativa
created_at - Saber quando foi criada
updated_at - Saber quando houve atualização

Produto
──────────────────────────

id - Indetificar a Produto
categoria - Qual categoria pertence
nome - Nome do produto (Talvez possa haver nomes iguais pertencentes a categorias diferentes)
descrição - Falar sobre o produto
ativa - Se o produto está disponivel
created_at - Saber quando foi criada
updated_at - Saber quando houve atualização

SKU
──────────────────────────

id - Indetificar o SKU
produto_id - Qual produto é
codigo - Código de produto
cor - Cor do produto
tamanho - Tamanho do produto 
sexo - Masculina, Feminino e Infantil (Ou unisex)
quantidade - Quantos tem disponivel
preco_custo - Preço que foi comprando em atacado ou para fazer a peça
preco_venda - Preço de venda
foto - Foto do produto 
created_at - Saber quando foi criada