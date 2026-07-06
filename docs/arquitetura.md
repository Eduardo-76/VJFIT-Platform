# Arquitetura da VJFIT Platform

## Objetivo

Construir uma plataforma para gestão de lojas de roupas, mantendo o sistema desktop como centro da operação e adicionando novos módulos de forma gradual.


---

## Princípios

1. Uma única fonte de dados.

2. Cada módulo possui apenas uma responsabilidade.

3. Nenhum módulo acessa responsabilidades de outro.

4. O sistema deve continuar funcionando mesmo durante futuras evoluções.

5. O Desktop continuará sendo o painel administrativo.

---

## Componentes

Desktop

Responsável por:

- Cadastro
- Estoque
- Venda
- Clientes

---

API

Responsável por:

- Comunicação

---

Catálogo

Responsável por:

- Exibir produtos

- Pesquisa

- Carrinho

---

Banco

Responsável por:

- Persistência

---

Infraestrutura

Responsável por:

- Docker

- PostgreSQL

- Backup
