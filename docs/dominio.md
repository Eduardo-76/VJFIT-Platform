Vamos lá: temos aqui a empresa física, temos o estoque e algumas peças para provador. Mas agora, neste notebook, ficará o sistema principal. Hoje em dia, as pessoas só querem comprar online. Haverá pessoas que virão até a loja física, mas será a minoria. Por isso, é preciso estar sempre de prontidão no WhatsApp.

O sistema já tem todo o estoque disponível com os preços cadastrados. Então, quando alguém escolher uma ou mais peças de roupa pelo catálogo online, basta você copiar e colar nesse campo. O sistema fará a leitura dos SKUs para você e montará o carrinho do cliente. Em seguida, você volta com a mensagem confirmando o pedido e finaliza a venda com ele. Depois, o estoque já baixa automaticamente e, no catálogo, se a peça estiver esgotada, ela já sinaliza para os demais clientes que está fora de estoque.

Importante: sempre averigue o estoque e veja quais peças podem ser encomendadas. Talvez seja interessante adicionar um botão para solicitar todas as peças com baixa no estoque. O sistema em si já permite ver o que está baixo, mas poder imprimir tudo de uma vez seria útil para ter um relatório melhor.

# Quando o cliente compra...
O que nasce?
Um Pedido?
Uma Venda?
Os dois?

Acredito que nasce os dois, eles vem um seguido do outro primeiro o pedido depois a venda.

# Quando uma venda é concluída...

Ela possui o quê?
Itens?
Valor?
Data?
Cliente?
Funcionário?

Ela possui valor

# O catálogo conhece o estoque?
Ou ele apenas consulta?

Ela apenas consulta, ela não faz nenhum tipo de alteração.

# O Desktop é dono do estoque?
Ou a API?

O PostgreSQL guarda os dados. A API aplica as regras de negócio e é a única autorizada a alterar o estoque. O Desktop deixa de alterar diretamente e passa a solicitar essas operações à API.

# O que realmente diminui?
Produto?
SKU?
Quantidade?

Quantidade disponível do SKU.

# O WhatsApp vende?
Ou apenas comunica?

É apenas o meio de comunicação quem vende é o negocio. 

# O Carrinho existe antes da venda?
Ou ele já é uma venda?

O carrinho hoje ele existe antes da venda, pois a sua compra não sera finalizada lá.


| Conceito | O que é?               | Responsabilidade                                 |
| -------- | ---------------------- | ------------------------------------------------ |
| Produto  | Modelo comercial       | Agrupar variações                                |
| SKU      | Unidade vendável       | Controlar estoque e preço                        |
| Carrinho | Seleção temporária     | Reunir itens antes do pedido                     |
| Pedido   | Intenção de compra     | Aguardar confirmação                             |
| Venda    | Pedido concluído       | Registrar a operação e disparar baixa de estoque |
| Catálogo | Vitrine online         | Exibir produtos e consultar disponibilidade      |
| API      | Camada de negócio      | Aplicar regras e alterar dados                   |
| Desktop  | Sistema administrativo | Operar o negócio diariamente                     |
