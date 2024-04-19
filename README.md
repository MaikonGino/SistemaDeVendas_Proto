## SISTEMA DE VENDAS - CAIXA SIMPLES

Este é um sistema desenvolvido em Excel para ajudar no controle de vendas, métodos de pagamento e valores devidos por cliente.
Ele é útil para pequenos negócios que desejam acompanhar suas vendas de forma organizada e eficiente.

### Funcionalidades:

- **Registro de Vendas:** Registra vendas feitas em dinheiro, cartão, Pix ou fiado.
- **Controle de Métodos de Pagamento:** Permite selecionar o método de pagamento para cada venda.
- **Cálculo Automático:** Calcula automaticamente o total das vendas em dinheiro, cartão e Pix.
- **Gestão de Fiados:** Registra os valores devidos por cliente para vendas fiadas.
- **Relatórios Detalhados:** Gera relatórios detalhados dos valores devidos por cliente, facilitando o acompanhamento e a cobrança.

### Funcionalidades de cada pasta de trabalho:
- **Pasta: Registro de Vendas:**

    Nesta pasta, o funcionário registra a quantidade de cada produto vendido, com produtos e valores totalmente personalizáveis.
    Ao final dos produtos, há uma célula para o valor total da venda (em vermelho). Além disso, há uma célula para o valor que o cliente deu em dinheiro, calculando automaticamente o troco.

Sub FinalizarVenda()  
    ' Macro para encerrar a venda  
    ' Realiza o cálculo do total da venda  
    ' Exibe uma caixa de diálogo para selecionar o método de pagamento: Dinheiro, Cartão, Pix ou Fiado  
    ' Com base na seleção do usuário, atribui o valor total da venda ao método de pagamento correspondente  
    ' Registra os valores na planilha "CAIXA"  
    ' Registra os valores na planilha "FIADO" (caso o método de pagamento seja Fiado)  
    ' Limpa o valor da venda na planilha de vendas  
    ' Exibe uma mensagem de finalização da venda  
End Sub

- **Pasta: Caixa:**

    **Instruções de Uso:**
    Nesta pasta, os valores dos diferentes métodos de pagamento são registrados automaticamente pelo sistema de vendas.
    Os valores são recebidos do macro "FinalizarVenda" e são organizados em colunas de acordo com o método de pagamento: dinheiro, cartão e Pix.

Sub SomarMetodosPagamento()
    ' Macro para fechar o caixa  
    ' Inicializa as variáveis necessárias para calcular os totais de cada método de pagamento  
    ' Encontra a última linha com dados na planilha  
    ' Calcula o total em dinheiro, cartão e Pix  
    ' Exibe os totais em uma mensagem de informação  
End Sub  

- **Nota:**

    A pasta "Caixa" serve como um resumo dos valores recebidos em diferentes métodos de pagamento ao longo do período de vendas.
    O botão "Fechar Caixa" permite gerar um relatório rápido desses valores, facilitando o controle financeiro e a prestação de contas.

- **Pasta: Fiado**

    Nesta pasta, são registrados os valores das vendas a prazo, ou seja, vendas para as quais o pagamento será feito em uma data futura.
    Os valores são recebidos automaticamente pelo sistema de vendas e organizados em duas colunas: uma para o nome do cliente e outra para o valor devido.

Sub CalcularValoresDevidos()
    ' Macro para calcular os valores devidos por cliente  
    ' Inicializa as variáveis necessárias para calcular os valores devidos por cliente  
    ' Percorre as linhas com dados na planilha de fiado  
    ' Extrai o nome do cliente e o valor devido de cada linha  
    ' Verifica se o nome do cliente já existe na lista de valores devidos. Se sim, adiciona o valor devido ao total existente. Se não, adiciona o cliente à lista  
    ' Cria um relatório com os valores devidos por cliente  
    ' Exibe o relatório em uma mensagem e pergunta ao usuário se deseja imprimir  
    ' Se o usuário optar por imprimir o relatório, cria uma nova planilha com os dados e configura a área de impressão  
End Sub  

- **Nota:**

    A pasta "Fiado" permite acompanhar os valores devidos por clientes e gerar relatórios para controle financeiro.
    Ao clicar no botão "Relatório", o usuário pode visualizar os valores devidos e, se desejar, imprimir um relatório para referência futura.
  
### Uso do Sistema:

- **Abertura do Arquivo:** Abra o arquivo "SISTEMA DE VENDAS - CAIXA SIMPLES.xlsx" no Excel.
- **Registro de Vendas:** Na planilha "VENDAS_FESTAS", registre as vendas nas células apropriadas.
- **Seleção do Método de Pagamento:** Selecione o método de pagamento para cada venda.
- **Finalização da Venda:** Clique no botão "Finalizar Venda" para calcular os totais e registrar a venda.
- **Visualização dos Valores Devidos:** Para visualizar os valores devidos por cliente, clique no botão "Calcular Valores Devidos".

- **Notas Adicionais:**

    - **Atenção ao VBA e Macros:** Este sistema faz uso extensivo de macros e códigos VBA (Visual Basic for Applications).
    Certifique-se de que os códigos estejam escritos corretamente e as macros atribuídas adequadamente para o correto funcionamento do sistema.
    - **Visualização em Tela Cheia:** A visualização em tela cheia e em 130% na tela de vendas é um diferencial, proporcionando uma melhor experiência de uso.
    - **Sugestões e Melhorias:** Se você tiver alguma sugestão de melhoria ou encontrar algum problema, por favor, entre em contato. Seu feedback é muito bem-vindo!

### Nota de Compatibilidade:

    Este sistema foi desenvolvido no Microsoft Office Excel 2016 e pode apresentar compatibilidade variada com outras versões do Office.
    Recomenda-se testar em versões mais antigas ou mais recentes do Office para garantir o funcionamento adequado de todas as funcionalidades.

