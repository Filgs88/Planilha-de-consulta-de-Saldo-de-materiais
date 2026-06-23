# Planilha de consulta de saldos de materiais usando VBA.

## RESUMO:
Este projeto visa a obtenção de forma mais rapida e simples de saldos de materiais dentro de um almoxarifado industrial, trazendo os saldos atualizados do banco de dados oracle.

## OBJETIVO:
Na empresa a qual trabalho me deparei com um problema de tempo de busca de um saldo do material, a tela do sistema ERP (Enterprise Relation Programming) é muito lenta, e como nessa empresa nós possuímos mais de um almoxarifado, era necessario uma forma mais rapida e organizada para a busca do material correto para dar a retirada do material requisitado.

## CRIAÇÃO:
Para isso eu precisei acessar o banco de dados da empresa (com as devidas permissões) para fazer a coleta dos dados, fazer o tratmento necessario e guarda-lo em uma planilha. Usei uma query SQL para isso. 

Primeiramente eu defini o almoxarifado que eu trabalho, peguei todos os materiais que possuiam um historico de saldo, porque não necessariamente todos os materiais do almoxarifado tiveram alguma compra, e peguei apenas colunas que me eram necessarias, para não deixar o dataset muito grande, sendo apenas 7 colunas.

Com a query pronta eu decidi tomar um caminho mais rapido entre o banco de dados e as planilhas. Utilizei o proprio VBA (linguagem de programação do Excel) para fazer a conexão com o banco de dados e rodar a query criada, salvando a consulta dentro da propria planilha criada com o código.

<img width="1012" height="399" alt="23 06 2026 10-09-44" src="https://github.com/user-attachments/assets/e109e372-439f-47a8-bad0-c0d4ce2e7110" />

O lado bom de se utilizar dessa forma é a velocidade de resposta para ter os dados mais atualizados, afinal, todas as vezes que se faz uma busca ele roda a query dentro do banco de dados e traz as informações mais atualizadas.

<img width="1299" height="959" alt="23 06 2026 10-10-52" src="https://github.com/user-attachments/assets/dc29e395-6677-4acb-a511-f37c92db789e" />

Depois que os dados estão obtidos, foi necessario criar uma filtragem, para que se possa buscar apenas as informações que o usuario precisa. Essa é a "Sub filtrar()" dentro da Base "FiltroAvançado.bas"

<img width="832" height="728" alt="23 06 2026 10-17-23" src="https://github.com/user-attachments/assets/0e966128-7874-473a-9a78-b5debd3f9a60" />

O que ele faz é simplesmente pegar os dados que atendam os critérios estabelecidos na segunda aba de planilha criada que é a pesquisa.

<img width="1742" height="413" alt="23 06 2026 10-18-29" src="https://github.com/user-attachments/assets/2c984f88-b383-4496-aeda-03f69d58809d" />

A partir daqui é uma questão do que voce está buscando, aqui vai alguns exemplos: 
- Materiais com o nome disco; 
- Materiais que custam menos de 1 real a unidade; 
- todos os materiais que são da Familia 19; 
- Todos materiais que tem mais de 10 unidade em estoque; 
- Todos os discos que estão com 0 de saldo; 
- e etc...

Eu criei 4 botões de comando na planilha de pesquisa, o primeiro "Filtrar" faz o filtro customizado que vc precisar, o segundo botão "Geral" ele tráz todos materiais que estão no dataset que trouxemos, mas com o campo de descrição customizado, o terceiro "Estoque", traz todos os materiais que possuem saldos excluindo os zerados, e por fim o botão "Limpar" apenas limpa os campos de critérios não sendo necessário selecionar as células e apagar uma por uma.

Para facilitar a vida de quem vai usar, eu inclui uma pequena sub que roda ao abrir o arquivo que mapeia dois atalhos, "Shift + Enter (Botão grande)" para fazer a consulta dos materiais que possuem saldo, e "Shift + Backspace" para fazer a limpeza dos critérios.

<img width="694" height="213" alt="23 06 2026 10-34-46" src="https://github.com/user-attachments/assets/da4305c0-a1b1-42e8-aa69-875432dc02c0" />

Para buscar essas informações no sistema ERP da empresa é muito trabalhoso e envolve ter que abrir varias janelas, além do tempo de resposta que é muito lento devido ao excesso de dados desnecessarios que ele trás. Com essa planilha, o tempo gasto e as possibilidades de análises são muito grandes, além da praticidade que o Excel nos traz.

## COMO EXECUTAR:
Para executar esse projeto vc precisa apenas do Microsoft Excel com as versões mais recentes, e não é necessario alguma referencia do VBA, ele faz tudo sozinho.

Um outro detalhe é para segurança da empresa eu retirei os dados de acesso ao banco de dados, então a Sub de conexão não irá funcionar.

Para utilizar esse projeto basta clonar esse projeto apenas utilize o comando do terminal: 
git clone https://github.com/Filgs88/Planilha-de-consulta-de-Saldo-de-materiais.git
e depois abrir a planilha que está na pasta.

Desde já agradeço e espero que gostem desse projeto simples. Segue meu Linkedin abaixo para contato.

Luan Filgueira Silva
https://www.linkedin.com/in/luanfilgueirasilva/
