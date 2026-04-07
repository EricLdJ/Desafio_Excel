# Desafio Prático de Excel: Limpeza, Análise e Busca Reversa

Organizar e analisar o desempenho de colaboradores e lojas. | Demonstrar a evolução de estudos autodidatas em Excel.

## Descrição

Esse post tem por objetivo evidenciar e comprovar os avanços de meus estudos autodidatas com Excel. Nele, utilizei desde fórmulas básicas a métodos de tratamento de base de dados. O desafio proposto foi feito pelo professor João Paulo Araujo, que leciona o curso “Excel Excelente”, e foi coletado no vídeo “Resolvendo na Prática uma Prova de Excel para Entrevista de Emprego! Planilha p/ Baixar”. 

Antes do vídeo ser visto por mim, eu resolvi o desafio da planilha e logo após confirmei meus resultados que estavam corretos. Logo de cara, percebi que apenas um passo meu estava “errado”: utilizei da funcionalidade texto para colunas onde era mais fácil usar a tecla de atalho CTRL + E ou Preenchimento Relâmpago. Mas isso se torna um aprendizado para futuros trabalhos. Desde já, quero deixar meus agradecimentos ao professor pelo ótimo conteúdo que ele disponibiliza de forma gratuita em seu canal no Youtube!

## Preparação e Tratamento dos Dados (Planilha1)

Meu primeiro passo foi abrir a aba “Planilha1”, ver as informações existentes e formatá-las de forma correta para resolver as questões propostas. As colunas Data de admissão, Salário e Cidade - Atuação estavam mal formatadas: Data de admissão e Salário estavam com formato Geral ao invés de, respectivamente, Data abreviada e Moeda. Feitas essas mudanças, por mais que não tenham sido solicitadas nas questões, é importante manter um passo a passo para continuar com minha linha de raciocínio. 

Por fim, notei que a coluna Cidade – Atuação precisava ser separada para o funcionamento da fórmula “MÉDIASES”, onde a questão pedia a média das lojas (LOJA 001, LOJA 002, LOJA 003 e LOJA 004). Para resolver o problema, selecionei a coluna, fui até a aba de dados e usei a funcionalidade “Texto para Colunas”. Escolhi a opção “Delimitado”, marquei a check box de “Outros”, coloquei “-“ e concluí. As colunas foram devidamente separadas e renomeadas para “LOJA” e “CIDADE”, mantendo em maiúsculo para respeitar a formatação original.

Posteriormente nesse README, você verá que essa formatação pela funcionalidade “Texto para Colunas” não resolveu todos os problemas. Na questão 3, tive barreiras para solucionar o erro de “DIV/0” que estava acontecendo mesmo com a fórmula correta. Com um olhar mais minucioso, reparei que quando formatei havia um espaço fantasma aonde estava o “-“. Logo, o Excel não estava conseguindo encontrar o critério solicitado (exemplo: o critério era “LOJA 001”, mas na Planilha1 estava como “LOJA 001 ”). Esse pequeno espaço me tomou 10 minutos, um erro que não acontecerá novamente. Para resolver esse problema, usei a função “ARRUMAR” selecionando a coluna LOJA. Os valores foram colocados nas células e só precisei copiar e colar como valores por cima da coluna original, resolvendo assim o misterioso problema do “DIV/0”.

## Resolução (Aba Desafios)

Logo após tratar os dados da Planilha1, fui analisar a aba Desafios, resolvendo de maneira rápida cada uma:

* **Questão 1:** Foi usada a fórmula `SOMA`, selecionando a coluna de salário inteira.

Soma: <img width="439" height="72" alt="Q1" src="https://github.com/user-attachments/assets/7453a9e4-928a-40b4-8f45-ae5393b0a32c" />

* **Questão 2:** Utilize a função `SOMASES`, viajando entre as informações da Planilha1 e Desafios. Selecionei do banco de dados a coluna funcionários e salário (trancando as células por precaução, mesmo percebendo em segunda análise que não era estritamente necessário) e utilizei o critério "Auxiliar" da aba Desafios. Após fazer o primeiro, com um duplo clique no canto inferior direito da célula, todos os outros critérios foram preenchidos com agilidade.

Somases: <img width="528" height="244" alt="Q2" src="https://github.com/user-attachments/assets/d3ef13b0-cff0-479b-aaab-431c1b613089" />

* **Questão 3 (O Espaço Fantasma):** Aqui utilizei a fórmula `MÉDIASES`. Contudo, me deparei com um erro de `DIV/0`. Com um olhar mais minucioso, reparei que a formatação "Texto para Colunas" deixou um espaço fantasma onde estava o "-" (ex: o critério exigia "LOJA 001", mas na base estava "LOJA 001 "). Esse detalhe me tomou 10 minutos. Resolvi o problema aplicando a função `ARRUMAR` na coluna LOJA, copiando e colando os valores por cima dos originais. Com isso corrigido, foi só repetir os passos da questão anterior para preencher as células.

Médiases: <img width="537" height="167" alt="Q3" src="https://github.com/user-attachments/assets/a1fbefe3-9e41-41ab-9582-5ea45d24e119" />

* **Questão 4 (Busca Reversa):** Utilizei as funções `MÁXIMO` e `MÍNIMO` selecionando a coluna de salários. O diferencial foi achar os nomes correspondentes a esses salários: fiz uma pesquisa reversa utilizando a fórmula `PROCX`. O `PROCV` tem o problema de só conseguir ler informações da esquerda para a direita, ou seja, se a matriz de retorno está à esquerda do valor procurado, ele acusa erro. O `PROCX` não tem essa barreira. Preenchendo a fórmula com os valores de máximo/mínimo como pesquisa, e as colunas de salários e nomes como matrizes, o resultado apareceu perfeitamente.

Máximo: <img width="784" height="86" alt="Q4_maior" src="https://github.com/user-attachments/assets/f3f27dd4-a8bb-426b-9028-cee9d0f1a8bc" />

Mínimo: <img width="782" height="79" alt="Q4_menor" src="https://github.com/user-attachments/assets/7e167eeb-7c7a-4e43-8db0-3ce498c21392" />

### Busca com PROCX:

<img width="782" height="79" alt="Q4_menor" src="https://github.com/user-attachments/assets/f504cbc8-cf93-4b8e-a3e4-e4b306ef4284" />

Prova do funcionamento da fórmula PROCX para a pesquisa reversa de nomes.

Com isso, todas as questões foram resolvidas em 40 minutos!

## Primeiros passos

### Dependências 
* Microsoft Excel (Office 365 recomendado para suporte ao PROCX)
* Git

### Instalação
Não é necessário instalar bibliotecas externas. Basta baixar a planilha e abri-la no seu Excel.

### Execução
Instale e abra o bash do Git na pasta que desejar, execute o comando git clone. 

`git clone https://github.com/EricLdJ/Desafio_Excel.git`

Ou, se desejar, baixe o arquivo no GitHub e abra o arquivo no seu Excel. Navegue entre as abas `Planilha1` e `Desafios` para visualizar os cálculos.

## Prints

### Aba da base de dados

Antes: <img width="1341" height="636" alt="Visao_geral_Planilha1" src="https://github.com/user-attachments/assets/7d19eb41-aeb6-4f89-8021-78c154cb9e4f" />

Depois: <img width="1338" height="635" alt="Visao_geral_Planilha1_resolvido" src="https://github.com/user-attachments/assets/45cfafab-7b39-40f5-9c94-a3c642888123" />

### Aba de Desafios

<img width="1254" height="554" alt="Visao_geral_DESAFIO" src="https://github.com/user-attachments/assets/9ed1430b-f603-47a7-9401-491ae22466b3" />

Essa é a visão geral do desafio sem estar concluído.

<img width="1292" height="581" alt="Visao_geral_DESAFIO_resolvido" src="https://github.com/user-attachments/assets/fd632817-d369-4ec2-bcfb-49ce705f69e7" />

Essa é a visão geral do desafio concluído, contendo as respostas automatizadas.


## Referências
Vídeo: https://www.youtube.com/watch?v=-ETCutmxH7M

## Autor
Eric Lima de Jesus

email: ericlj333@gmail.com

LinkedIn: https://www.linkedin.com/in/eric-l-jesus/

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data_Analysis-000000?style=for-the-badge&logo=google-analytics&logoColor=white)
![Problem Solving](https://img.shields.io/badge/Problem_Solving-0052CC?style=for-the-badge&logo=code-climate&logoColor=white)
