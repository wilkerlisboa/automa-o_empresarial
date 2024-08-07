<p align="center">
	<img src="https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white"/>
	<img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black" />
	<img src="https://img.shields.io/badge/Node.js-43853D?style=for-the-badge&logo=node.js&logoColor=white"/> 
</p>

<h1 align="center">Automação Empresarial: Desenvolvimento de uma Solução para Calcular a Média de Chegada de Produtos Utilizando Node.js, JavaScript e Excel</h1>

Este projeto tem como objetivo calcular e analisar a média de chegada de produtos com base em dados armazenados em um arquivo Excel. O código gera um relatório detalhado com médias por ano, mês, semana, dia e hora. Além disso, a análise é exportada para um novo arquivo Excel.

<h2> Resumo </h2>
O projeto é dividido em duas partes principais:

1.  **Geração do Arquivo Excel com Dados de Produtos**
2.  **Análise e Cálculo das Médias dos Dados**

Caso já tenha o arquivo Excel com os detalhes do fluxo de chegada, pode descartar a etapa "1. Geração do Arquivo Excel com Dados de Produtos". Essa opção é para caso não tenha um fluxo de chegadas de produtos, podendo gerar um fluxo fictício.

<h2>1. Geração do Arquivo Excel com Dados de Produtos</h2>

### Código JavaScript
O código a seguir gera um arquivo Excel contendo dados sobre produtos, incluindo nome, data de chegada e quantidade. O arquivo gerado é chamado `chegada_produtos.xlsx`.

Importação de Bibliotecas javascript

	const xlsx = require('xlsx');
	const moment = require('moment');
	const fs = require('fs');
		
- `xlsx:` Esta biblioteca é usada para criar e manipular arquivos Excel.
- `moment:` Esta biblioteca é usada para manipulação e formatação de datas.
- `fs:` Esta biblioteca do Node.js é usada para manipulação de arquivos no sistema de arquivos.
  
Função para Gerar Produtos

	function gerarProdutos(numProdutos) {
	  const produtos = [];
	  const startDate = moment('2024-01-01 08:00:00');
	  for (let i = 0; i < numProdutos; i++) {
	    const dataChegada = startDate.clone().add(4 * i, 'hours').format('YYYY-MM-DD HH:mm:ss');
	    const quantidade = Math.floor(Math.random() * 20) + 1; // Quantidade aleatória entre 1 e 20
	    produtos.push({
	      Nome: `Produto ${i + 1}`,
	      DataChegada: dataChegada,
	      Quantidade: quantidade,
	    });
	  }
	  return produtos;
	}

- `gerarProdutos(numProdutos):` Esta função gera um array de objetos, onde cada objeto representa um produto.
- `produtos:` Inicializa um array vazio que armazenará os produtos.
- `startDate:` Define a data inicial de chegada dos produtos como 1º de janeiro de 2024 às 08:00.
- `Loop for:` Itera de 0 até numProdutos para criar cada produto
- `dataChegada:` Calcula a data de chegada para cada produto, adicionando 4 horas para cada produto subsequente.
- `quantidade:` Gera uma quantidade aleatória entre 1 e 20 para cada produto.
- `produtos.push(...):` Adiciona um objeto ao array produtos com o nome, data de chegada e quantidade.
  
Função para Gerar o Arquivo Excel

	function gerarExcel(produtos, filePath) {
	  const worksheet = xlsx.utils.json_to_sheet(produtos);
	  const workbook = xlsx.utils.book_new();
	  xlsx.utils.book_append_sheet(workbook, worksheet, 'ChegadaProdutos');
	  xlsx.writeFile(workbook, filePath);
	}

- `gerarExcel(produtos, filePath):` Esta função gera um arquivo Excel a partir do array de produtos.
- `worksheet:` Converte o array de produtos em uma planilha.
- `workbook:` Cria um novo livro de trabalho (workbook) do Excel.
- `xlsx.utils.book_append_sheet(workbook, worksheet, 'ChegadaProdutos'):` Adiciona a planilha criada ao livro de trabalho com o nome 'ChegadaProdutos'.
- `xlsx.writeFile(workbook, filePath):` Escreve o livro de trabalho no caminho especificado (filePath).

Geração dos Produtos e Criação do Arquivo Excel

	const produtos = gerarProdutos(100);
	const filePath = './chegada_produtos.xlsx';
	gerarExcel(produtos, filePath);
	
	console.log(`Arquivo Excel gerado em: ${filePath}`);

- `const produtos = gerarProdutos(100):` Gera um array de 100 produtos usando a função gerarProdutos.
- `const filePath = './chegada_produtos.xlsx':` Define o caminho onde o arquivo Excel será salvo.
- `gerarExcel(produtos, filePath):` Chama a função gerarExcel para criar o arquivo Excel com os produtos gerados e salvar no caminho especificado.
- `console.log(...):` Exibe uma mensagem no console indicando que o arquivo Excel foi gerado e o local onde ele foi salvo.

O nome do arquivo, `chegada_produtos.xlsx`, é importante, pois ele será utilizado em outro momento na etapa "2. Análise e Cálculo das Médias dos Dados".
### Execução
Para rodar o código e gerar um fluxo fictício, presumo que você tenha instalado os arquivos e dependências necessários, como o VS Code (Visual Studio Code) e o Node.js (ambiente de execução JavaScript gratuito).

Caso ainda não tenha instalado, seguem os links abaixo com todas as orientações necessárias para realizar as instalações:

		https://nodejs.org/

		https://code.visualstudio.com/
	
Após realizar todas as instalações, abra o diretório onde você irá executar o código. Neste caso, foi criado um diretório chamado `produtos`. Dentro desse diretório, foi criado o arquivo `gerar_produtos.js`. É nesse arquivo que você deve colocar o código anterior.



![image](https://github.com/user-attachments/assets/a4fbc2ed-1386-429e-83c7-e1dd4d8e1609)



Para instalar as dependências necessárias, como  `xlsx` e ` moment`, execute o comando abaixo:


		npm install xlsx moment

Após isso, abra o terminal (seja no Windows ou no Linux) e execute o comando:

		gerar_produtos.js

![image](https://github.com/user-attachments/assets/d1e007ca-d7f9-4433-a569-479239517861)


Assim, seu diretório padrão conterá um arquivo JavaScript e um arquivo Excel chamado `chegada_produtos.xlsx`. Agora, finalizamos a primeira etapa e podemos prosseguir para a segunda etapa deste trabalho.

<h2>2. Análise e Cálculo das Médias dos Dados</h2>

### Código JavaScript
O código abaixo lê o arquivo `chegada_produtos.xlsx`, calcula as médias de chegada dos produtos e salva essas médias em um novo arquivo Excel chamado `media_chegada_produtos.xlsx`.
	
Importação de Bibliotecas

 	const xlsx = require('xlsx');
	const moment = require('moment');
	const fs = require('fs');

- `xlsx:` Biblioteca para criar e manipular arquivos Excel.
- `moment:` Biblioteca para manipulação de datas e horas.
- `fs:` Biblioteca do Node.js para operações com o sistema de arquivos.

Definição de Caminhos dos Arquivos

	const filePath = './chegada_produtos.xlsx';
	const outputFilePath = './media_chegada_produtos.xlsx';
 
- `filePath:` Caminho para o arquivo Excel de entrada, que contém os dados de chegada dos produtos.
- `outputFilePath:` Caminho para o arquivo Excel de saída, que conterá as médias calculadas.

Estruturas para Armazenar Dados de Chegada

	const chegadaPorAno = {};
	const chegadaPorMes = {};
	const chegadaPorSemana = {};
	const chegadaPorDia = {};
	const chegadaPorHora = {};

- Esses objetos armazenarão a quantidade total e o número de registros por diferentes intervalos de tempo (ano, mês, semana, dia e hora).

Função para Calcular a Média de Chegada dos Produtos

	function calcularMediaChegada() {
	  const workbook = xlsx.readFile(filePath);
	  const sheetName = workbook.SheetNames[0];
	  const sheet = workbook.Sheets[sheetName];
	  const data = xlsx.utils.sheet_to_json(sheet);
	
	  data.forEach(row => {
	    const dataHora = moment(row.DataChegada, 'YYYY-MM-DD HH:mm:ss');
	    const quantidade = parseInt(row.Quantidade, 10);
	    const nomeProduto = row.Nome;

- `workbook = xlsx.readFile(filePath):` Lê o arquivo Excel.
- `sheetName = workbook.SheetNames[0]:` Obtém o nome da primeira planilha.
- `sheet = workbook.Sheets[sheetName]:` Obtém a primeira planilha.
- `data = xlsx.utils.sheet_to_json(sheet):` Converte a planilha em um array de objetos JSON.
- `data.forEach(...):` Itera sobre cada linha da planilha.

Atualização dos Objetos de Armazenamento por Intervalo de Tempo
Para cada intervalo de tempo (ano, mês, semana, dia e hora), os seguintes passos são realizados:

	// Para média por ano
	const ano = dataHora.format('YYYY');
	if (!chegadaPorAno[ano]) {
	  chegadaPorAno[ano] = { total: 0, count: 0, produtos: {} };
	}
	if (!chegadaPorAno[ano].produtos[nomeProduto]) {
	  chegadaPorAno[ano].produtos[nomeProduto] = { total: 0, count: 0 };
	}
	chegadaPorAno[ano].produtos[nomeProduto].total += quantidade;
	chegadaPorAno[ano].produtos[nomeProduto].count += 1;
	chegadaPorAno[ano].total += quantidade;
	chegadaPorAno[ano].count += 1;
 
- `const ano = dataHora.format('YYYY'):` Extrai o ano da data de chegada.
- Verifica se o ano existe no objeto `chegadaPorAno`. Se não, inicializa com valores padrão.
- Verifica se o produto existe no ano. Se não, inicializa com valores padrão.
- Atualiza os totais e contagens para o produto e o ano.

O mesmo processo é repetido para mês, semana, dia e hora, usando os formatos apropriados de data e hora ('YYYY-MM', 'YYYY-[W]WW', 'YYYY-MM-DD', 'YYYY-MM-DD HH:mm').

	const calcularMedia = (obj, label) => {
	  const resultados = [];
	  console.log(`Média de chegada por ${label}: \n`);
	  for (const key in obj) {
	    const media = obj[key].total / obj[key].count;
	    console.log(`${key}: Média total: ${media.toFixed(2)} unidades \n`);
	    resultados.push({ Período: `${key} (Média Total)`, Média: `${media.toFixed(2)} unidades` });
	    console.log('Detalhes por produto: \n');
	    for (const produto in obj[key].produtos) {
	      const produtoMedia = obj[key].produtos[produto].total / obj[key].produtos[produto].count;
	      console.log(`  ${produto}: ${produtoMedia.toFixed(2)} unidades`);
	      resultados.push({ Período: `${key} (Produto: ${produto})`, Média: `${produtoMedia.toFixed(2)} unidades` });
	    }
	    console.log('');
	  }
	  return resultados;
	};

- `calcularMedia:` Função que calcula e imprime a média de chegada dos produtos por um intervalo de tempo especificado.
- `obj:` O objeto contendo os dados de chegada para o intervalo de tempo.
- `label:` O rótulo do intervalo de tempo (ano, mês, semana, dia ou hora).

Calcular as Médias e Criar o Arquivo Excel

	const mediasAno = calcularMedia(chegadaPorAno, 'ano');
	const mediasMes = calcularMedia(chegadaPorMes, 'mês');
	const mediasSemana = calcularMedia(chegadaPorSemana, 'semana');
	const mediasDia = calcularMedia(chegadaPorDia, 'dia');
	const mediasHora = calcularMedia(chegadaPorHora, 'hora');
	
	const wsAno = xlsx.utils.json_to_sheet(mediasAno, { header: ['Período', 'Média'] });
	const wsMes = xlsx.utils.json_to_sheet(mediasMes, { header: ['Período', 'Média'] });
	const wsSemana = xlsx.utils.json_to_sheet(mediasSemana, { header: ['Período', 'Média'] });
	const wsDia = xlsx.utils.json_to_sheet(mediasDia, { header: ['Período', 'Média'] });
	const wsHora = xlsx.utils.json_to_sheet(mediasHora, { header: ['Período', 'Média'] });
	
	const wb = xlsx.utils.book_new();
	xlsx.utils.book_append_sheet(wb, wsAno, 'Média por Ano');
	xlsx.utils.book_append_sheet(wb, wsMes, 'Média por Mês');
	xlsx.utils.book_append_sheet(wb, wsSemana, 'Média por Semana');
	xlsx.utils.book_append_sheet(wb, wsDia, 'Média por Dia');
	xlsx.utils.book_append_sheet(wb, wsHora, 'Média por Hora');
	
	xlsx.writeFile(wb, outputFilePath);
	
	console.log(`Arquivo Excel gerado com as médias em: ${outputFilePath}`);

- Calcula as médias para cada intervalo de tempo usando a função `calcularMedia.`
- Converte os resultados em planilhas Excel usando `xlsx.utils.json_to_sheet.`
- Cria um novo livro de trabalho Excel.
- Adiciona cada planilha ao livro de trabalho.
-Escreve o livro de trabalho no arquivo de saída `(outputFilePath).`

Executar a Função para Calcular a Média de Chegada dos Produtos

	calcularMediaChegada();
 
- Chama a função `calcularMediaChegada` para executar todo o processo descrito acima.
  
### Descrição do Código

1.  **Leitura do Arquivo Excel**: Lê os dados do arquivo `chegada_produtos.xlsx`.
2.  **Estrutura de Armazenamento**: Cria objetos para armazenar a quantidade total e o número de registros por ano, mês, semana, dia e hora.
3.  **Processamento dos Dados**: Itera pelos dados do Excel e atualiza os objetos de armazenamento com a quantidade de produtos e o número de registros.
4.  **Cálculo das Médias**: Calcula a média de chegada dos produtos para cada período.
5.  **Geração do Arquivo Excel com Médias**: Cria um novo arquivo Excel chamado `media_chegada_produtos.xlsx` contendo as médias calculadas.

### Exemplo de Dados de Médias no Excel
![image](https://github.com/user-attachments/assets/c5fcab9f-f1b2-488e-9bde-5ea0817931b8)


### Execução
Para iniciar o projeto, digite o comando abaixo para gerar o relatório do fluxo de chegada de produtos:

	node calcular_medias.js

 **Verifique se o arquivo foi gerado**:
    
`media_chegada_produtos.xlsx`: Contém as médias calculadas para cada período (ano, mês, semana, dia e hora).

<h2>Conclusão</h2>

Este projeto demonstrou como utilizar Node.js, JavaScript e Excel para automatizar a análise de dados empresariais. Com a geração de um fluxo de dados fictício ou real, seguido pela análise detalhada desses dados, é possível obter médias de chegada de produtos em diferentes períodos (ano, mês, semana, dia e hora). O resultado é um relatório compreensivo, exportado para um arquivo Excel, que facilita a tomada de decisões informadas sobre a logística e a gestão de estoque.

O uso de bibliotecas como xlsx para manipulação de arquivos Excel e moment para gerenciamento de datas mostra a flexibilidade e a potência do ecossistema JavaScript para resolver problemas do mundo real. Esta solução automatizada não só economiza tempo, mas também garante a precisão e a consistência na análise dos dados.

Por fim, este projeto serve como um exemplo prático de como integrar diferentes tecnologias para criar uma ferramenta robusta de análise de dados, contribuindo significativamente para a eficiência operacional das empresas.

<h2>Autor 🤝</h2>


<table>
  <tr>
    <td align="center">
      <a href="#" title="defina o título do link">
        <img src="https://avatars.githubusercontent.com/u/73085812" width="100px;" alt="Foto do wilker lisboa no  github"/><br>
        <sub>
          <b>Wilker Lisboa</b>
        </sub>
      </a>
    </td>
  </tr>
</table>

