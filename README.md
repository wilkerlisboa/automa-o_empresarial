<p align="center">
  <img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black" />
  <img src="https://img.shields.io/badge/Node.js-43853D?style=for-the-badge&logo=node.js&logoColor=white"/>
</p>

<h1 align="center">Automação Empresarial: Desenvolvimento de uma Solução para Calcular a Média de Chegada de Produtos Utilizando Node.js, JavaScript e Excel</h1>

Este projeto tem como objetivo calcular e analisar a média de chegada de produtos com base em dados armazenados em um arquivo Excel. O código gera um relatório detalhado com médias por ano, mês, semana, dia e hora. Além disso, a análise é exportada para um novo arquivo Excel.

<h2> Resumo </h2>
O projeto é dividido em duas partes principais:

1.  **Geração do Arquivo Excel com Dados de Produtos**
2.  **Análise e Cálculo das Médias dos Dados**

Caso já tenha ao arquivo excel com o detalhes do fluxo de chegada pode descartar a etapa "1 Geração do Arquivo Excel com Dados de Produtos".
Essa opção e para caso não tenha um fluxo de chegadas de produtos, podendo gerar um fluxo fictício.

<h2>1. Geração do Arquivo Excel com Dados de Produtos</h2>

### Código JavaScript
O código a seguir gera um arquivo Excel contendo dados sobre produtos, incluindo nome, data de chegada e quantidade. O arquivo gerado é chamado `chegada_produtos.xlsx`.
	
	const xlsx = require('xlsx');
	const moment = require('moment');
	const fs = require('fs');

	// Função para gerar um array de produtos com datas de chegada e quantidades aleatórias
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

	// Função para gerar o arquivo Excel
	function gerarExcel(produtos, filePath) {
	  const worksheet = xlsx.utils.json_to_sheet(produtos);
	  const workbook = xlsx.utils.book_new();
	  xlsx.utils.book_append_sheet(workbook, worksheet, 'ChegadaProdutos');
	  xlsx.writeFile(workbook, filePath);
	}

	// Gerar 100 produtos e salvar em um arquivo Excel
	const produtos = gerarProdutos(100);
	const filePath = './chegada_produtos.xlsx';
	gerarExcel(produtos, filePath);

	console.log(`Arquivo Excel gerado em: ${filePath}`);
O nome do arquivo `chegada_produtos.xlsx` e importante pois ele sera chamado em outro momento na etapa " 2.  Análise e Cálculo das Médias dos Dados "
### Execução
Colocar para rodar o codigo anterior para gerar um fluxo ficticio, para executar o arquivo presumo que tenha instalado os arquivos e dependências como o VS code ( Visual studio Code ) e o NodeJS ( ambiente de execução JavaScript gratuito).

Caso não tenha instalador segue os links abaixo com todas as orientações que precisa para realizar as instalações

		https://nodejs.org/

		https://code.visualstudio.com/
	
Após realizar todas as instalações abra o diretorio a onde ira realizar a execução do código,  nesse caso foi criado um diretorio produtos dentro desse diretorio foi criado o arquivo `gerar_produtos.js` será a onde vamos colocar o código anterior dentro dele

imagem

Agora instalar as dependências do como a xlsx moment execute o comando a baixo para isso

		npm install xlsx moment

Após isso abra o terminal seja ele widowns ou link e execute o comando

		gerar_produtos.js

image

assim fica seu diretorio padão um arquivo de javascript e um arquivo de excel chamado de `chegada_produtos.xlsx`, agora finalizamos a primeira etapa, podemos ir para segunda etapa desse trabalho.

<h2>2. Análise e Cálculo das Médias dos Dados</h2>

### Código JavaScript
O código abaixo lê o arquivo `chegada_produtos.xlsx`, calcula as médias de chegada dos produtos e salva essas médias em um novo arquivo Excel chamado `media_chegada_produtos.xlsx`.
	
	const xlsx = require('xlsx');
	const moment = require('moment');
	const fs = require('fs');

	// Caminho para o arquivo Excel
	const filePath = './chegada_produtos.xlsx';
	const outputFilePath = './media_chegada_produtos.xlsx';

	// Objetos para armazenar a quantidade total e o número de registros por ano, mês, semana, dia e hora
	const chegadaPorAno = {};
	const chegadaPorMes = {};
	const chegadaPorSemana = {};
	const chegadaPorDia = {};
	const chegadaPorHora = {};

	// Função para calcular a média de chegada dos produtos
	function calcularMediaChegada() {
	  const workbook = xlsx.readFile(filePath);
	  const sheetName = workbook.SheetNames[0];
	  const sheet = workbook.Sheets[sheetName];
	  const data = xlsx.utils.sheet_to_json(sheet);

	  data.forEach(row => {
	    const dataHora = moment(row.DataChegada, 'YYYY-MM-DD HH:mm:ss');
	    const quantidade = parseInt(row.Quantidade, 10);
	    const nomeProduto = row.Nome;

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

	    // Para média por mês
	    const mes = dataHora.format('YYYY-MM');
	    if (!chegadaPorMes[mes]) {
	      chegadaPorMes[mes] = { total: 0, count: 0, produtos: {} };
	    }
	    if (!chegadaPorMes[mes].produtos[nomeProduto]) {
	      chegadaPorMes[mes].produtos[nomeProduto] = { total: 0, count: 0 };
	    }
	    chegadaPorMes[mes].produtos[nomeProduto].total += quantidade;
	    chegadaPorMes[mes].produtos[nomeProduto].count += 1;
	    chegadaPorMes[mes].total += quantidade;
	    chegadaPorMes[mes].count += 1;

	    // Para média por semana
	    const semana = dataHora.format('YYYY-[W]WW');
	    if (!chegadaPorSemana[semana]) {
	      chegadaPorSemana[semana] = { total: 0, count: 0, produtos: {} };
	    }
	    if (!chegadaPorSemana[semana].produtos[nomeProduto]) {
	      chegadaPorSemana[semana].produtos[nomeProduto] = { total: 0, count: 0 };
	    }
	    chegadaPorSemana[semana].produtos[nomeProduto].total += quantidade;
	    chegadaPorSemana[semana].produtos[nomeProduto].count += 1;
	    chegadaPorSemana[semana].total += quantidade;
	    chegadaPorSemana[semana].count += 1;

	    // Para média por dia
	    const dia = dataHora.format('YYYY-MM-DD');
	    if (!chegadaPorDia[dia]) {
	      chegadaPorDia[dia] = { total: 0, count: 0, produtos: {} };
	    }
	    if (!chegadaPorDia[dia].produtos[nomeProduto]) {
	      chegadaPorDia[dia].produtos[nomeProduto] = { total: 0, count: 0 };
	    }
	    chegadaPorDia[dia].produtos[nomeProduto].total += quantidade;
	    chegadaPorDia[dia].produtos[nomeProduto].count += 1;
	    chegadaPorDia[dia].total += quantidade;
	    chegadaPorDia[dia].count += 1;

	    // Para média por hora
	    const hora = dataHora.format('YYYY-MM-DD HH:mm');
	    if (!chegadaPorHora[hora]) {
	      chegadaPorHora[hora] = { total: 0, count: 0, produtos: {} };
	    }
	    if (!chegadaPorHora[hora].produtos[nomeProduto]) {
	      chegadaPorHora[hora].produtos[nomeProduto] = { total: 0, count: 0 };
	    }
	    chegadaPorHora[hora].produtos[nomeProduto].total += quantidade;
	    chegadaPorHora[hora].produtos[nomeProduto].count += 1;
	    chegadaPorHora[hora].total += quantidade;
	    chegadaPorHora[hora].count += 1;
	  });

	  // Função para calcular e imprimir a média
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

	  // Calcular as médias
	  const mediasAno = calcularMedia(chegadaPorAno, 'ano');
	  const mediasMes = calcularMedia(chegadaPorMes, 'mês');
	  const mediasSemana = calcularMedia(chegadaPorSemana, 'semana');
	  const mediasDia = calcularMedia(chegadaPorDia, 'dia');
	  const mediasHora = calcularMedia(chegadaPorHora, 'hora');

	  // Criar e salvar o arquivo Excel
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
	}

	// Executar a função para calcular a média de chegada dos produtos
	calcularMediaChegada();
### Imagem do Código

### Descrição do Código

1.  **Leitura do Arquivo Excel**: Lê os dados do arquivo `chegada_produtos.xlsx`.
2.  **Estrutura de Armazenamento**: Cria objetos para armazenar a quantidade total e o número de registros por ano, mês, semana, dia e hora.
3.  **Processamento dos Dados**: Itera pelos dados do Excel e atualiza os objetos de armazenamento com a quantidade de produtos e o número de registros.
4.  **Cálculo das Médias**: Calcula a média de chegada dos produtos para cada período.
5.  **Geração do Arquivo Excel com Médias**: Cria um novo arquivo Excel chamado `media_chegada_produtos.xlsx` contendo as médias calculadas.

### Exemplo de Dados de Médias no Excel

<h2>Como Usar</h2>

1. **Execute o código para gerar o arquivo com dados de produtos**:

		node gerar_produtos.js
2. **Execute o código para calcular as médias e gerar o arquivo de médias**:

		node calcular_medias.js

3.  **Verifique os arquivos gerados**:
    
    -   `chegada_produtos.xlsx`: Contém os dados de chegada dos produtos.
    -   `media_chegada_produtos.xlsx`: Contém as médias calculadas para cada período (ano, mês, semana, dia e hora).

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

