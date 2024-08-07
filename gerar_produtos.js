const xlsx = require('xlsx');
const moment = require('moment');
const fs = require('fs');

// Lista de nomes de produtos (roupas)
const nomesRoupas = [
  'Camisa Polo',
  'Calça Jeans',
  'Vestido Floral',
  'Jaqueta de Couro',
  'Saia Midi',
  'Blusa de Frio',
  'Shorts de Verão',
  'Camisa Social',
  'Cardigan',
  'Tênis Esportivo',
  'Bermuda',
  'Suéter de Lã',
  'Camisa de Seda',
  'Casaco Impermeável',
  'Vestido de Festa',
  'Blusa de Algodão',
];

// Função para gerar um array de produtos com datas de chegada e quantidades aleatórias
function gerarProdutos(numProdutos) {
  const produtos = [];
  const startDate = moment('2024-01-01 08:00:00');
  for (let i = 0; i < numProdutos; i++) {
    const dataChegada = startDate.clone().add(4 * i, 'hours').format('YYYY-MM-DD HH:mm:ss');
    const quantidade = Math.floor(Math.random() * 20) + 1; // Quantidade aleatória entre 1 e 20
    const nomeProduto = nomesRoupas[Math.floor(Math.random() * nomesRoupas.length)]; // Nome aleatório da lista
    produtos.push({
      Nome: nomeProduto,
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
