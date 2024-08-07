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