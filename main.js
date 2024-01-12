const fs = require('fs');
const xlsx = require('xlsx');
const xml2js = require('xml2js');
const path = require('path');

function findFiles(directory) {
  if (!directory) directory = './app/xml';
  else directory = `${directory}`;

  const files = fs.readdirSync(directory);
  const filenames = [];

  for (const file of files) {
    const basename = path.basename(file);
    if (/\w+\.xml/.test(basename)) filenames.push(path.basename(file));
  }
  if (!filenames.length) {
    console.log('The folder is empty or the files are not in the correct format!');
    console.log('Supported file format is "xslx"!');
    return;
  };
  return filenames;
}

// Função para converter um arquivo XML em um objeto JavaScript
function parseXML(xmlString) {
  return new Promise((resolve, reject) => {
    const parser = new xml2js.Parser({ explicitArray: false });
    parser.parseString(xmlString, (err, result) => {
      if (err) {
        reject(err);
      } else {
        resolve(result);
      }
    });
  });
}

const enums = {
  CNPJ: 'cnpj',
  xNome: 'nome',
  enderDest: 'endereco',
  xLgr: 'logradouro',
  nro: 'numero',
  xBairro: 'bairro',
  cMun: 'ibge',
  xMun: 'municipio',
  UF: 'uf',
  CEP: 'cep',
  cPais: 'codigoPais',
  xPais: 'pais',
  fone: 'telefone',
  IE: 'ie',
};

function renameKeysRecursively(obj, keyMap) {
  const renamedObj = {};

  for (const [key, value] of Object.entries(obj)) {
    if (typeof value === 'object' && value !== null) {
      renamedObj[keyMap[key] || key] = renameKeysRecursively(value, keyMap);
    } else {
      renamedObj[keyMap[key] || key] = value;
    }
  }

  return renamedObj;
}

function convertToXLSX(data, outputPath) {
  const ws = xlsx.utils.json_to_sheet(data);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Sheet 1');
  xlsx.writeFile(wb, outputPath);
}

async function processXMLFiles(xmlFiles, outputPath) {
  const dataArray = [];

  for (const xmlFile of xmlFiles) {
    const xmlFilePath = `app/xml/Odair/${xmlFile}`;
    console.log('read: ', xmlFile);
    const xmlData = fs.readFileSync(xmlFilePath, 'utf-8');
    const jsonData = await parseXML(xmlData);
    const data = renameKeysRecursively(jsonData.sistema.dest, enums);
    const formattedData = {
      cnpj: data.cnpj,
      nome: data.nome,
      ...data.endereco,
      ie: data.ie,
    }
    dataArray.push(formattedData);
  }
  console.log(dataArray[0]);
  convertToXLSX(dataArray, outputPath);
}

// Exemplo de uso
const xmlFiles = findFiles('app/xml/Odair');
const outputPath = 'app/output/output.xlsx';

processXMLFiles(xmlFiles, outputPath);
