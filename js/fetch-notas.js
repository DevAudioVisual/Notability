const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// URL da planilha (troque com o link da sua planilha)
const sheetID = process.argv[2];
const sheetURL = 'https://docs.google.com/spreadsheets/d/' + sheetID + '/export?format=xlsx';

// Caminho temporário para salvar o arquivo XLSX baixado
const tempFilePath = path.join(__dirname, 'temp_sheet.xlsx');

// Função para baixar o arquivo XLSX
async function downloadSheet(url, destination) {
    const fetch = (await import('node-fetch')).default;
    const response = await fetch(url);
    const fileStream = fs.createWriteStream(destination);
    await new Promise((resolve, reject) => {
        response.body.pipe(fileStream);
        response.body.on('error', reject);
        fileStream.on('finish', resolve);
    });
}

// Função para processar o arquivo XLSX
function processSheet(filePath) {
    // Lê o arquivo XLSX
    const workbook = XLSX.readFile(filePath);

    // Seleciona a primeira aba da planilha
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Define o intervalo inicial
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const desiredData = [];

    // Expressão regular para buscar valores no formato número.número
    const regex = /^\d+\.\d+$/; // Expressão regular para "número ponto número"
    
    // Encontra a linha com um valor no formato número.número na coluna A
    let startRow = null;
    for (let R = 0; R <= range.e.r; ++R) { // Itera sobre as linhas da planilha
        const cellAddress = { c: 0, r: R }; // Coluna A (c: 0)
        const cell = worksheet[XLSX.utils.encode_cell(cellAddress)];
        
        // Verifica se a célula existe e se o valor corresponde à expressão regular
        if (cell && regex.test(cell.v.toString())) { 
            startRow = R; // Define a linha inicial
            break; // Sai do loop assim que encontrar
        }
    }

    // Se não encontrar nenhum valor no formato número.número, avisa e para a execução
    if (startRow === null) {
        console.error('Erro: Nenhum valor no formato número.número foi encontrado na coluna A.');
        return;
    }

    // Itera sobre as linhas e extrai dados começando da linha encontrada
    for (let R = startRow; R <= range.e.r; ++R) { 
        const row = {};
        for (let C = 0; C <= 4; ++C) { // Colunas A (0), B (1), C (2), D (3)
            const cellAddress = { c: C, r: R };
            const cell = worksheet[XLSX.utils.encode_cell(cellAddress)];
            row[`Column${C+1}`] = cell ? cell.v : ''; // Adiciona o valor da célula ou uma string vazia com chave "Column1", "Column2", etc.
        }
        desiredData.push(row);
    }

    // Salva os dados extraídos em um arquivo JSON
    const jsonData = JSON.stringify(desiredData, null, 2);
    fs.writeFileSync(path.join(__dirname, 'extracted_data.json'), jsonData);

    console.log('Dados extraídos e salvos como JSON.');
}

// Função principal para executar os passos
async function main() {
    try {
        await downloadSheet(sheetURL, tempFilePath);
        console.log('Planilha baixada com sucesso.');
        
        processSheet(tempFilePath);

        // Remove o arquivo temporário
        fs.unlinkSync(tempFilePath);
        console.log('Arquivo temporário deletado.');
    } catch (error) {
        console.error('Erro:', error);
    }
}

main();
