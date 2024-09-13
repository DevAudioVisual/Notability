const fs = require('fs');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');

// URL da planilha (troque com o link da sua planilha)
const sheetID = process.argv[2];
const sheetURL = 'https://docs.google.com/spreadsheets/d/' + sheetID + '/export?format=xlsx';

// Caminho temporário para salvar o arquivo XLSX baixado
const tempFilePath = path.join(os.tmpdir(), 'temp_sheet.xlsx');

// Caminho temporário para salvar o arquivo JSON extraído
const tempJsonPath = path.join(os.tmpdir(), 'extracted_data.json');

// Caminho para salvar o arquivo JSON limpo na área de trabalho
const desktopDir = path.join(os.homedir(), 'Desktop');
const outputFilePath = path.join(desktopDir, 'cleaned_data.json');

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

// Função para processar o arquivo XLSX e salvar como JSON
function processSheet(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const desiredData = [];

    const regex = /^\d+\.\d+$/;
    let startRow = null;

    for (let R = 0; R <= range.e.r; ++R) {
        const cellAddress = { c: 0, r: R };
        const cell = worksheet[XLSX.utils.encode_cell(cellAddress)];
        if (cell && regex.test(cell.v.toString())) {
            startRow = R;
            break;
        }
    }

    if (startRow === null) {
        console.error('Erro: Nenhum valor no formato número.número foi encontrado na coluna A.');
        return;
    }

    for (let R = startRow; R <= range.e.r; ++R) {
        const row = {};
        for (let C = 0; C <= 4; ++C) {
            const cellAddress = { c: C, r: R };
            const cell = worksheet[XLSX.utils.encode_cell(cellAddress)];
            row[`Column${C+1}`] = cell ? cell.v : '';
        }
        desiredData.push(row);
    }

    const jsonData = JSON.stringify(desiredData, null, 2);
    fs.writeFileSync(tempJsonPath, jsonData);
    console.log(`Dados extraídos e salvos como JSON no diretório temporário: ${tempJsonPath}`);
}

// Função para limpar o JSON e preencher os espaços vazios na Column1
function cleanJSON(filePath, outputFilePath) {
    const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));

    let lastNonEmptyColumn1 = '';

    const cleanedData = data
        .filter(row => {
            const isColumnAVazia = row.Column1 === '';
            const isColumnBVazia = row.Column2 === '';
            const isColumnCIgualWhite = row.Column3 === 'White';
            const isColumnDVazia = row.Column4 === '';
            const isColumnEVazia = row.Column5 === '';
            const isEmptyRow = row.Column1 === '' && row.Column2 === '' && row.Column3 === '' && row.Column4 === '' && row.Column5 === '';

            return !(isColumnAVazia && isColumnBVazia && isColumnCIgualWhite && isColumnDVazia && isColumnEVazia) && !isEmptyRow;
        })
        .map(row => {
            if (row.Column1 === '') {
                row.Column1 = lastNonEmptyColumn1;
            } else {
                lastNonEmptyColumn1 = row.Column1;
            }
            return row;
        });

    fs.writeFileSync(outputFilePath, JSON.stringify(cleanedData, null, 2));
    console.log('JSON limpo, completado e salvo como:', outputFilePath);
}

// Função principal para executar os passos
async function main() {
    try {
        await downloadSheet(sheetURL, tempFilePath);
        console.log('Planilha baixada com sucesso.');
        
        processSheet(tempFilePath);

        // Remove o arquivo temporário XLSX
        fs.unlinkSync(tempFilePath);
        console.log('Arquivo temporário XLSX deletado.');

        // Limpa o JSON gerado e salva na área de trabalho
        cleanJSON(tempJsonPath, outputFilePath);

        // Remove o arquivo JSON temporário
        fs.unlinkSync(tempJsonPath);
        console.log('Arquivo JSON temporário deletado.');
        
    } catch (error) {
        console.error('Erro:', error);
    }
}

main();
