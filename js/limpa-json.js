const fs = require('fs');
const path = require('path');
const os = require('os');

// Caminho para o arquivo JSON original na pasta temporária
const tempDir = os.tmpdir();
const inputFilePath = path.join(os.tmpdir(), 'extracted_data.json');

// Caminho para salvar o arquivo JSON limpo na área de trabalho
const desktopDir = path.join(os.homedir(), 'Desktop');
const outputFilePath = path.join(desktopDir, 'cleaned_data.json');

// Função para limpar o JSON e preencher os espaços vazios na Column1
function cleanJSON(filePath, outputFilePath) {
    // Lê o arquivo JSON
    const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));

    // Variável para armazenar o último valor não vazio de Column1
    let lastNonEmptyColumn1 = '';

    // Filtra e completa os dados com base nas condições
    const cleanedData = data
        .filter(row => {
            // Checa as condições para ignorar linhas indesejadas
            const isColumnAVazia = row.Column1 === '';
            const isColumnBVazia = row.Column2 === '';
            const isColumnCIgualWhite = row.Column3 === 'White';
            const isColumnDVazia = row.Column4 === '';
            const isColumnEVazia = row.Column5 === '';
            const isEmptyRow = row.Column1 === '' && row.Column2 === '' && row.Column3 === '' && row.Column4 === '' && row.Column5 === '';

            // Filtra as linhas que não precisam ser processadas
            return !(isColumnAVazia && isColumnBVazia && isColumnCIgualWhite && isColumnDVazia && isColumnEVazia) && !isEmptyRow;
        })
        .map(row => {
            // Se a `Column1` estiver vazia, preencher com o último valor não vazio
            if (row.Column1 === '') {
                row.Column1 = lastNonEmptyColumn1;
            } else {
                // Se não estiver vazia, atualizar o último valor não vazio
                lastNonEmptyColumn1 = row.Column1;
            }

            return row;
        });

    // Salva o JSON limpo e completo em um novo arquivo na área de trabalho
    fs.writeFileSync(outputFilePath, JSON.stringify(cleanedData, null, 2));

    console.log('JSON limpo, completado e salvo como:', outputFilePath);
}

// Executa a função de limpeza e preenchimento
cleanJSON(inputFilePath, outputFilePath);
