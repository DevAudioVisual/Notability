function escapeString(str) {
    return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function runAll() {
    var cs = new CSInterface();
    var form1 = document.getElementById("form1").value;
    cs.evalScript('var sheetURL = "' + escapeString(form1) + '"; $.runScript.fetchNotas();', function(response) {
        // After fetching, update the dropdown
        loadDataFromProjectFolder();
    });
}

function loadDataFromProjectFolder() {
    var cs = new CSInterface();
    cs.evalScript('$.runScript.getProjectFolderPath();', function(projectFolderPath) {
        if (projectFolderPath) {
            var cleanedDataPath = projectFolderPath + '/cleaned_data.json';

            // Read the cleaned_data.json file
            var readResult = window.cep.fs.readFile(cleanedDataPath);
            if (readResult.err === 0) {
                var content = readResult.data;
                var jsonData = JSON.parse(content);
                populateDropdown(jsonData);
            } else {
                alert('Erro ao ler o arquivo cleaned_data.json: ' + readResult.err);
            }
        } else {
            alert('Não foi possível obter o caminho da pasta do projeto.');
        }
    });
}

function populateDropdown(jsonData) {
    var dropdown = document.getElementById('myDropdown');
    dropdown.innerHTML = '';

    var uniqueValues = [];
    jsonData.forEach(function(item) {
        if (item.Column1 && uniqueValues.indexOf(item.Column1) === -1) {
            uniqueValues.push(item.Column1);
        }
    });

    uniqueValues.forEach(function(value) {
        var option = document.createElement('option');
        option.textContent = value;
        option.value = value;
        dropdown.appendChild(option);
    });
}

function markerNaTL() {
    var cs = new CSInterface();
    var qualAula = document.getElementById("myDropdown").value;
    if (!qualAula) {
        alert('Por favor, selecione uma Aula.');
        return;
    }
    cs.evalScript('var qualAula = "' + escapeString(qualAula) + '"; $.runScript.markerNaTL();');
}

function markerNaTrack() {
    var cs = new CSInterface();
    var qualAula = document.getElementById("myDropdown").value;
    if (!qualAula) {
        alert('Por favor, selecione uma Aula.');
        return;
    }
    cs.evalScript('var qualAula = "' + escapeString(qualAula) + '"; $.runScript.markerNaTrack();');
}

function init() {
    document.getElementById('loadNotesButton').addEventListener('click', runAll);
    document.getElementById('updateDataButton').addEventListener('click', loadDataFromProjectFolder);
}

document.addEventListener('DOMContentLoaded', init);
