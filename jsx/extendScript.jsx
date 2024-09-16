// Ensure that the XMP library is available
if (ExternalObject.AdobeXMPScript === undefined) {
    ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
}

$.runScript = {

    fetchNotas: function() {

        var url = sheetURL;
        var sheetID = extractSheetID(url);

        var projectFolderPath = $.runScript.getProjectFolderPath();

        createAndRunBatchFile(sheetID, projectFolderPath);

        function extractSheetID(url) {
            var regex = /\/d\/([a-zA-Z0-9-_]+)\/|\/d\/([a-zA-Z0-9-_]+)(\/|$)/;
            var match = url.match(regex);

            if (match) {
                var sheetID = match[1] || match[2];
                return sheetID;
            } else {
                throw new Error("ID da planilha não encontrado na URL.");
            }
        }

        function createAndRunBatchFile(sheetID, projectFolderPath) {
            var batContent = '@echo off\n' +
                             'set "SHEET_ID=' + sheetID + '"\n' +
                             'set "OUTPUT_DIR=' + projectFolderPath + '"\n' +
                             '"C:\\Program Files\\nodejs\\node.exe" "C:\\Program Files (x86)\\Common Files\\Adobe\\CEP\\extensions\\Notability\\js\\process-sheet.js" %SHEET_ID% "%OUTPUT_DIR%"\n';

            var batFilePath = Folder.temp.fsName + "/runScript.bat";
            var f = new File(batFilePath);
            f.open("w");
            f.writeln(batContent);
            f.close();
            f.execute();

            $.sleep(3000); 
        }
    },

    getProjectFolderPath: function() {
        var projectPath = app.project.path;
        if (!projectPath) {
            return '';
        }
        var projectFile = new File(projectPath);
        var projectFolder = projectFile.parent.fsName;
        return projectFolder;
    },

     markerNaTL: function() {
        var projectFolderPath = $.runScript.getProjectFolderPath();
        if (!projectFolderPath) {
            alert('Nenhum projeto aberto.');
            return;
        }
        var dataFilePath = projectFolderPath + "/cleaned_data.json";
        processMarkersFromFile(dataFilePath, 'timeline');
    },

    markerNaTrack: function() {
        var projectFolderPath = $.runScript.getProjectFolderPath();
        if (!projectFolderPath) {
            alert('Nenhum projeto aberto.');
            return;
        }
        var dataFilePath = projectFolderPath + "/cleaned_data.json";
        processMarkersFromFile(dataFilePath, 'track');
    }

};

// Function to read data from file and process markers
function processMarkersFromFile(dataFilePath, target) {
    var file = new File(dataFilePath);
    if (!file.exists) {
        alert('Arquivo cleaned_data.json não encontrado no diretório do projeto.');
        return;
    }

    file.open("r");    
    var dataContent = file.read();
    file.close();
    var data = JSON.parse(dataContent);

    processMarkers(data, target);
}

// Function to process markers
function processMarkers(data, target) {
    // Ensure 'qualAula' is defined
    if (typeof qualAula === 'undefined' || !qualAula) {
        alert('Aula não definida. Por favor, selecione uma aula.');
        return;
    }

    var aula = [];
    var tempo = [];
    var tipo = [];
    var texto = [];
    var grupos = {};

    // Populate arrays and group data by 'aula'
    for (var i = 0; i < data.length; i++) {
        var item = data[i];
        aula.push(item["Column1"]);
        tempo.push(item["Column2"]);
        tipo.push(item["Column4"]);
        texto.push(item["Column5"]);

        var valor = item["Column1"];
        if (!grupos[valor]) {
            grupos[valor] = [];
        }
        grupos[valor].push(i); // Store index
    }

    // Access the indices for the selected 'qualAula'
    var indices = grupos[qualAula];
    if (!indices) {
        alert('Aula não encontrada: ' + qualAula);
        return;
    }

    var numeroAula = indices.length;
    var sequence = app.project.activeSequence;
    if (!sequence) {
        alert('Nenhuma sequência ativa.');
        return;
    }

    var markers;
    if (target === 'timeline') {
        markers = sequence.markers;
    } else if (target === 'track') {
        var selected = sequence.getSelection();
        if (selected.length === 0) {
            alert('Nenhum item selecionado na timeline.');
            return;
        }
        var projectItem = selected[0].projectItem;
        markers = projectItem.getMarkers();
    } else {
        alert('Destino desconhecido para os marcadores.');
        return;
    }

    for (var t = 0; t < numeroAula; t++) {
        var idx = indices[t]; // Index into the arrays
        var tempoFormated = tempo[idx];

        // Handle time parsing, considering possible formats
        var tempoSec = parseTimeToSeconds(tempoFormated);

        if (isNaN(tempoSec)) {
            alert('Formato de tempo inválido');
            continue;
        }

        var markerTime = tempoSec; // Time in seconds

        var newMark = markers.createMarker(markerTime);
        newMark.name = tipo[idx];
        newMark.comments = texto[idx];

        // Set marker color based on 'tipo'
        switch (tipo[idx]) {
            case "Corte":
                newMark.setColorByIndex(1);
                break;
            case "Destaque":
                newMark.setColorByIndex(2);
                break;
            case "Inserção":
                newMark.setColorByIndex(4);
                break;
            case "Lettering":
                newMark.setColorByIndex(6);
                break;
            case "Tópicos":
                newMark.setColorByIndex(0);
                break;
            default:
                newMark.setColorByIndex(5); // Default color
                break;
        }
    }
}

// Helper function to parse time string to seconds
function parseTimeToSeconds(timeStr) {
    var partes = timeStr.split(':');
    if (partes.length === 2) {
        var minutos = parseInt(partes[0], 10);
        var segundos = parseFloat(partes[1]);
        return minutos * 60 + segundos;
    } else if (partes.length === 3) {
        var horas = parseInt(partes[0], 10);
        var minutos = parseInt(partes[1], 10);
        var segundos = parseFloat(partes[2]);
        return horas * 3600 + minutos * 60 + segundos;
    } else {
        return NaN;
    }
}
