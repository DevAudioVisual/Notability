var selectedFilePath = "";

function escapeFilePath(path) {
    return path.replace(/\\/g, '\\\\');
}

function runAll() {
    var cs = new CSInterface();
    var form1 = document.getElementById("form1").value;
    cs.evalScript('var sheetURL = "' + form1 + '";');
    cs.evalScript('$.runScript.fetchNotas()');
}

function selectFile() {
    var cs = new CSInterface();
    var result = window.cep.fs.showOpenDialog(false, false, 'Select a file', null, null);
    if (result.err === window.cep.fs.NO_ERROR) {
        var filePath = result.data[0];
        selectedFilePath = filePath;
        var readResult = window.cep.fs.readFile(filePath);
        if (readResult.err === 0) {
            var content = readResult.data;
            var dropdown = document.getElementById('myDropdown');
            dropdown.innerHTML = '';

            var jsonData = JSON.parse(content);
            var uniqueValues = [];
            jsonData.forEach(function(item) {
                if (item.Column1 && uniqueValues.indexOf(item.Column1) === -1) {
                    uniqueValues.push(item.Column1);
                }
            });

            uniqueValues.forEach(function(value) {
                var option = document.createElement('option');
                option.textContent = value;
                option.value = value; // Use the actual value
                dropdown.appendChild(option);
            });
        }
    }
}

function markerNaTL() {
    var cs = new CSInterface();
    var qualAula = document.getElementById("myDropdown").value;
    if (!qualAula) {
        alert('Please select an Aula.');
        return;
    }
    cs.evalScript('var qualAula = "' + qualAula + '"; var selectedFilePath = "' + escapeFilePath(selectedFilePath) + '"; $.runScript.markerNaTL();');
}

function markerNaTrack() {
    var cs = new CSInterface();
    var qualAula = document.getElementById("myDropdown").value;
    if (!qualAula) {
        alert('Please select an Aula.');
        return;
    }
    cs.evalScript('var qualAula = "' + qualAula + '"; var selectedFilePath = "' + escapeFilePath(selectedFilePath) + '"; $.runScript.markerNaTrack();');
}

function init() {
    document.getElementById('selectFileButton').addEventListener('click', selectFile);
}

document.addEventListener('DOMContentLoaded', init);
