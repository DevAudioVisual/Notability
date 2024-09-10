        function runAll() {
            var cs = new CSInterface();
            var form1 = document.getElementById("form1").value;
            cs.evalScript('var sheetURL = "' + form1 + '";');
            cs.evalScript('$.runScript.fetchNotas()');
        }   

      function markerNaTL () {
            var cs = new CSInterface();
            var qualAula = document.getElementById("myDropdown").value;
            cs.evalScript('var qualAula = "' + qualAula + '";')
            cs.evalScript('$.runScript.markerNaTL()');
        }

       function markerNaTrack () {
            var cs = new CSInterface();
            var qualAula = document.getElementById("myDropdown").value;
            cs.evalScript('var qualAula = "' + qualAula + '";')
            cs.evalScript('$.runScript.markerNaTrack()');
        }

        function handleFileSelect(event) {
    var file = event.target.files[0];
    if (file) {
        var reader = new FileReader();
        reader.onload = function(e) {
            var content = e.target.result;
            var dropdown = document.getElementById('myDropdown');
            // Remove todas as opções do dropdown
            dropdown.innerHTML = '';

            var jsonData = JSON.parse(content);
            var uniqueValues = [];
            jsonData.forEach(function(item) {
                if (item.Column1 && uniqueValues.indexOf(item.Column1) === -1) {
                    uniqueValues.push(item.Column1);
                }
            });

            // Adiciona as opções no dropdown
            uniqueValues.forEach(function(value, index) {
                var option = document.createElement('option');
                option.textContent = value;
                option.value = index; // Valor do índice
                dropdown.appendChild(option);
            });
        };
        reader.readAsText(file);
    }
}

function init() {
    document.getElementById('fileInputNotes').addEventListener('change', handleFileSelect);
}

document.addEventListener('DOMContentLoaded', init);



