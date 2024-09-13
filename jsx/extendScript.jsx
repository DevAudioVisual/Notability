$.runScript = {

  fetchNotas: function() {

    var url = sheetURL;
    var sheetID = extractSheetID(url);

    createAndRunBatchFile(sheetID);


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

    function createAndRunBatchFile(sheetID) {
        var batContent = '@echo off\n' +
                         'set "SHEET_ID=' + sheetID + '"\n' +
                         '"C:\\Program Files\\nodejs\\node.exe" "C:\\Program Files (x86)\\Common Files\\Adobe\\CEP\\extensions\\Notability\\js\\process-sheet.js" %SHEET_ID%\n';

        var batFilePath = Folder.temp.fsName + "/runScript.bat";
        var f = new File(batFilePath);
        f.open("w");
        f.writeln(batContent);
        f.close();
        f.execute();

        $.sleep(1500); 
    }
  },

  markerNaTL: function() {
      #include ./filterClass.jsx
      

      readJSONFile(File(Folder.desktop.fsName + "/cleaned_data.json"));
      separarIguais(aula);
      separaUpdateLista();
      tempoMarker(); 

      function tempoMarker() {
          var numeroAula = aulaSeparadas[qualAula].length.toString();
          for (var t = 0; t < numeroAula; t++) {
              var tempoFormated = tempo[t];
              var partes = tempoFormated.split(':');
              var minutos = parseInt(partes[0], 10);
              var segundos = parseInt(partes[1], 10);
              var tempoSec = minutos * 60 + segundos;

              var newMark = app.project.activeSequence.markers.createMarker(tempoSec)
              newMark.name = tipo[t]
              newMark.comments = texto[t]

              if (tipo[t] == "Corte") {
                  newMark.setColorByIndex(1,t)
              }
              if (tipo[t] == "Destaque") {
                  newMark.setColorByIndex(2,t)
              }
              if (tipo[t] == "Inserção") {
                  newMark.setColorByIndex(4,t)
              }
              if (tipo[t] == "Lettering") {
                  newMark.setColorByIndex(6,t)
              }
              if (tipo[t] == "Tópicos") {
                  newMark.setColorByIndex(0,t)
              }      
          }
      }
  },

  markerNaTrack: function () {
      #include ./filterClass.jsx


      readJSONFile(File(Folder.desktop.fsName + "/cleaned_data.json")); 
      separarIguais(aula);
      separaUpdateLista();
      tempoMarker(); 

      function tempoMarker() {
          var numeroAula = aulaSeparadas[qualAula].length.toString();
          var selected = app.project.activeSequence.getSelection();
          var projectItem = selected[0].projectItem;

          for (var t = 0; t < numeroAula; t++) {
              var tempoFormated = tempo[t];
              var partes = tempoFormated.split(':');
              var minutos = parseInt(partes[0], 10);
              var segundos = parseInt(partes[1], 10);
              var tempoSec = minutos * 60 + segundos;

              var newMark = projectItem.getMarkers().createMarker(tempoSec)
              newMark.name = tipo[t]
              newMark.comments = texto[t]

              if (tipo[t] == "Corte") {
                  newMark.setColorByIndex(1,t)
              }
              if (tipo[t] == "Destaque") {
                  newMark.setColorByIndex(2,t)
              }
              if (tipo[t] == "Inserção") {
                  newMark.setColorByIndex(4,t)
              }
              if (tipo[t] == "Lettering") {
                  newMark.setColorByIndex(6,t)
              }
              if (tipo[t] == "Tópicos") {
                  newMark.setColorByIndex(0,t)
              }
          }
      }
  }

}; // Fecha $.runScript
