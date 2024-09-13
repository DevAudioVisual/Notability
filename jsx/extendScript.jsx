$.runScript = {

  fetchNotas: function() {
      // URL da planilha
      var url = sheetURL;

      // Extraí o ID da planilha
      var sheetID = extractSheetID(url);

      // Exemplo de uso
      createAndRunBatchFile(sheetID, limpaCelulas);

      // Função para extrair o ID da planilha da URL
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

      // Função para criar e executar o arquivo .bat
      function createAndRunBatchFile(sheetID, callback) {
          var batContent = '@echo off\n' +
                           'set "SHEET_ID=' + sheetID + '"\n' +
                           '"C:\\Program Files\\nodejs\\node.exe" "C:\\Program Files (x86)\\Common Files\\Adobe\\CEP\\extensions\\Notability\\js\\fetch-notas.js" %SHEET_ID%\n' +
                           'pause'; // Adiciona pause para manter a janela aberta

          var batFilePath = Folder.temp.fsName + "/runScript.bat";
          var f = new File(batFilePath);
          f.open("w");
          f.writeln(batContent);
          f.close();
          f.execute();
          
          // Chama o callback após um atraso (ajuste o tempo conforme necessário)
          $.sleep(1500); // Aguarda 1,5 segundos
          if (typeof callback === "function") {
              callback();
          }
      }

      // Função para criar e executar o arquivo .bat para limpar células
      function limpaCelulas() {
          var batContent2 = '"C:\\Program Files\\nodejs\\node.exe" "C:\\Program Files (x86)\\Common Files\\Adobe\\CEP\\extensions\\Notability\\js\\limpa-json.js"\n' +
                            'pause'; // Adiciona pause para manter a janela aberta
          var batFilePath2 = Folder.temp.fsName + "/limpaJson.bat";
          
          var x = new File(batFilePath2);
          x.open("w");
          x.writeln(batContent2);
          x.close();
          x.execute();
      }
  },

  markerNaTL: function() {
      #include ./filterClass.jsx
      
      // Agora o arquivo JSON é lido e gravado no Desktop
      readJSONFile(File(Folder.desktop.fsName + "/cleaned_data.json")); // Usando o Desktop para o JSON
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

      // Agora o arquivo JSON é lido e gravado no Desktop
      readJSONFile(File(Folder.desktop.fsName + "/cleaned_data.json")); // Usando o Desktop para o JSON
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
