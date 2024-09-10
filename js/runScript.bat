@echo off
setlocal enabledelayedexpansion

rem Define a variável SHEET_ID
set "SHEET_ID=%SHEET_ID%"

rem Executa o script Node.js com a sheetID como argumento
node "C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\Notability\js\fetch-notas.js" %SHEET_ID%

