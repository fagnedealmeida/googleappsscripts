function onEdit(e) {


//Aba da planilha de origem dos dados
  let sheetName = "RETENÇÃO";
//Aba da planilha de destino dos dados;
  let targetSheetName = "RA'S COBRADOS"
//Coluna referencia da ação
  let colToCheck = 17 // Coluna G do check
// Desconsidera do cabeçalho
  let headerRow = 2 // Linha de inicio dos dados


  var sheet = e.source.getActiveSheet()
  var range = e.range


// Verifica se a edição foi na ABA correta e na COLUNA correta e na LINHA correta sendo maior do que a referencia
  if (sheet.getName() === sheetName && range.getColumn() === colToCheck && range.getRow() >= headerRow && range.getValue() === true) {


    //Indicando a intervalo da linha para copiar
    let rowToCopy = range.getRow()


    //Armazenando a variavel da ABA que receberá os dados na variavel local do IF.
    let targetSheet = e.source.getSheetByName(targetSheetName)


// Copia os dados para a outra aba e seleciona os intervalos
    //Salvando a ultima linha na variavel da planilha
    let lastRow = targetSheet.getLastRow()
    // Indicando o número de linhas para serem copiadas
    let numRows = 1 //numero de linhas
    //Indicando o número de colunas a serem copiadas
    let numCols = 6 //numero de colunas
    //Indica a linha para inserir os dados
    let startRow = lastRow + 1


    let endRow = lastRow + numRows - 1
    let endCol = numCols


// Variavel para guardar a referência do que vai ser colado e os intervalos.
// Referenciando as variaveis anteriores para o range da planilha destino
    let targetRange = targetSheet.getRange(startRow, 1, numRows, numCols)


//Processo de transferir o intervalo que vai ser copiado entre as planilhas
    sheet.getRange(rowToCopy, 8, 1, numCols).copyTo(targetRange)
    // 1 = intervalo de linhas copiadas apartir da linha referencia
    // 8 = numero da coluna inicial da cópia


// Exclui a linha original
// sheet.deleteRow(rowToCopy)
}}
