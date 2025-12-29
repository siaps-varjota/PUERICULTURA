function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'success', message: 'Web app está ativo. Use requisições POST para atualizar a planilha.' })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    // Parseia o corpo da requisição
    const data = JSON.parse(e.postData.contents);
    const row = parseInt(data.row);
    const col = parseInt(data.col);
    const value = data.value;

    // Abre a planilha pelo ID
    const spreadsheetId = '1YhOzK9c4m8Dhgl83APq_-xFc8K4SSY2QtZTThX7x_-w'; // ID da planilha
    const sheetName = 'Sheet1'; // Substitua pelo nome da aba, se necessário
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

    // Atualiza a célula especificada
    const columnLetter = String.fromCharCode(64 + col); // Exemplo: 4 -> D
    const cell = `${columnLetter}${row}`;
    sheet.getRange(cell).setValue(value);

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', message: 'Planilha atualizada' })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
