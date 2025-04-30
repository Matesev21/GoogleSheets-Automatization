function onEdit(e) {
  try {
    const sheetName = "Daily To Dos";
    const targetSheetName = "Completed";
    const statusColumn = 2; // Columna B

    const editedSheet = e.range.getSheet();
    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();

    if (editedSheet.getName() !== sheetName || editedCol !== statusColumn) return;

    const value = e.range.getValue();

    if (value === "DONE") {
      const numCols = editedSheet.getLastColumn();
      const rowData = editedSheet.getRange(editedRow, 1, 1, numCols).getValues()[0];

      // Añadir fecha actual
      const currentDate = new Date();
      rowData.splice(rowData.length - 1, 0, currentDate);// se agrega al final de la fila

      const targetSheet = e.source.getSheetByName(targetSheetName);
      if (!targetSheet) {
        SpreadsheetApp.getUi().alert("La hoja 'Completed' no existe.");
        return;
      }

      targetSheet.appendRow(rowData);
      editedSheet.deleteRow(editedRow);
    }
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}
