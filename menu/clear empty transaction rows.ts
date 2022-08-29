function clearEmptyTransactionRows() {
  documentLock(_clearEmptyTransactionRows);
}
function _clearEmptyTransactionRows() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const { trx_Dates } = getReferenceRanges(spreadsheet);

  const txnSheet = trx_Dates.getSheet();

  const firstDateRowNumber = trx_Dates.getRow();

  const dateRows = trx_Dates.getValues().map(([cell]) => cell);

  const emptyRows = dateRows.map((v) => v == "");

  const lastNonEmptyRowIdx = emptyRows.lastIndexOf(false);

  if (lastNonEmptyRowIdx !== -1) {
    // delete every empty row before this one
    for (let i = lastNonEmptyRowIdx; i >= 0; --i) {
      if (!emptyRows[i]) continue;
      console.log("Deleting row", i + firstDateRowNumber);
      txnSheet.deleteRow(i + firstDateRowNumber);
    }
  } else {
    console.log("Nothing to delete");
  }
  console.log("Done");
}
