function initialise() {
  scriptLock(_initialise);
}
function _initialise() {
  console.log("initialise");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (configSheet) {
    let result = ui.alert(
      "Nordigen config already exists. Do you want to override it?",
      ui.ButtonSet.YES_NO
    );
    switch (result) {
      case ui.Button.NO:
      case ui.Button.CLOSE:
        return;
      case ui.Button.YES:
        config = void 0;
        spreadsheet.deleteSheet(configSheet);
        break;
    }
  }

  let result = ui.prompt(
    "Please enter your Nordigen refresh token:",
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    return;
  }

  let activeSheet = spreadsheet.getActiveSheet();
  configSheet = spreadsheet.insertSheet().setName(CONFIG_SHEET_NAME);
  spreadsheet.setActiveSheet(configSheet);
  spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());

  configSheet.appendRow([REFRESH_TOKEN_KEY, text]);

  spreadsheet.setActiveSheet(activeSheet);
  configSheet.hideSheet();
}
