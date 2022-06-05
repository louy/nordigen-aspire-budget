import { scriptLock } from "../lock";
import { getAccessToken, getReferenceRanges, INSTITUTIONS_SHEET_NAME, nordigenRequest } from "../utils";

const ADD_LOGOS = false;

function loadInstitutions() { scriptLock(_loadInstitutions) }
function _loadInstitutions() {
  const ui = SpreadsheetApp.getUi();
  
  let result = ui.prompt(
      'Enter your country code (default GB):',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {
    return
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

  const {
    cfg_Accounts,
  } = getReferenceRanges(spreadsheet);

  let institutionsSheet = spreadsheet.getSheetByName(INSTITUTIONS_SHEET_NAME)

  if (!institutionsSheet) {
    institutionsSheet = spreadsheet.insertSheet().setName(INSTITUTIONS_SHEET_NAME)
    spreadsheet.setActiveSheet(institutionsSheet);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());
    institutionsSheet.appendRow(['ID', 'Name', 'Max transaction days', 'Logo'])

    // FIXME - hacky?
    spreadsheet.getSheetByName('Balances').getRange('B7:B7')
      .copyFormatToRange(
        institutionsSheet.getSheetId(),
        1, 4,
        1, 1,
      );

    institutionsSheet.setColumnWidth(2, 300);
  }

  spreadsheet.setActiveSheet(institutionsSheet);

  institutionsSheet.getRange(2, 1, institutionsSheet.getMaxRows(), institutionsSheet.getMaxColumns()).clear();
  institutionsSheet.getImages().forEach(image => image.remove());

  const accessToken = getAccessToken();

  const data = nordigenRequest<{
    id: string,
    name: string,
    transaction_total_days: number,
    logo: string,
  }[]>('/api/v2/institutions/?country=' + encodeURIComponent(text || 'gb'), {
    headers: {
      Authorization: "Bearer " + accessToken,
    },
  });

  const LOGO_MAX_WIDTH = 150
  const LOGO_MAX_HEIGHT = 50

  if (ADD_LOGOS) {
    institutionsSheet.setColumnWidth(4, LOGO_MAX_WIDTH);
  }

  for (const row of data) {
    console.log(row);

    institutionsSheet.appendRow([row.id, row.name, row.transaction_total_days])
    
    // FIXME - hacky?
    spreadsheet.getSheetByName('Balances').getRange('E8:E8')
      .copyFormatToRange(
        institutionsSheet.getSheetId(),
        1, 4,
        institutionsSheet.getLastRow(), institutionsSheet.getLastRow(),
      );

    if (ADD_LOGOS) {
      institutionsSheet.setRowHeight(institutionsSheet.getLastRow(), LOGO_MAX_HEIGHT)

      if (row.logo) {
        const logo = institutionsSheet
          .insertImage(row.logo, 4, institutionsSheet.getLastRow())
        fitImage(logo, LOGO_MAX_HEIGHT, LOGO_MAX_WIDTH);
      }
    }
  }
}
