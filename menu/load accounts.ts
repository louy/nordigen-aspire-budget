interface Account {
  name: string;
  id: string;
  details: string;
  lastFetched: string;
  lastBalance: number | string;
  lastBalanceDate: string;
  currency: string;
  displayName: string;
  product: string;
  ownerName: string;
  bban: string;
  maskedPan: string;
  institutionId: string;
  status: string;
  message: string;
}
const columns: {
  key: keyof Account;
  label: string;
  format: "text" | "money" | "date";
}[] = [
  { key: "name", label: "Name", format: "text" },
  { key: "id", label: "ID", format: "text" },
  { key: "details", label: "Account Details", format: "text" },
  { key: "status", label: "Status", format: "text" },
  { key: "message", label: "Message", format: "text" },
  { key: "lastFetched", label: "Last fetched", format: "date" },
  { key: "lastBalance", label: "Last balance", format: "money" },
  { key: "lastBalanceDate", label: "Last balance date", format: "date" },
  { key: "currency", label: "Currency", format: "text" },
  { key: "displayName", label: "Display name", format: "text" },
  { key: "product", label: "Product", format: "text" },
  { key: "ownerName", label: "Owner name", format: "text" },
  { key: "bban", label: "BBAN", format: "text" },
  { key: "maskedPan", label: "Masked PAN", format: "text" },
  { key: "institutionId", label: "Instutition ID", format: "text" },
];

function loadAccounts() {
  scriptLock(_loadAccounts);
}
function _loadAccounts() {
  // console.log('loadAccounts')
  const ui = SpreadsheetApp.getUi();

  const accessToken = getAccessToken();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let requisitionsSheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME);

  if (!requisitionsSheet) {
    ui.alert(
      "Please link an account before using this command",
      ui.ButtonSet.OK
    );
    return;
  }

  let accountsSheet = spreadsheet.getSheetByName(ACCOUNTS_SHEET_NAME);

  if (!accountsSheet) {
    let activeSheet = spreadsheet.getActiveSheet();
    accountsSheet = spreadsheet.insertSheet().setName(ACCOUNTS_SHEET_NAME);
    spreadsheet.setActiveSheet(accountsSheet);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());
    accountsSheet.appendRow(columns.map((h) => h.label));
    formatAccountsTable(spreadsheet, accountsSheet);
    spreadsheet.setActiveSheet(activeSheet);
    accountsSheet.hideSheet();
  } else {
    accountsSheet
      .getRange(1, 1, 1, columns.length)
      .setValues([columns.map((h) => h.label)]);
  }

  // console.log({accessToken})

  const requisitions = requisitionsSheet.getSheetValues(
    2,
    1,
    requisitionsSheet.getLastRow() - 1,
    3
  );

  let accounts = [];
  const errors = [];

  for (const [index, [id, status, institutionId]] of Object.entries(
    requisitions
  )) {
    if (!id) continue;
    try {
      const data = nordigenRequest<{
        accounts: string[];
      }>("/api/v2/requisitions/" + encodeURIComponent(id) + "/", {
        headers: {
          Authorization: "Bearer " + accessToken,
        },
      });

      // console.log(data) // @TODO: Update requisition status

      if (data.accounts) {
        accounts.push(
          ...data.accounts.map((accountId: string) => ({
            id: accountId,
            requisitionId: id,
            institutionId,
          }))
        );
      }
    } catch (error) {
      errors.push(error);
    }
  }
  if (errors.length) console.error(errors);

  for (const { id: accountId, institutionId } of accounts) {
    try {
      const { account } = nordigenRequest<any>(
        "/api/v2/accounts/" + encodeURIComponent(accountId) + "/details/",
        {
          headers: {
            Authorization: "Bearer " + accessToken,
          },
        }
      );
      const { balances } = nordigenRequest<{
        balances: {
          balanceAmount: { amount: number };
          referenceDate: string;
          balanceType: string;
        }[];
      }>("/api/v2/accounts/" + encodeURIComponent(accountId) + "/balances/", {
        headers: {
          Authorization: "Bearer " + accessToken,
        },
      });

      let balance = balances.find(
        ({ balanceType }) => balanceType === "interimCleared"
      );
      if (!balance)
        balance = balances.find(
          ({ balanceType }) => balanceType === "interimBooked"
        );
      if (!balance)
        balance = balances.find(
          ({ balanceType }) => balanceType === "interimAvailable"
        );
      if (!balance)
        balance = balances.find(
          ({ balanceType }) => balanceType === "expected"
        );

      console.log({ account, balances, balance });

      updateAccount(accountsSheet, {
        id: accountId,
        details: account.details,
        // lastFetched: endDate,
        lastBalance: balance?.balanceAmount.amount,
        lastBalanceDate: balance?.referenceDate,

        currency: account.currency,
        displayName: account.displayName,
        product: account.product,
        ownerName: account.ownerName,
        bban: account.bban,
        maskedPan: account.maskedPan,
        institutionId,
        status: "OK",
        message: "",
      });
    } catch (error) {
      updateAccount(accountsSheet, {
        id: accountId,
        institutionId,
        status: "ERROR",
        message: error.message,
      });
    }
  }

  formatAccountsTable(spreadsheet, accountsSheet);
}

function updateAccount(
  accountsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  input: Partial<Account>
) {
  const rowNumber = getRowForAccount(accountsSheet, input.id);
  const range = accountsSheet.getRange(rowNumber, 1, 1, columns.length);
  const row = range.getValues()[0];

  columns.forEach(({ key }, index) => {
    const value = input[key];
    if (value !== undefined) {
      row[index] = value;
    }
  });

  range.setValues([row]);
}

function getRowForAccount(
  accountsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  accountId: string
): number {
  const ids = accountsSheet
    .getRange(2, 2, accountsSheet.getLastRow() - 1 || 1, 1)
    .getValues()
    .map(([cell]) => cell)
    .concat([""]);

  let idx = ids.indexOf(accountId);
  if (idx === -1) {
    // account doesnt exist. find empty row
    idx = ids.indexOf("");

    // Create an auto complete
    accountsSheet
      .getRange(idx + 2, 1, 1, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(
            SpreadsheetApp.getActiveSpreadsheet().getRangeByName(
              "TransactionAccounts"
            ),
            true
          )
          .build()
      );
  }
  return idx + 2;
}

function formatAccountsTable(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  accountsSheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const { trx_Dates } = getReferenceRanges(spreadsheet);

  // header
  trx_Dates
    .getSheet()
    .getRange(trx_Dates.getRow() - 1, trx_Dates.getColumn(), 1, 1)
    .copyFormatToRange(accountsSheet.getSheetId(), 1, 1, 1, 1);
  // FIXME - hacky?
  spreadsheet
    .getSheetByName("Balances")
    .getRange("B7:B7")
    .copyFormatToRange(accountsSheet.getSheetId(), 2, columns.length, 1, 1);

  // FIXME - hacky?
  const money = spreadsheet.getSheetByName("Balances").getRange("D8:D8");
  const date = spreadsheet.getSheetByName("Balances").getRange("B8:B8");
  const text = spreadsheet.getSheetByName("Balances").getRange("E8:E8");

  columns.forEach(({ format }, index) => {
    let source = text;
    if (format === "money") source = money;
    if (format === "date") source = date;

    source.copyFormatToRange(
      accountsSheet.getSheetId(),
      index + 1,
      index + 1,
      2,
      Math.max(2, accountsSheet.getLastRow())
    );
  });
}

function getAccounts(accountsSheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const data = accountsSheet
    .getRange(2, 1, accountsSheet.getLastRow() - 1 || 1, columns.length)
    .getValues();

  return data.map<Account>((row) =>
    row.reduce((acc, item, idx) => {
      acc[columns[idx].key] = item;
      return acc;
    }, {} as Partial<Account>)
  );
}
