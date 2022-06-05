function loadAccounts() { scriptLock(_loadAccounts) }
function _loadAccounts() {
  // console.log('loadAccounts')
  const ui = SpreadsheetApp.getUi();

  const accessToken = getAccessToken();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let requisitionsSheet = spreadsheet.getSheetByName(REQUISITIONS_SHEET_NAME)

  if (!requisitionsSheet) {
    ui.alert('Please link an account before using this command', ui.ButtonSet.OK)
    return
  }

  let accountsSheet = spreadsheet.getSheetByName(ACCOUNTS_SHEET_NAME)

  if (!accountsSheet) {
    let activeSheet = spreadsheet.getActiveSheet();
    accountsSheet = spreadsheet.insertSheet().setName(ACCOUNTS_SHEET_NAME)
    spreadsheet.setActiveSheet(accountsSheet);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());
    accountsSheet.appendRow([
      'Name', 'ID', 'Account Details', 'Last fetched', 'Last balance', 'Last balance date',
      'Currency', 'Display name', 'Product', 'Owner name', 'BBAN', 'Masked PAN', 'Instutition ID',
    ])
    formatAccountsTable(spreadsheet, accountsSheet)
    spreadsheet.setActiveSheet(activeSheet);
    accountsSheet.hideSheet();
  }

  // console.log({accessToken})

  const requisitions = requisitionsSheet.getSheetValues(2, 1, requisitionsSheet.getLastRow()-1, 3)

  let accounts = []

  for (const [index, [id, status, institutionId]] of Object.entries(requisitions)) {
    if (!id) continue
    const data = nordigenRequest('/api/v2/requisitions/'+encodeURIComponent(id)+'/', {
      headers: {
        Authorization: "Bearer " + accessToken,
      },
    })
    
    // console.log(data) // @TODO: Update requisition status

    if (data.accounts) {
      accounts.push(
        ...data.accounts.map(accountId => ({id: accountId, requisitionId: id, institutionId}))
      )
    }
  }

  for (const {id: accountId, institutionId} of accounts) {
    const {account} = nordigenRequest('/api/v2/accounts/'+encodeURIComponent(accountId)+'/details/', {
      headers: {
        Authorization: "Bearer " + accessToken,
      },
    })
    const {balances} = nordigenRequest('/api/v2/accounts/'+encodeURIComponent(accountId)+'/balances/', {
      headers: {
        Authorization: "Bearer " + accessToken,
      },
    });

    let balance = balances.find(({balanceType}) => balanceType === 'interimCleared')
    if (!balance) balance = balances.find(({balanceType}) => balanceType === 'interimBooked')
    if (!balance) balance = balances.find(({balanceType}) => balanceType === 'interimAvailable')
    if (!balance) balance = balances.find(({balanceType}) => balanceType === 'expected')

    console.log({account, balances, balance})

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
    });
  }

  formatAccountsTable(spreadsheet, accountsSheet)
}

function updateAccount(accountsSheet, {
  id, details, lastFetched, lastBalance, lastBalanceDate,
  currency, displayName, product, ownerName, bban, maskedPan, institutionId,
}) {
  const rowNumber = getRowForAccount(accountsSheet, id);
  const row = accountsSheet.getRange(rowNumber, 2, 1, 12).getValues()[0];
  
  [id, details, lastFetched, lastBalance, lastBalanceDate, currency, displayName, product, ownerName, bban, maskedPan, institutionId, ]
    .forEach((value, index) => {
      if (value !== undefined) {
        row[index] = value
      }
    });
  
  accountsSheet.getRange(rowNumber, 2, 1, 12).setValues([row]);
}

function getRowForAccount(accountsSheet, accountId) {
  const ids = accountsSheet.getRange(2, 2, accountsSheet.getLastRow()-1 || 1, 1).getValues().map(([cell]) => cell).concat(['']);

  let idx = ids.indexOf(accountId)
  if (idx === -1) {
    // account doesnt exist. find empty row
    idx = ids.indexOf('')

    // Create an auto complete
    accountsSheet.getRange(idx+2, 1, 1, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRangeByName('TransactionAccounts'), true)
        .build()
      );
  }
  return idx + 2;
}

function formatAccountsTable(spreadsheet, accountsSheet) {
  const {
    trx_Dates,
  } = getReferenceRanges(spreadsheet);

  // header
  trx_Dates.getSheet()
    .getRange(
      trx_Dates.getRow() - 1,
      trx_Dates.getColumn(),
      1,
      1,
    )
    .copyFormatToRange(
      accountsSheet.getSheetId(),
      1, 1,
      1, 1,
    )
  // FIXME - hacky?
  spreadsheet.getSheetByName('Balances').getRange('B7:B7')
    .copyFormatToRange(
      accountsSheet.getSheetId(),
      2, 13,
      1, 1,
    )

  // FIXME - hacky?
  const money = spreadsheet.getSheetByName('Balances').getRange('D8:D8')
  const date = spreadsheet.getSheetByName('Balances').getRange('B8:B8')
  const text = spreadsheet.getSheetByName('Balances').getRange('E8:E8')
  text.copyFormatToRange(accountsSheet.getSheetId(), 1, 13, 2, Math.max(2, accountsSheet.getLastRow()))
  date.copyFormatToRange(accountsSheet.getSheetId(), 4, 4, 2, Math.max(2, accountsSheet.getLastRow()))
  money.copyFormatToRange(accountsSheet.getSheetId(), 5, 5, 2, Math.max(2, accountsSheet.getLastRow()))
  date.copyFormatToRange(accountsSheet.getSheetId(), 6, 6, 2, Math.max(2, accountsSheet.getLastRow()))
}

function getAccounts(accountsSheet) {
  const data = accountsSheet.getRange(
    2, 1, 
    accountsSheet.getLastRow()-1 || 1, 13
  ).getValues()
  
  return data.map(([
    name, id, details, lastFetched, lastBalance, lastBalanceDate, currency, displayName, product, ownerName, bban, maskedPan, institutionId,
  ]) => ({
    name, id, details, lastFetched, lastBalance, lastBalanceDate, currency, displayName, product, ownerName, bban, maskedPan, institutionId,
  }))
}
