export function onInstall() {
  onOpen();
  // Perform additional setup as needed.
}

export function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Initialise', 'initialise')
    .addItem('Load institutions', 'loadInstitutions')
    .addItem('Link an account', 'linkAccount')
    .addItem('Load accounts', 'loadAccounts')
    .addItem('Load transactions', 'loadTransactions')
    
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Utils')
      .addItem('Clear empty transaction rows', 'clearEmptyTransactionRows'))

    .addToUi();
}
