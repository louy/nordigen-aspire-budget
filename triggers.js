function onInstall(e) {
  onOpen(e);
  // Perform additional setup as needed.
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu()
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
