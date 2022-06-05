function scriptLock(fn) {
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(1000)) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Another script is running in the background. Please wait for it to finish then try again')
    return
  }
  try {
    fn()
  } finally {
    lock.releaseLock()
  }
}

function documentLock(fn) {
  const lock = LockService.getDocumentLock()
  if (!lock.tryLock(1000)) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Document is locked. Please wait and try again later')
    return
  }
  try {
    fn()
  } finally {
    lock.releaseLock()
  }
}
