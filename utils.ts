const CONFIG_SHEET_NAME = 'NordigenData'
const REFRESH_TOKEN_KEY = 'Refresh token'

const INSTITUTIONS_SHEET_NAME = 'NordigenInstitutions'
const REQUISITIONS_SHEET_NAME = 'NordigenRequisitions'
const ACCOUNTS_SHEET_NAME = 'NordigenAccounts'

let config: [string, string][]

function nordigenRequest<T extends {}>(url: string, {headers, ...options}: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions): T {
  const request = UrlFetchApp.fetch("https://ob.nordigen.com" + url, {
    ...options,
    headers: {
      accept: "application/json",
      ...headers,
    },
  });
  const data = JSON.parse(request.getContentText());
  return data
}

function getAccessToken() {
  if (!config) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const configSheet = spreadsheet.getSheetByName('NordigenData')

    if (!configSheet) {
      throw new Error("Nordigen integration has not been initialised. Please initialise first.")
    }

    config = configSheet.getSheetValues(1, 1, configSheet.getLastRow(), 2) as string[][] as typeof config
  }

  const refreshToken = config.find(([key]) => key === REFRESH_TOKEN_KEY)?.[1]
  if (!refreshToken) throw new Error("Missing refresh token from data. Please re-initialise")

  const data = nordigenRequest<{access: string}>('/api/v2/token/refresh/', {
    method: 'post',
    headers: {
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({refresh:refreshToken}),
  });
  // console.log({data})
  const {access} = data
  return access;
}

function getReferenceValues(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  return [
    'v_Today',
    'v_ReportableCategorySymbol',
    'v_NonReportableCategorySymbol',
    'v_DebtAccountSymbol',
    'v_CategoryGroupSymbol',
    'v_ApprovedSymbol',
    'v_PendingSymbol',
    'v_BreakSymbol',
    'v_AccountTransfer',
    'v_BalanceAdjustment',
    'v_StartingBalance'
  ].reduce((acc, name) => {
    Object.defineProperty(acc, name, {
      get() {
        const value = spreadsheet.getRangeByName(name).getValue()
        Object.defineProperty(acc, name, {value})
        return value;
      },
      configurable: true,
    })
    return acc
  }, {} as Record<string, string|number|Date>)
}

function getReferenceRanges(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  return [
    'trx_Dates',
    'trx_Outflows',
    'trx_Inflows',
    'trx_Categories',
    'trx_Accounts',
    'trx_Statuses',
    'trx_Memos',
    'trx_Uuids',
    'ntw_Dates',
    'ntw_Amounts',
    'ntw_Categories',
    'cts_Dates',
    'cts_Amounts',
    'cts_FromCategories',
    'cts_ToCategories',
    'cfg_Accounts',
    'cfg_Cards',
  ].reduce((acc, name) => {
    Object.defineProperty(acc, name, {
      get() {
        const value = spreadsheet.getRangeByName(name)
        Object.defineProperty(acc, name, {value})
        return value;
      },
      configurable: true,
    })
    return acc
  }, {} as Record<string, GoogleAppsScript.Spreadsheet.Range>)
}
