function formatDate(date) {
  return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")
}

function parseDate(date) {
  let [year, month, day] = date.split('-')
  return new Date(year, month - 1, day, 0, 0, 0, 0);
}

function addDays(date, days) {
  date.setDate(date.getDate() + days)
  return date
}

function maxDate(...args) {
  return formatDate(new Date(Math.max(...args.map(date=>parseDate(date)))));
}
function minDate(...args) {
  return formatDate(new Date(Math.min(...args.map(date=>parseDate(date)))));
}

const _MS_PER_DAY = 1000 * 60 * 60 * 24;
function dateDifferenceInDays(a, b) {
  // Discard the time and time-zone information.
  const utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

  return Math.floor((utc2 - utc1) / _MS_PER_DAY);
}
