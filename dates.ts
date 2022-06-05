export function formatDate(date: Date) {
  return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd")
}

export function parseDate(date: string) {
  let [year, month, day] = date.split('-')
  return new Date(+year, +month - 1, +day, 0, 0, 0, 0);
}

export function addDays(date: Date, days: number) {
  date.setDate(date.getDate() + days)
  return date
}

export function maxDate(...args: string[]) {
  return formatDate(new Date(Math.max(...args.map(date=>+parseDate(date)))));
}
export function minDate(...args: string[]) {
  return formatDate(new Date(Math.min(...args.map(date=>+parseDate(date)))));
}

const _MS_PER_DAY = 1000 * 60 * 60 * 24;
export function dateDifferenceInDays(a: Date, b: Date) {
  // Discard the time and time-zone information.
  const utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

  return Math.floor((utc2 - utc1) / _MS_PER_DAY);
}
