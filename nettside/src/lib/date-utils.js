const DAY_MS = 24 * 60 * 60 * 1000;

function utcDate(year, month, day) {
  return new Date(Date.UTC(year, month - 1, day));
}

function dateKey(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function easterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return utcDate(year, month, day);
}

function addUtcDays(date, days) {
  return new Date(date.getTime() + days * DAY_MS);
}

export function norwegianPublicHolidayKeys(year) {
  const easter = easterSunday(year);
  return new Set([
    dateKey(utcDate(year, 1, 1)),
    dateKey(addUtcDays(easter, -3)),
    dateKey(addUtcDays(easter, -2)),
    dateKey(addUtcDays(easter, 1)),
    dateKey(utcDate(year, 5, 1)),
    dateKey(utcDate(year, 5, 17)),
    dateKey(addUtcDays(easter, 39)),
    dateKey(addUtcDays(easter, 50)),
    dateKey(utcDate(year, 12, 25)),
    dateKey(utcDate(year, 12, 26))
  ]);
}

export function parseReportDate(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return utcDate(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
  }
  if (typeof value === "number" && value >= 1 && value < 2958466) {
    const serial = Math.floor(value);
    if (serial === 60) return null; // Excel's fictitious 29 February 1900.
    return new Date(Date.UTC(1899, 11, (serial < 60 ? 31 : 30) + serial));
  }
  const text = String(value).trim();
  const local = text.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2}|\d{4})$/);
  if (local) {
    const year = Number(local[3].length === 2 ? `20${local[3]}` : local[3]);
    return checkedDate(year, Number(local[2]), Number(local[1]));
  }
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  return iso ? checkedDate(Number(iso[1]), Number(iso[2]), Number(iso[3])) : null;
}

function checkedDate(year, month, day) {
  const date = utcDate(year, month, day);
  return date.getUTCFullYear() === year && date.getUTCMonth() + 1 === month && date.getUTCDate() === day ? date : null;
}

export function norwegianWorkingDayKeys(startValue, endValue, year) {
  const start = parseReportDate(startValue);
  const end = parseReportDate(endValue);
  const numericYear = Number(year);
  if (!start || !end || !Number.isInteger(numericYear) || numericYear < 1900 || numericYear > 2200 || start > end) return [];
  const first = new Date(Math.max(start.getTime(), utcDate(numericYear, 1, 1).getTime()));
  const last = new Date(Math.min(end.getTime(), utcDate(numericYear, 12, 31).getTime()));
  const holidays = norwegianPublicHolidayKeys(numericYear);
  const keys = [];
  for (let date = first; date <= last; date = addUtcDays(date, 1)) {
    if (date.getUTCDay() !== 0 && date.getUTCDay() !== 6 && !holidays.has(dateKey(date))) keys.push(dateKey(date));
  }
  return keys;
}

export function countNorwegianWorkingDays(startValue, endValue, year) {
  return norwegianWorkingDayKeys(startValue, endValue, year).length;
}

export function leaveDaysForRow(row, year) {
  return countNorwegianWorkingDays(row?.Start, row?.Slutt, year);
}
