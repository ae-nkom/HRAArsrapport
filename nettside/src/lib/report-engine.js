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
  if (typeof value === "number" && value > 20000 && value < 70000) {
    return new Date(Date.UTC(1899, 11, 30 + value));
  }
  const text = String(value).trim();
  const local = text.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2}|\d{4})$/);
  if (local) {
    const year = Number(local[3].length === 2 ? `20${local[3]}` : local[3]);
    return utcDate(year, Number(local[2]), Number(local[1]));
  }
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  return iso ? utcDate(Number(iso[1]), Number(iso[2]), Number(iso[3])) : null;
}

export function countNorwegianWorkingDays(startValue, endValue, year) {
  const start = parseReportDate(startValue);
  const end = parseReportDate(endValue);
  const numericYear = Number(year);
  if (!start || !end || !Number.isInteger(numericYear) || start > end) return 0;

  const first = new Date(Math.max(start.getTime(), utcDate(numericYear, 1, 1).getTime()));
  const last = new Date(Math.min(end.getTime(), utcDate(numericYear, 12, 31).getTime()));
  if (first > last) return 0;

  const holidays = norwegianPublicHolidayKeys(numericYear);
  let days = 0;
  for (let date = first; date <= last; date = addUtcDays(date, 1)) {
    const weekday = date.getUTCDay();
    if (weekday !== 0 && weekday !== 6 && !holidays.has(dateKey(date))) days += 1;
  }
  return days;
}

export function leaveDaysForRow(row, year) {
  return countNorwegianWorkingDays(row?.Start, row?.Slutt, year);
}

export function selectExactSnapshotFiles(files, year, periodKey) {
  if (!year || !periodKey) return [];
  const targetKey = `${year}-${periodKey}`;
  return (Array.isArray(files) ? files : []).filter((file) => file?.snapshotKey === targetKey);
}

export function validateWorkbookFile(file, maxBytes = 15 * 1024 * 1024) {
  if (!file || !/\.xlsx$/i.test(String(file.name || ""))) {
    throw new Error(`${file?.name || "Filen"} avvises. Last opp en XLSX-fil.`);
  }
  if (!Number.isFinite(file.size) || file.size < 0 || file.size > maxBytes) {
    throw new Error(`${file.name} er større enn grensen på ${maxBytes / 1024 / 1024} MB.`);
  }
  return true;
}

export function consumeSelectedFiles(fileInput) {
  const files = Array.from(fileInput?.files || []);
  if (fileInput && files.length) fileInput.value = "";
  return files;
}

export function buildParentalLeaveEmployeesByYear(rows, year) {
  const byPerson = new Map();

  for (const row of rows) {
    if (!row?.["Kjønn"] || !row?.["Fornavn"] || !row?.["Etternavn"]) continue;
    const days = leaveDaysForRow(row, year);
    if (!days) continue;
    const key = `${row["Etternavn"]}|${row["Fornavn"]}|${row["Kjønn"]}`;
    const current = byPerson.get(key) || {
      name: `${row["Etternavn"]}, ${row["Fornavn"]}`,
      gender: row["Kjønn"],
      days: 0
    };
    current.days += days;
    byPerson.set(key, current);
  }

  return [...byPerson.values()]
    .map((employee) => ({ ...employee, weightedDays: employee.days, weeks: employee.days / 5 }))
    .sort((left, right) => right.days - left.days || left.name.localeCompare(right.name, "nb"));
}

export function summarizeParentalLeaveByYear(rows, year) {
  const employees = buildParentalLeaveEmployeesByYear(rows, year);
  const women = employees.filter((employee) => employee.gender === "Kvinne");
  const men = employees.filter((employee) => employee.gender === "Mann");
  const womenDays = women.reduce((sum, employee) => sum + employee.days, 0);
  const menDays = men.reduce((sum, employee) => sum + employee.days, 0);
  const totalDays = womenDays + menDays;

  return {
    totalEmployees: employees.length,
    womenCount: women.length,
    menCount: men.length,
    womenDays,
    menDays,
    totalDays,
    womenAvgWeeks: women.length ? womenDays / women.length / 5 : 0,
    menAvgWeeks: men.length ? menDays / men.length / 5 : 0,
    womenShareDays: totalDays ? womenDays / totalDays * 100 : 0,
    menShareDays: totalDays ? menDays / totalDays * 100 : 0
  };
}
