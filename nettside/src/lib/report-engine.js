import { norwegianWorkingDayKeys, leaveDaysForRow } from './date-utils.js';
export { countNorwegianWorkingDays, norwegianPublicHolidayKeys, leaveDaysForRow, parseReportDate } from './date-utils.js';
import { normalizeText, normalizePersonName, parseNumber, mean, formatCurrency } from './formatting.js';
import { parseReportDate as parseCellDate } from './date-utils.js';
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
  if (!file.size) throw new Error(`${file.name} er tom.`);
  return true;
}

export function consumeSelectedFiles(fileInput) {
  const files = Array.from(fileInput?.files || []);
  if (fileInput && files.length) fileInput.value = "";
  return files;
}

export function buildParentalLeaveEmployeesByYear(rows, year) {
  const byPerson = new Map();
  const idsByName = buildSourceIdentityIndex(rows, (row) => `${row.Etternavn || ""} ${row.Fornavn || ""}`);

  for (const row of rows) {
    if (!row?.["Kjønn"] || !row?.["Fornavn"] || !row?.["Etternavn"]) continue;
    const dayKeys = norwegianWorkingDayKeys(row.Start, row.Slutt, year);
    if (!dayKeys.length) continue;
    const fullName = `${row.Etternavn} ${row.Fornavn}`;
    const key = sourcePersonKey(row, fullName, idsByName);
    const current = byPerson.get(key) || {
      personKey: key,
      employeeId: sourceEmployeeId(row),
      name: `${row["Etternavn"]}, ${row["Fornavn"]}`,
      gender: normalizeGender(row["Kjønn"]),
      dayKeys: new Set(),
      overlappingDays: 0
    };
    for (const key of dayKeys) {
      if (current.dayKeys.has(key)) current.overlappingDays += 1;
      current.dayKeys.add(key);
    }
    byPerson.set(key, current);
  }

  return [...byPerson.values()]
    .map(({ dayKeys, ...employee }) => ({ ...employee, days: dayKeys.size, weightedDays: dayKeys.size, weeks: dayKeys.size / 5 }))
    .sort((left, right) => right.days - left.days || left.name.localeCompare(right.name, "nb"));
}

export function summarizeParentalLeaveByYear(rows, year) {
  const employees = buildParentalLeaveEmployeesByYear(rows, year);
  const women = employees.filter((employee) => employee.gender === "Kvinne");
  const men = employees.filter((employee) => employee.gender === "Mann");
  const womenDays = women.reduce((sum, employee) => sum + employee.days, 0);
  const menDays = men.reduce((sum, employee) => sum + employee.days, 0);
  const totalDays = employees.reduce((sum, employee) => sum + employee.days, 0);

  return {
    totalEmployees: employees.length,
    womenCount: women.length,
    menCount: men.length,
    womenDays,
    menDays,
    totalDays,
    womenAvgWeeks: women.length ? womenDays / women.length / 5 : null,
    menAvgWeeks: men.length ? menDays / men.length / 5 : null,
    womenShareDays: totalDays ? womenDays / totalDays * 100 : null,
    menShareDays: totalDays ? menDays / totalDays * 100 : null
  };
}

export const groupOrder = [
  "Totalt antall ansatte",
  "Direktørgruppen",
  "Seksjonssjefgruppen",
  "Fagsjefgruppen",
  "Seniorpersonale",
  "Øvrige saksbehandlere og andre",
  "Uavklart rapportgruppe"
];

export const unassignedGroup = "Uavklart rapportgruppe";
export const excludedGroup = "Utelatt fra rapporten";
export const reportGroups = groupOrder.slice(1).filter((group) => group !== unassignedGroup);

export function normalizeGender(value) {
  const gender = normalizeText(value);
  if (["kvinne", "kvinner", "k", "female"].includes(gender)) return "Kvinne";
  if (["mann", "menn", "m", "male"].includes(gender)) return "Mann";
  return "Uoppgitt";
}

export function sourceEmployeeId(row) {
  const value = row["Ansattnummer"] ?? row["Ansattnr."] ?? row["Ansattnr"] ?? row["Ansatt-ID"];
  const text = String(value ?? "").trim();
  return /^\d+$/.test(text) ? text.replace(/^0+(?=\d)/, "") : "";
}

// Some SAP layouts use the employee-number column as a row counter. Such an ID
// must not merge different people. Blank IDs may be filled by an unambiguous name.
export function buildSourceIdentityIndex(rows, getName) {
  const idsByName = new Map();
  const namesById = new Map();
  for (const row of rows) {
    const name = normalizePersonName(getName(row));
    const id = sourceEmployeeId(row);
    if (!name || !id) continue;
    if (!idsByName.has(name)) idsByName.set(name, new Set());
    idsByName.get(name).add(id);
    if (!namesById.has(id)) namesById.set(id, new Set());
    namesById.get(id).add(name);
  }
  return { idsByName, namesById };
}

export function sourcePersonKey(row, name, index) {
  const normalized = normalizePersonName(name);
  const ownId = sourceEmployeeId(row);
  const nameIds = index?.idsByName.get(normalized);
  const id = ownId || (nameIds?.size === 1 ? [...nameIds][0] : "");
  if (id && index?.namesById.get(id)?.size === 1) return `id:${id}`;
  return `name:${normalized}`;
}

export function employeeKey(row) {
  const id = sourceEmployeeId(row);
  return id ? `id:${id}` : `name:${normalizePersonName(`${row.Etternavn || ""} ${row.Fornavn || ""}`)}`;
}

export function detectEmploymentDateColumns(rows) {
  const headers = rows.length ? Object.keys(rows[0]) : [];
  const normalizedHeaders = headers.map((header) => ({ header, normalized: normalizeText(header) }));
  const startColumn =
    normalizedHeaders.find(({ normalized }) =>
      ["ansettelsesdato", "tiltredelsesdato", "startdato", "gyldig fra", "ansatt fra", "fom"].some((candidate) =>
        normalized.includes(candidate)
      )
    )?.header ??
    headers.find((header) => header === "VB - 0041") ??
    null;
  const endColumn =
    normalizedHeaders.find(({ normalized }) =>
      ["sluttdato", "gyldig til", "ansatt til", "tom", "stoppdato", "sluttetdato"].some((candidate) =>
        normalized.includes(candidate)
      )
    )?.header ?? null;

  return { startColumn, endColumn };
}

export function isActiveOnDate(record, snapshotDate, dateColumns) {
  if (!snapshotDate) return true;
  const startDate = dateColumns.startColumn ? parseCellDate(record[dateColumns.startColumn]) : null;
  const endDate = dateColumns.endColumn ? parseCellDate(record[dateColumns.endColumn]) : null;
  if (startDate && startDate > snapshotDate) return false;
  if (endDate && endDate < snapshotDate) return false;
  return true;
}

export function parseSnapshotKeyDate(snapshotKey) {
  return parseCellDate(snapshotKey);
}

export function periodDateForYear(year, periodKey) {
  if (!/^20\d{2}$/.test(String(year)) || !["05-01", "12-31"].includes(periodKey)) return null;
  return parseCellDate(`${year}-${periodKey}`);
}

export function yearFromDate(date) {
  return date ? String(date.getUTCFullYear()) : "";
}

export function employeeGroup(record, assignments = {}) {
  const override = assignments[employeeKey(record)] || record.Rapportgruppe;
  if ([...reportGroups, excludedGroup].includes(override)) return override;
  const title = String(record["Stillingsgruppe betegnelse"] || "").trim();
  const [code] = title.split(/\s+/);
  if (code === "9106" || code === "1060") return "Direktørgruppen";
  if (code === "1059") {
    const responsibility = normalizeText(record.Personalansvar);
    if (["ja", "yes", "true", "1"].includes(responsibility)) return "Direktørgruppen";
    if (["nei", "no", "false", "0"].includes(responsibility)) return "Fagsjefgruppen";
    return unassignedGroup;
  }
  if (code === "1211") return "Seksjonssjefgruppen";
  if (["1220", "1538", "1072"].includes(code)) return "Fagsjefgruppen";
  if (["1088", "1181", "1114", "1364"].includes(code)) return "Seniorpersonale";
  if (["1087", "1434", "1408", "1065", "1124", "1184"].includes(code)) {
    return "Øvrige saksbehandlere og andre";
  }
  return unassignedGroup;
}

export function annualSalary(record) {
  const individual = parseNumber(record["107A - Individuell lønn årsbe"]);
  const leader = parseNumber(record["1006-Årslønn lederlønnstab."]);
  return individual > 0 ? individual : leader > 0 ? leader : null;
}

export function compareFastlonnCandidates(left, right, snapshotDate) {
  const leftExact = snapshotDate && left.snapshotDate ? left.snapshotDate.getTime() === snapshotDate.getTime() : false;
  const rightExact = snapshotDate && right.snapshotDate ? right.snapshotDate.getTime() === snapshotDate.getTime() : false;
  if (leftExact !== rightExact) return leftExact ? -1 : 1;

  const leftHasDate = Boolean(left.snapshotDate);
  const rightHasDate = Boolean(right.snapshotDate);
  if (leftHasDate !== rightHasDate) return leftHasDate ? -1 : 1;

  if (snapshotDate && left.snapshotDate && right.snapshotDate) {
    const leftDistance = Math.abs(left.snapshotDate.getTime() - snapshotDate.getTime());
    const rightDistance = Math.abs(right.snapshotDate.getTime() - snapshotDate.getTime());
    if (leftDistance !== rightDistance) return leftDistance - rightDistance;
  }

  if (left.snapshotDate && right.snapshotDate && left.snapshotDate.getTime() !== right.snapshotDate.getTime()) {
    return right.snapshotDate.getTime() - left.snapshotDate.getTime();
  }

  return 0;
}

export function buildEmployeeIndex(rows, snapshotDate, assignments = {}) {
  const chosenByPerson = new Map();
  const dateColumns = detectEmploymentDateColumns(rows);
  const unknownByCode = new Map();
  const duplicatePeople = new Set();
  const candidates = rows
    .map((record) => {
      const aliases = [
        `${record["Etternavn"] || ""} ${record["Fornavn"] || ""}`,
        `${record["Fornavn"] || ""} ${record["Etternavn"] || ""}`,
        record["Ansattnr - navn"]
      ]
        .map((alias) => normalizePersonName(alias))
        .filter(Boolean);
      const personKey = employeeKey(record);
      const partTimePercent = parseNumber(record["Deltids-% 0007"]);
      const employee = {
        personKey,
        employeeId: sourceEmployeeId(record),
        name: `${record["Etternavn"] || ""} ${record["Fornavn"] || ""}`.trim(),
        gender: normalizeGender(record["Nøkkel for kjønn"] || record["Kjønn"]),
        group: employeeGroup(record, assignments),
        requiresGroupChoice: employeeGroup(record) === unassignedGroup,
        position: String(record["Stillingsgruppe betegnelse"] || "").trim(),
        salary: annualSalary(record),
        employmentGroup: String(record["Medarbeidergruppe"] || ""),
        partTimePercent: partTimePercent !== null && partTimePercent >= 0 && partTimePercent <= 100 ? partTimePercent : null
      };
      return {
        employee,
        aliases,
        snapshotDate: parseSnapshotKeyDate(record.__snapshotKey || ""),
        record
      };
    })
    .filter(({ employee, record }) => employee.name && isActiveOnDate(record, snapshotDate, dateColumns));

  for (const candidate of candidates) {
    const current = chosenByPerson.get(candidate.employee.personKey);
    if (current && compareFastlonnCandidates(candidate, current, snapshotDate) === 0 &&
      JSON.stringify(candidate.employee) !== JSON.stringify(current.employee)) duplicatePeople.add(candidate.employee.personKey);
    if (!current || compareFastlonnCandidates(candidate, current, snapshotDate) < 0) {
      chosenByPerson.set(candidate.employee.personKey, candidate);
    }
  }

  const byName = new Map();
  const byId = new Map();
  const ambiguousNames = new Set();
  const allEmployees = [...chosenByPerson.values()].map((candidate) => {
    for (const alias of candidate.aliases) {
      if (duplicatePeople.has(candidate.employee.personKey) || (byName.has(alias) && byName.get(alias)?.personKey !== candidate.employee.personKey)) {
        ambiguousNames.add(alias);
        byName.set(alias, null);
      } else if (!ambiguousNames.has(alias)) byName.set(alias, candidate.employee);
    }
    if (candidate.employee.employeeId && !duplicatePeople.has(candidate.employee.personKey)) byId.set(candidate.employee.employeeId, candidate.employee);
    return candidate.employee;
  });
  const employees = allEmployees.filter((employee) => employee.group !== excludedGroup && !duplicatePeople.has(employee.personKey));
  for (const employee of employees.filter((employee) => employee.group === unassignedGroup)) {
    const code = employee.position.split(/\s+/)[0] || "Mangler kode";
    unknownByCode.set(code, (unknownByCode.get(code) || 0) + 1);
  }

  return {
    employees,
    allEmployees,
    byName,
    byId,
    ambiguousNames,
    duplicatePeople,
    excludedEmployees: allEmployees.filter((employee) => employee.group === excludedGroup),
    unknownPositionCodes: [...unknownByCode.entries()].map(([code, count]) => ({ code, count }))
  };
}

export function matchEmployee(row, employeeIndex, name = row["Etternavn, fornavn"]) {
  if (!employeeIndex) return null;
  const id = sourceEmployeeId(row);
  // Prefer a real employee ID. Fall back only to a unique name; never guess.
  const employee = (id && employeeIndex.byId?.get(id)) || employeeIndex.byName.get(normalizePersonName(name));
  return employee && employee.group !== excludedGroup ? employee : null;
}

export function summarizeEmployment(employees) {
  const women = employees.filter((employee) => employee.gender === "Kvinne");
  const men = employees.filter((employee) => employee.gender === "Mann");
  const isTemporary = (employee) => ["Midlert./engasj. tjm", "Vikarer"].includes(employee.employmentGroup);
  const isPartTime = (employee) => employee.partTimePercent !== null && employee.partTimePercent < 100;

  return {
    totalWomen: women.length,
    totalMen: men.length,
    temporary: {
      women: women.filter(isTemporary).length,
      men: men.filter(isTemporary).length
    },
    partTime: {
      women: women.filter(isPartTime).length,
      men: men.filter(isPartTime).length
    }
  };
}

export function variablePayGroup(group) {
  if (group === "Direktørgruppen" || group === "Seksjonssjefgruppen") {
    return "Ledere med personalansvar";
  }
  return group;
}

export const variablePayOrder = [
  "Mottakere i valgt uttrekk",
  "Ledere med personalansvar",
  "Fagsjefgruppen",
  "Seniorpersonale",
  "Øvrige saksbehandlere og andre",
  unassignedGroup
];

export function averageAmount(records) {
  return records.length ? amountSum(records) / records.length : null;
}

export function amountSum(records) {
  return records.reduce((sum, record) => sum + (Number(record.amount) || 0), 0);
}

export function summarizeVariablePay(rows, employeeIndex) {
  const participants = buildVariablePayParticipants(rows, employeeIndex);
  return variablePayOrder.map((groupName) => {
    const subset = groupName === variablePayOrder[0] ? participants : participants.filter((employee) => variablePayGroup(employee.group) === groupName);
    const women = subset.filter((employee) => employee.gender === "Kvinne");
    const men = subset.filter((employee) => employee.gender === "Mann");
    const womenAvg = averageAmount(women);
    const menAvg = averageAmount(men);
    return {
      group: groupName, n: subset.length, womenAvg, menAvg,
      totalAvg: averageAmount(subset),
      womenPctOfMen: womenAvg !== null && menAvg > 0 ? womenAvg / menAvg * 100 : null
    };
  }).filter((row) => row.group !== unassignedGroup || row.n > 0);
}

export function summarizeFastlonn(rows, snapshotDate, assignments = {}) {
  const employeeIndex = buildEmployeeIndex(rows, snapshotDate, assignments);
  const employees = employeeIndex.employees;

  const genderBalance = groupOrder.map((groupName) => {
    const subset =
      groupName === "Totalt antall ansatte"
        ? employees
        : employees.filter((employee) => employee.group === groupName);
    const women = subset.filter((employee) => employee.gender === "Kvinne").length;
    const men = subset.filter((employee) => employee.gender === "Mann").length;
    const total = subset.length;
    return {
      group: groupName,
      women,
      womenShare: total ? (women / total) * 100 : 0,
      men,
      menShare: total ? (men / total) * 100 : 0,
      total
    };
  });

  const fastlonn = groupOrder.map((groupName) => {
    const subset =
      groupName === "Totalt antall ansatte"
        ? employees
        : employees.filter((employee) => employee.group === groupName);
    const womenSalaries = subset.filter((employee) => employee.gender === "Kvinne").map((employee) => employee.salary);
    const menSalaries = subset.filter((employee) => employee.gender === "Mann").map((employee) => employee.salary);
    const allSalaries = subset.map((employee) => employee.salary);
    const womenAvg = mean(womenSalaries);
    const menAvg = mean(menSalaries);
    return {
      group: groupName,
      n: subset.length,
      womenAvg,
      menAvg,
      totalAvg: mean(allSalaries),
      womenPctOfMen: womenAvg !== null && menAvg > 0 ? (womenAvg / menAvg) * 100 : null
    };
  });

  return { employees, employeeIndex, genderBalance: genderBalance.filter((row) => row.group !== unassignedGroup || row.total > 0), fastlonn: fastlonn.filter((row) => row.group !== unassignedGroup || row.n > 0), unknownPositionCodes: employeeIndex.unknownPositionCodes };
}

export function buildFastlonnRows(files) {
  return files.flatMap((file) =>
    file.rows.map((row) => ({
      ...row,
      __snapshotKey: file.snapshotKey || "",
      __snapshotLabel: file.snapshotLabel || "",
      __fileName: file.fileName,
      __fileYear: file.fileYear || ""
    }))
  );
}

export function fileMatchesYear(file, year) {
  return (file.fileYear || file.snapshotKey?.slice(0, 4)) === year;
}

export function buildRowsForRole(files, role) {
  return files
    .filter((file) => file.role === role && file.source === "rådata")
    .flatMap((file) => file.rows);
}

export function getFastlonnFilesForSelection(files, selectedYear, selectedPeriodKey) {
  const availableFastlonnFiles = files.filter((file) => file.role === "fastlønn" && file.source === "rådata");
  const fastlonnFilesForYear = selectedYear
    ? availableFastlonnFiles.filter((file) => fileMatchesYear(file, selectedYear))
    : availableFastlonnFiles;
  const targetSnapshotDate = periodDateForYear(selectedYear, selectedPeriodKey);

  const fastlonnFilesForPeriod = targetSnapshotDate
    ? selectExactSnapshotFiles(fastlonnFilesForYear, selectedYear, selectedPeriodKey)
    : [];

  return {
    availableFastlonnFiles,
    fastlonnFilesForYear,
    fastlonnFilesForPeriod,
    targetSnapshotDate
  };
}

export function filterVariablePayRowsByYear(rows, year) {
  if (!year) return rows;
  return rows.filter((row) => yearFromDate(parseCellDate(row["Utbetalingsdato"])) === year);
}

export function filterParentalLeaveRowsByYear(rows, year) {
  if (!year) return rows;
  const yearStart = new Date(Date.UTC(Number(year), 0, 1));
  const yearEnd = new Date(Date.UTC(Number(year), 11, 31));

  return rows.filter((row) => {
    const startDate = parseCellDate(row["Start"]);
    const endDate = parseCellDate(row["Slutt"]);

    if (startDate && endDate) {
      return startDate <= yearEnd && endDate >= yearStart;
    }

    if (startDate) {
      return yearFromDate(startDate) === year;
    }

    if (endDate) {
      return yearFromDate(endDate) === year;
    }

    return false;
  });
}

export function buildVariablePayParticipants(rows, employeeIndex) {
  return aggregateVariablePay(rows, employeeIndex).filter((person) => person.employee && person.amount > 0)
    .map(({ employee, ...person }) => ({ ...person, gender: employee.gender, group: employee.group }))
    .sort((left, right) => right.amount - left.amount || left.name.localeCompare(right.name, "nb"));
}

export function aggregateVariablePay(rows, employeeIndex) {
  const totals = new Map();
  for (const { personKey, name, amount, employee } of resolveVariablePayRows(rows, employeeIndex)) {
    const person = totals.get(personKey) || { personKey, name, employee, amount: 0 };
    person.amount += amount;
    totals.set(personKey, person);
  }
  return [...totals.values()];
}

export function resolveVariablePayRows(rows, employeeIndex) {
  const resolved = [];
  const identityIndex = buildSourceIdentityIndex(rows, (row) => row["Etternavn, fornavn"]);
  for (const row of rows) {
    const name = String(row["Etternavn, fornavn"] || "").trim();
    const amount = parseNumber(row.Beløp);
    if (!name || amount === null) continue;
    const nameKey = normalizePersonName(name);
    const id = sourceEmployeeId(row);
    const ambiguousSourceName = identityIndex.idsByName.get(nameKey)?.size > 1;
    const repeatedCounter = id && identityIndex.namesById.get(id)?.size > 1;
    const employee = ambiguousSourceName && !employeeIndex?.byId?.has(id) ? null
      : matchEmployee(repeatedCounter ? { ...row, Ansattnummer: "", "Ansattnr.": "" } : row, employeeIndex, name);
    const personKey = employee?.personKey || sourcePersonKey(row, name, identityIndex);
    resolved.push({ row, personKey, name, amount, employee });
  }
  return resolved;
}

export function buildVariablePayReconciliation(rows, employeeIndex) {
  const all = aggregateVariablePay(rows, employeeIndex);
  const positiveSource = all.filter((person) => person.amount > 0);
  const matched = positiveSource.filter((person) => person.employee);
  const unmatchedPeople = positiveSource.filter((person) => !person.employee)
    .map(({ employee, ...person }) => person)
    .sort((left, right) => right.amount - left.amount || left.name.localeCompare(right.name, "nb"));
  return {
    sourceCount: positiveSource.length, matchedCount: matched.length,
    unmatchedCount: unmatchedPeople.length,
    sourceAmount: amountSum(positiveSource), matchedAmount: amountSum(matched),
    unmatchedAmount: amountSum(unmatchedPeople), unmatchedPeople,
    netSourceAmount: amountSum(all),
    correctionAmount: amountSum(all.filter((person) => person.amount < 0)),
    zeroCount: all.filter((person) => person.amount === 0).length
  };
}

export function summarizeVariablePayParticipants(participants, reconciliation = null) {
  const women = participants.filter((participant) => participant.gender === "Kvinne").length;
  const men = participants.filter((participant) => participant.gender === "Mann").length;
  return {
    total: reconciliation?.sourceCount ?? participants.length,
    matched: participants.length,
    unmatched: reconciliation?.unmatchedCount ?? 0,
    sourceAmount: reconciliation?.sourceAmount ?? amountSum(participants),
    matchedAmount: reconciliation?.matchedAmount ?? amountSum(participants),
    unmatchedAmount: reconciliation?.unmatchedAmount ?? 0,
    women,
    men
  };
}

export function buildParentalLeaveEmployees(rows, year) {
  return buildParentalLeaveEmployeesByYear(rows, year);
}

export function buildParentalLeaveGroupedEmployees(rows, employeeIndex, year) {
  return buildParentalLeaveEmployeesByYear(rows, year).flatMap((person) => {
    const employee = matchEmployee({ Ansattnummer: person.employeeId }, employeeIndex, person.name);
    return employee ? [{ ...person, group: employee.group }] : [];
  });
}

export function buildReport(files, selectedYear, selectedPeriodKey, assignments = {}) {
  const hasSource = (role) => files.some((file) => file.role === role && file.source === 'rådata' && (file.years || [file.fileYear]).includes(selectedYear));
  const { fastlonnFilesForPeriod, targetSnapshotDate } = getFastlonnFilesForSelection(files, selectedYear, selectedPeriodKey);
  const fastlonnRows = buildFastlonnRows(fastlonnFilesForPeriod);
  const overtidRows = filterVariablePayRowsByYear(buildRowsForRole(files, "overtid"), selectedYear);
  const vakttilleggRows = filterVariablePayRowsByYear(buildRowsForRole(files, "vakttillegg"), selectedYear);
  const foreldrepermisjonRows = filterParentalLeaveRowsByYear(buildRowsForRole(files, "foreldrepermisjon"), selectedYear);
  const fastlonn = fastlonnRows.length ? summarizeFastlonn(fastlonnRows, targetSnapshotDate, assignments) : null;
  const employment = fastlonn ? summarizeEmployment(fastlonn.employees) : null;
  const overtid = hasSource('overtid') && fastlonn ? summarizeVariablePay(overtidRows, fastlonn.employeeIndex) : null;
  const vakttillegg = hasSource('vakttillegg') && fastlonn ? summarizeVariablePay(vakttilleggRows, fastlonn.employeeIndex) : null;
  const foreldrepermisjon = hasSource('foreldrepermisjon') ? summarizeParentalLeaveByYear(foreldrepermisjonRows, selectedYear) : null;
  const overtidParticipants = fastlonn ? buildVariablePayParticipants(overtidRows, fastlonn.employeeIndex) : [];
  const vakttilleggParticipants = fastlonn ? buildVariablePayParticipants(vakttilleggRows, fastlonn.employeeIndex) : [];
  const overtidReconciliation = fastlonn ? buildVariablePayReconciliation(overtidRows, fastlonn.employeeIndex) : null;
  const vakttilleggReconciliation = fastlonn ? buildVariablePayReconciliation(vakttilleggRows, fastlonn.employeeIndex) : null;

  const notes = [];
  if (files.some((file) => file.source === "bearbeidet")) {
    notes.push("Bearbeidede filer er lastet opp. Målet er at løsningen skal fungere med rådata alene.");
  }
  if (files.some((file) => file.source === "referanse")) {
    notes.push("Utregningsskjema er lastet opp som referanse. Det skal ikke være nødvendig i den ferdige automatiseringen.");
  }
  if (fastlonn?.unknownPositionCodes?.length) {
    const description = fastlonn.unknownPositionCodes.map(({ code, count }) => `${code} (${count})`).join(", ");
    notes.push(`Ansatte med disse stillingskodene står i «Uavklart rapportgruppe»: ${description}. Avklar kodene før publisering.`);
  }
  if (selectedYear) {
    if (!fastlonnFilesForPeriod.length) {
      notes.push(`Mangler fastlønnfil for ${selectedPeriodKey === "05-01" ? "01.05" : "31.12"} i ${selectedYear}.`);
    }
    if (!hasSource('overtid')) {
      notes.push(`Fant ingen overtidsrader for ${selectedYear} basert på Utbetalingsdato.`);
    }
    if (!hasSource('vakttillegg')) {
      notes.push(`Fant ingen vakttilleggsrader for ${selectedYear} basert på Utbetalingsdato.`);
    }
    if (!hasSource('foreldrepermisjon')) {
      notes.push(`Fant ingen foreldrepermisjonsrader for ${selectedYear} basert på Start og Slutt.`);
    }
  }
  if (overtidReconciliation?.unmatchedCount) {
    notes.push(`${overtidReconciliation.unmatchedCount} ansatte med overtid er ikke i valgt fastlønnsuttrekk. Beløpet ${formatCurrency(overtidReconciliation.unmatchedAmount)} er med i kildetotalen, men ikke fordelt på stillingsgruppe eller kjønn.`);
  }
  if (vakttilleggReconciliation?.unmatchedCount) {
    notes.push(`${vakttilleggReconciliation.unmatchedCount} ansatte med vakttillegg er ikke i valgt fastlønnsuttrekk. Beløpet ${formatCurrency(vakttilleggReconciliation.unmatchedAmount)} er med i kildetotalen, men ikke fordelt på stillingsgruppe eller kjønn.`);
  }

  if (fastlonn) {
    const index = fastlonn.employeeIndex;
    if (index.excludedEmployees.length) notes.push(`${index.excludedEmployees.length} ansatte er eksplisitt utelatt gjennom lokale gruppevalg. Antallet er ikke med i rapporttabellene.`);
    if (index.duplicatePeople.size) notes.push(`${index.duplicatePeople.size} ansatte har motstridende rader i samme uttrekk og er utelatt. Rett kilden før rapportering.`);
    if (index.ambiguousNames.size) notes.push("Navnelike ansatte kan ikke kobles entydig med navn. Bruk ansattnummer i begge kilder.");
    const missingSalary = fastlonn.employees.filter((employee) => employee.salary === null).length;
    if (missingSalary) notes.push(`${missingSalary} ansatte mangler gyldig årslønn. Berørte lønnsgjennomsnitt vises som utilgjengelige.`);
    const missingPartTime = fastlonn.employees.filter((employee) => employee.partTimePercent === null).length;
    if (missingPartTime) notes.push(`${missingPartTime} ansatte mangler gyldig stillingsprosent og kan ikke klassifiseres som heltid eller deltid.`);
    const missingGender = fastlonn.employees.filter((employee) => employee.gender === "Uoppgitt").length;
    if (missingGender) notes.push(`${missingGender} ansatte har uoppgitt eller ukjent kjønn. De inngår i totaltallet, men ikke i kvinne- og mannsandelene.`);
  }
  for (const [label, reconciliation] of [["Overtid", overtidReconciliation], ["Vakttillegg", vakttilleggReconciliation]]) {
    if (reconciliation?.correctionAmount) notes.push(`${label}: ${formatCurrency(reconciliation.correctionAmount)} gjelder personer med negativ nettoutbetaling. Dette inngår i kildens nettobeløp, men ikke i gjennomsnitt blant mottakere med positivt årsbeløp.`);
  }
  const leaveEmployees = buildParentalLeaveEmployeesByYear(foreldrepermisjonRows, selectedYear);
  if (leaveEmployees.some((employee) => employee.gender === 'Uoppgitt')) notes.push('Foreldrepermisjon: uoppgitt kjønn inngår i samlet uttak, men ikke i kvinne- og mannsandelene.');
  const overlap = leaveEmployees.reduce((sum, employee) => sum + employee.overlappingDays, 0);
  if (overlap) notes.push(`Foreldrepermisjon: ${overlap} overlappende arbeidsdager er telt én gang per person. Kontroller overlappene i kilden.`);
  const groupedLeave = fastlonn ? buildParentalLeaveGroupedEmployees(foreldrepermisjonRows, fastlonn.employeeIndex, selectedYear) : [];
  if (leaveEmployees.length > groupedLeave.length) notes.push(`${leaveEmployees.length - groupedLeave.length} ansatte med foreldrepermisjon kan ikke kobles til valgt fastlønnsuttrekk. Totalt uttak inkluderer dem, mens gruppetabellene bare viser koblede ansatte.`);
  for (const file of files.filter((file) => (file.years || [file.fileYear]).includes(selectedYear))) {
    for (const warning of file.warnings || []) notes.push(`${file.label}: ${warning}`);
    if (file.role === 'foreldrepermisjon' && file.fileYear && file.fileYear !== selectedYear) notes.push(`Foreldrepermisjon: filen er oppgitt for ${file.fileYear}, men inneholder perioder som overlapper ${selectedYear}. Kontroller at alle permisjoner i rapportåret er med.`);
  }

  return {
    generatedAt: new Date().toISOString(),
    fastlonnSnapshotLabel: targetSnapshotDate ? `${selectedPeriodKey === "05-01" ? "01.05" : "31.12"}.${selectedYear}` : "",
    fastlonnSnapshotKey: targetSnapshotDate ? `${selectedYear}-${selectedPeriodKey}` : "",
    fastlonn,
    employment,
    overtid,
    overtidParticipants: summarizeVariablePayParticipants(overtidParticipants, overtidReconciliation),
    vakttillegg,
    vakttilleggParticipants: summarizeVariablePayParticipants(vakttilleggParticipants, vakttilleggReconciliation),
    foreldrepermisjon,
    notes
  };
}
