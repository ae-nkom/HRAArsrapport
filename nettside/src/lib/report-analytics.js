import { groupOrder, variablePayOrder, variablePayGroup, amountSum, resolveVariablePayRows, buildVariablePayParticipants, unassignedGroup } from './report-engine.js';
import { quantileSorted } from './formatting.js';
import { parseReportDate as parseCellDate } from './date-utils.js';

export const monthNames = ["Jan", "Feb", "Mar", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Des"];

export function monthKeyFromDate(date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

export function formatMonthLabel(monthKey) {
  const [year, month] = String(monthKey).split("-");
  const monthIndex = Number(month) - 1;
  return `${monthNames[monthIndex] || month} ${year}`;
}

export function buildVariablePayGroupMetrics(participants, employees) {
  const eligibleEmployees = employees;

  return variablePayOrder.map((groupName) => {
    const employeeSubset =
      groupName === "Mottakere i valgt uttrekk"
        ? eligibleEmployees
        : eligibleEmployees.filter((employee) => variablePayGroup(employee.group) === groupName);
    const participantSubset =
      groupName === "Mottakere i valgt uttrekk"
        ? participants
        : participants.filter((employee) => variablePayGroup(employee.group) === groupName);
    const womenEmployees = employeeSubset.filter((employee) => employee.gender === "Kvinne");
    const menEmployees = employeeSubset.filter((employee) => employee.gender === "Mann");
    const womenParticipants = participantSubset.filter((employee) => employee.gender === "Kvinne");
    const menParticipants = participantSubset.filter((employee) => employee.gender === "Mann");
    const totalEmployees = employeeSubset.length;
    const totalParticipants = participantSubset.length;
    const totalAmount = amountSum(participantSubset);

    return {
      group: groupName,
      totalEmployees,
      womenEmployees: womenEmployees.length,
      menEmployees: menEmployees.length,
      participants: totalParticipants,
      womenParticipants: womenParticipants.length,
      menParticipants: menParticipants.length,
      participationShare: totalEmployees ? (totalParticipants / totalEmployees) * 100 : 0,
      womenParticipationShare: womenEmployees.length ? (womenParticipants.length / womenEmployees.length) * 100 : 0,
      menParticipationShare: menEmployees.length ? (menParticipants.length / menEmployees.length) * 100 : 0,
      womenParticipantShare: totalParticipants ? (womenParticipants.length / totalParticipants) * 100 : 0,
      menParticipantShare: totalParticipants ? (menParticipants.length / totalParticipants) * 100 : 0,
      totalAmount,
      womenAmount: amountSum(womenParticipants),
      menAmount: amountSum(menParticipants),
      totalAvgPerParticipant: totalParticipants ? totalAmount / totalParticipants : 0
    };
  }).filter((row) => row.group !== unassignedGroup || row.totalEmployees > 0);
}

export function buildVariablePaySpreadRows(participants) {
  return variablePayOrder
    .filter((groupName) => groupName !== "Mottakere i valgt uttrekk")
    .map((groupName) => {
      const amounts = participants
        .filter((employee) => variablePayGroup(employee.group) === groupName)
        .map((employee) => employee.amount)
        .filter((amount) => amount > 0)
        .sort((left, right) => left - right);

      if (!amounts.length) return null;

      return {
        gruppe: groupName,
        min: amounts[0],
        q1: quantileSorted(amounts, 0.25),
        median: quantileSorted(amounts, 0.5),
        q3: quantileSorted(amounts, 0.75),
        max: amounts[amounts.length - 1]
      };
    })
    .filter(Boolean);
}

export function buildVariablePayMonthlyRows(rows, employeeIndex) {
  const totals = new Map();
  const participants = new Set(buildVariablePayParticipants(rows, employeeIndex).map((person) => person.personKey));

  for (const { row, employee, amount, personKey } of resolveVariablePayRows(rows, employeeIndex)) {
    const payoutDate = parseCellDate(row["Utbetalingsdato"]);
    if (!payoutDate || !employee || !participants.has(personKey)) continue;

    const key = monthKeyFromDate(payoutDate);
    const current = totals.get(key) || {
      måned: formatMonthLabel(key),
      sortKey: key,
      kvinner: 0,
      menn: 0,
      totalt: 0
    };

    if (employee.gender === "Kvinne") {
      current.kvinner += amount;
    } else if (employee.gender === "Mann") {
      current.menn += amount;
    }

    current.totalt += amount;
    totals.set(key, current);
  }

  // Empty months within an imported year are actual zeros, not skipped positions
  // on the time axis. The files represent a complete annual extract.
  const years = new Set([...totals.keys()].map((key) => key.slice(0, 4)));
  for (const year of years) for (let month = 1; month <= 12; month++) {
    const key = `${year}-${String(month).padStart(2, '0')}`;
    if (!totals.has(key)) totals.set(key, { måned: formatMonthLabel(key), sortKey: key, kvinner: 0, menn: 0, totalt: 0 });
  }
  return [...totals.values()].sort((left, right) => left.sortKey.localeCompare(right.sortKey));
}

export function buildVariablePayMonthHeatmap(rows, employeeIndex) {
  const totals = new Map();
  const participants = new Set(buildVariablePayParticipants(rows, employeeIndex).map((person) => person.personKey));

  for (const { row, employee, amount, personKey } of resolveVariablePayRows(rows, employeeIndex)) {
    const payoutDate = parseCellDate(row["Utbetalingsdato"]);
    if (!payoutDate || !participants.has(personKey)) continue;
    const group = employee?.group ? variablePayGroup(employee.group) : "";
    if (!group) continue;

    const monthKey = monthKeyFromDate(payoutDate);
    const compoundKey = `${group}::${monthKey}`;
    const current = totals.get(compoundKey) || {
      gruppe: group,
      måned: formatMonthLabel(monthKey),
      sortKey: monthKey,
      verdi: 0
    };

    current.verdi += amount;
    totals.set(compoundKey, current);
  }

  return [...totals.values()].sort((left, right) => {
    const groupDiff = variablePayOrder.indexOf(left.gruppe) - variablePayOrder.indexOf(right.gruppe);
    if (groupDiff !== 0) return groupDiff;
    return left.sortKey.localeCompare(right.sortKey);
  });
}

export function buildParentalLeaveGroupCountChartData(employees) {
  return groupOrder
    .filter((group) => group !== 'Totalt antall ansatte')
    .map((groupName) => {
      const subset = employees.filter((employee) => employee.group === groupName);
      return [
        { gruppe: groupName, kjønn: "Kvinner", verdi: subset.filter((employee) => employee.gender === "Kvinne").length },
        { gruppe: groupName, kjønn: "Menn", verdi: subset.filter((employee) => employee.gender === "Mann").length }
      ];
    })
    .flat()
    .filter((row) => row.verdi > 0);
}

export function buildParentalLeaveGroupParticipationData(participants, employees) {
  return groupOrder
    .filter((group) => group !== 'Totalt antall ansatte')
    .map((groupName) => {
      const employeeSubset = employees.filter((employee) => employee.group === groupName);
      const participantSubset = participants.filter((employee) => employee.group === groupName);
      return {
        gruppe: groupName,
        andel: employeeSubset.length ? participantSubset.length / employeeSubset.length : 0
      };
    })
    .filter((row) => row.andel > 0);
}

export function buildParentalLeaveSpreadRows(employees) {
  return groupOrder
    .filter((group) => group !== 'Totalt antall ansatte')
    .map((groupName) => {
      const values = employees
        .filter((employee) => employee.group === groupName)
        .map((employee) => employee.weeks)
        .filter((value) => value > 0)
        .sort((left, right) => left - right);

      if (!values.length) return null;

      return {
        gruppe: groupName,
        min: values[0],
        q1: quantileSorted(values, 0.25),
        median: quantileSorted(values, 0.5),
        q3: quantileSorted(values, 0.75),
        max: values[values.length - 1]
      };
    })
    .filter(Boolean);
}
