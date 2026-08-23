import test from "node:test";
import assert from "node:assert/strict";
import {
  countNorwegianWorkingDays,
  norwegianPublicHolidayKeys,
  selectExactSnapshotFiles,
  summarizeParentalLeaveByYear,
  validateWorkbookFile
} from "../src/lib/report-engine.js";

test("norske bevegelige helligdager for 2025 følger kalenderen", () => {
  const holidays = norwegianPublicHolidayKeys(2025);
  for (const date of ["2025-04-17", "2025-04-18", "2025-04-21", "2025-05-29", "2025-06-09"]) {
    assert.equal(holidays.has(date), true, `${date} skal være helligdag`);
  }
});

test("opplastingsgrensen avviser feil format og store arbeidsbøker", () => {
  assert.throws(() => validateWorkbookFile({ name: "ansatte.xls", size: 100 }), /XLSX-fil/);
  assert.throws(() => validateWorkbookFile({ name: "ansatte.xlsx", size: 16 * 1024 * 1024 }), /15 MB/);
  assert.equal(validateWorkbookFile({ name: "ansatte.XLSX", size: 1024 }), true);
});

test("et uttrekk fra 31.12 kan ikke brukes som fasit for 01.05", () => {
  const files = [
    { snapshotKey: "2025-12-31", fileName: "desember.xlsx" },
    { snapshotKey: "2024-05-01", fileName: "mai-i-fjor.xlsx" }
  ];
  assert.deepEqual(selectExactSnapshotFiles(files, "2025", "05-01"), []);
  assert.equal(selectExactSnapshotFiles(files, "2025", "12-31")[0].fileName, "desember.xlsx");
});

test("arbeidsdager avgrenses til rapportåret og utelater helg og helligdag", () => {
  assert.equal(countNorwegianWorkingDays("23.12.2024", "06.01.2025", 2025), 3);
  assert.equal(countNorwegianWorkingDays("16.04.2025", "22.04.2025", 2025), 2);
});

test("foreldrepermisjon bruker kalenderperioden, ikke SAP-total eller arbeidsførhet", () => {
  const rows = [
    { Fornavn: "Ada", Etternavn: "Nord", Kjønn: "Kvinne", Start: "2024-12-23", Slutt: "2025-01-06", "Frav.dager": 99, Arbeidsførhet: 50 },
    { Fornavn: "Ada", Etternavn: "Nord", Kjønn: "Kvinne", Start: "2025-04-16", Slutt: "2025-04-22", "Frav.dager": 99, Arbeidsførhet: 0 },
    { Fornavn: "Ola", Etternavn: "Sør", Kjønn: "Mann", Start: "2025-05-02", Slutt: "2025-05-05", "Frav.dager": 50, Arbeidsførhet: 60 }
  ];

  assert.deepEqual(summarizeParentalLeaveByYear(rows, 2025), {
    totalEmployees: 2,
    womenCount: 1,
    menCount: 1,
    womenDays: 5,
    menDays: 2,
    totalDays: 7,
    womenAvgWeeks: 1,
    menAvgWeeks: 0.4,
    womenShareDays: 71.42857142857143,
    menShareDays: 28.57142857142857
  });
});
