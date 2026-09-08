import { formatInteger, formatPercent, formatNumber, parseNumber } from './formatting.js';

export function createEmptyManualReportInputs() {
  return {
    nyansatteCount: "",
    samletSykefravaer: "",
    sykefravaerKvinner: "",
    sykefravaerMenn: ""
  };
}

export function normalizeManualReportInputs(value) {
  const defaults = createEmptyManualReportInputs();
  const input = value && typeof value === "object" ? value : {};
  return Object.fromEntries(Object.keys(defaults).map((key) => [key,
    typeof input[key] === 'string' || typeof input[key] === 'number' ? String(input[key]).slice(0, 100) : ''
  ]));
}

export function validateManualField(field, value) {
  if (!String(value ?? '').trim()) return 'Fyll inn tallet.';
  const text = String(value).trim();
  const number = parseNumber(field === 'nyansatteCount' ? text : text.replace(/\s*%$/, ''));
  if (field === 'nyansatteCount') return number !== null && Number.isInteger(number) && number >= 0 ? '' : 'Oppgi et heltall som er 0 eller større.';
  return number !== null && number >= 0 && number <= 100 ? '' : 'Oppgi en prosent mellom 0 og 100.';
}

export function manualValue(value, fallbackLabel) {
  const trimmed = String(value ?? "").trim();
  return !validateManualField('nyansatteCount', trimmed) ? formatInteger(parseNumber(trimmed)) : `[MANUELL: ${fallbackLabel}]`;
}

export function manualPercentValue(value, fallbackLabel) {
  const trimmed = String(value ?? "").trim();
  if (validateManualField('percent', trimmed)) return `[MANUELL: ${fallbackLabel}]`;
  return formatPercent(parseNumber(trimmed.replace(/\s*%$/, '')));
}

export function totalGenderRow(summary) {
  return summary?.fastlonn?.genderBalance?.find((row) => row.group === "Totalt antall ansatte") ?? null;
}

export function leaderSummary(summary) {
  const rows = summary?.fastlonn?.genderBalance ?? [];
  const leaderGroups = ["Direktørgruppen", "Seksjonssjefgruppen"];
  const selected = rows.filter((row) => leaderGroups.includes(row.group));
  return {
    total: selected.reduce((sum, row) => sum + row.total, 0),
    women: selected.reduce((sum, row) => sum + row.women, 0)
  };
}

export function standardReportNarratives(summary, selectedYear, manualInputs = createEmptyManualReportInputs()) {
  const year = selectedYear || new Date(summary.generatedAt).getFullYear();
  const totalRow = totalGenderRow(summary);
  const leaders = leaderSummary(summary);
  const snapshot = summary.fastlonnSnapshotLabel || 'valgt uttrekksdato';
  const overtimeParticipants = summary?.overtidParticipants?.total ?? summary?.overtid?.find((row) => row.group === "Mottakere i valgt uttrekk")?.n ?? 0;
  const vakttilleggParticipants = summary?.vakttilleggParticipants?.total ?? summary?.vakttillegg?.find((row) => row.group === "Mottakere i valgt uttrekk")?.n ?? 0;

  return {
    introduction:
      totalRow
        ? `Det ble ansatt ${manualValue(manualInputs.nyansatteCount, "antall nyansatte")} nye medarbeidere i ${year}. Per ${snapshot} inngår ${formatInteger(totalRow.total)} ansatte i rapportgrunnlaget. Virksomheten hadde et samlet sykefravær på ${manualPercentValue(manualInputs.samletSykefravaer, "samlet sykefravær i prosent")}.`
        : "",
    genderBalance:
      totalRow
        ? `Per ${snapshot} var det ${formatPercent(totalRow.womenShare)} kvinner og ${formatPercent(totalRow.menShare)} menn i rapportgrunnlaget. Antall ledere med personalansvar var ${formatInteger(leaders.total)}, hvorav ${formatInteger(leaders.women)} var kvinner.${leaders.total ? ` Andel kvinner i ledelsen er dermed ${formatPercent((leaders.women / leaders.total) * 100)}.` : ''}`
        : "",
    fastlonn:
      summary?.fastlonnSnapshotLabel
        ? `Lønnskartleggingen er basert på fastlønn per ${summary.fastlonnSnapshotLabel}. Tabellen viser gjennomsnittlig oppgitt årslønn per stillingsgruppe fordelt på kvinner og menn. Lønnen vektes ikke med stillingsprosent. Manglende lønnsgrunnlag eller manglende sammenligningsgruppe vises med tankestrek.`
        : "",
    overtid:
      summary?.overtid
        ? variablePayNarrative('overtid', year, overtimeParticipants, summary.overtidParticipants, snapshot, totalRow?.total)
        : "",
    vakttillegg:
      summary?.vakttillegg
        ? variablePayNarrative('vakttillegg', year, vakttilleggParticipants, summary.vakttilleggParticipants, snapshot, totalRow?.total)
        : "",
    employment:
      summary?.employment
        ? `Tabellen viser midlertidige ansatte og faktisk deltid per ${snapshot}, fordelt på kjønn. Deltid betyr registrert stillingsprosent under 100. Opplysningene skiller ikke mellom frivillig og ufrivillig deltid.`
        : "",
    foreldrepermisjon:
      summary?.foreldrepermisjon
        ? `Det var i ${year} totalt ${formatInteger(summary.foreldrepermisjon.totalEmployees)} medarbeidere som tok ut foreldrepermisjon, ${formatInteger(summary.foreldrepermisjon.womenCount)} kvinner og ${formatInteger(summary.foreldrepermisjon.menCount)} menn. Uttak av foreldrepermisjon basert på gjennomsnittlig antall uker fordeler seg med henholdsvis ${formatNumber(summary.foreldrepermisjon.womenAvgWeeks)} uker for kvinner og ${formatNumber(summary.foreldrepermisjon.menAvgWeeks)} uker for menn. Uttak av foreldrepermisjon basert på andel dager av total er på henholdsvis ${formatPercent(summary.foreldrepermisjon.womenShareDays)} for kvinner og ${formatPercent(summary.foreldrepermisjon.menShareDays)} for menn.`
        : "",
    sykefravaer:
      `Det samlede sykefraværet for ${year} var ${manualPercentValue(manualInputs.samletSykefravaer, "samlet sykefravær %")}. Det legemeldte sykefraværet var ${manualPercentValue(manualInputs.sykefravaerKvinner, "legemeldt sykefravær kvinner %")} for kvinner og ${manualPercentValue(manualInputs.sykefravaerMenn, "legemeldt sykefravær menn %")} for menn.`
  };
}

function variablePayNarrative(label, year, total, participants, snapshot, employees) {
  const matched = participants?.matched || 0;
  const unmatched = participants?.unmatched || 0;
  return `I ${year} hadde ${formatInteger(total)} ansatte positiv samlet utbetaling av ${label}. ` +
    `${formatInteger(matched)} av mottakerne kan kobles til fastlønnsuttrekket per ${snapshot}, herav ${formatInteger(participants?.women || 0)} kvinner og ${formatInteger(participants?.men || 0)} menn. ` +
    (unmatched ? `${formatInteger(unmatched)} mottakere er ikke fordelt på kjønn eller gruppe. ` : '') +
    (employees ? `De koblede mottakerne utgjør ${formatPercent(matched / employees * 100)} av de ansatte i dette uttrekket. ` : '') +
    `Tabellen viser gjennomsnitt blant koblede mottakere med positivt årsbeløp. Personer uten utbetaling inngår ikke i gjennomsnittet. Året bestemmes av utbetalingsdatoen.`;
}

export function reportHistoryYears(selectedYear, summary, count = 6) {
  const endYear = Number(selectedYear || new Date(summary.generatedAt).getFullYear());
  return Array.from({ length: count }, (_, index) => String(endYear - (count - 1) + index));
}
