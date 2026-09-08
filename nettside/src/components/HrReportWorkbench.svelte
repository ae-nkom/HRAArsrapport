<script>
  import ReportTable from './ReportTable.svelte';
  import GroupAssignments from './GroupAssignments.svelte';
  import { readExcelFiles } from '../lib/workbook-client.js';
  import { fileCoversYear } from '../lib/workbook-import.js';
  import { validateManualField } from '../lib/report-text.js';
  let groupAssignmentCache = {};
  $: assignmentKey = selectedFastlonnYear + '-' + selectedFastlonnPeriodKey;
  $: currentGroupAssignments = groupAssignmentCache[assignmentKey] || {};
  $: classificationEmployees = report?.fastlonn?.employeeIndex.allEmployees.filter((employee) => employee.requiresGroupChoice) || [];
  function setGroupAssignment(personKey, group) {
    groupAssignmentCache = { ...groupAssignmentCache, [assignmentKey]: { ...currentGroupAssignments, [personKey]: group } };
  }
  function resetReplacedAssignments(files) {
    const next = { ...groupAssignmentCache };
    for (const file of files) if (file.role === 'fastlønn') delete next[file.snapshotKey];
    groupAssignmentCache = next;
  }

  import { formatNumber, formatPercent, formatInteger, formatCurrency } from "../lib/formatting.js";
  import { buildRowsForRole, filterVariablePayRowsByYear, filterParentalLeaveRowsByYear, buildVariablePayParticipants, buildVariablePayReconciliation, buildParentalLeaveEmployees, buildReport } from "../lib/report-engine.js";
  
  import { compareSnapshotKeys, fileIdentity, mergeFiles } from "../lib/workbook-import.js";
  import { createEmptyManualReportInputs, normalizeManualReportInputs } from "../lib/report-text.js";
  import { parseReportDate as parseCellDate } from "../lib/date-utils.js";
  async function downloadReport() {
    if (!report?.fastlonn || exporting || hasInvalidManualInputs) return;
    const summary = { ...report, notes: [...new Set([...report.notes, ...datasetReadinessIssues.map((issue) => issue.detail || issue.title)])] };
    const options = { year: selectedFastlonnYear, files: [...uploadedFiles], manualInputs: { ...manualReportInputs } };
    exporting = true;
    error = "";
    try {
      const { createReportDocument } = await import("../lib/report-document.js");
      const { Packer } = await import("docx");
      const doc = createReportDocument(summary, options);
      const url = URL.createObjectURL(await Packer.toBlob(doc));
      const link = document.createElement("a");
      link.href = url;
      link.download = `hr-arsrapport-${options.year}.docx`;
      document.body.append(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch {
      error = "Kunne ikke lage Word-rapporten. Prøv igjen.";
    } finally {
      exporting = false;
    }
  }
  let exporting = false;
  let mounted = false;

  import BarChart from "./charts/BarChart.svelte";

  import { onMount, tick } from "svelte";
  import { consumeSelectedFiles } from "../lib/report-engine.js";

  const genderColumns = [
    {key:'group',label:'Stillingsgruppe'}, {key:'women',label:'Kvinner',numeric:true},
    {key:'womenShare',label:'Andel kvinner',numeric:true}, {key:'men',label:'Menn',numeric:true},
    {key:'menShare',label:'Andel menn',numeric:true}, {key:'total',label:'Totalt',numeric:true}
  ];
  const salaryColumns = [
    {key:'group',label:'Stillingsgruppe'}, {key:'n',label:'Antall',numeric:true},
    {key:'women',label:'Kvinner',numeric:true}, {key:'men',label:'Menn',numeric:true},
    {key:'total',label:'Totalt',numeric:true}, {key:'ratio',label:'Kvinner i % av menn',numeric:true}
  ];
  const employmentColumns = [
    {key:'kategori',label:'Arbeidsforhold'}, {key:'kvinnerAntall',label:'Kvinner',numeric:true},
    {key:'kvinnerAndel',label:'Andel kvinner',numeric:true}, {key:'mennAntall',label:'Menn',numeric:true},
    {key:'mennAndel',label:'Andel menn',numeric:true}
  ];
  const leaveColumns = [
    {key:'måling',label:'Måling'}, {key:'kvinner',label:'Kvinner',numeric:true}, {key:'menn',label:'Menn',numeric:true}
  ];
  function moneyTableRows(rows) {
    return (rows || []).map(row => ({group:row.group.replace(/\s*\(N=\d+\)$/, ''), n:formatInteger(row.n),
      women:formatCurrency(row.womenAvg), men:formatCurrency(row.menAvg), total:formatCurrency(row.totalAvg), ratio:formatPercent(row.womenPctOfMen)}));
  }
  $: genderTableRows = genderBalanceRows.map(row => ({...row, womenShare:formatPercent(row.womenShare), menShare:formatPercent(row.menShare)}));
  $: employmentTableRows = employmentAnalysisRows.map(row => ({...row, kvinnerAndel:formatPercent(row.kvinnerAndel), mennAndel:formatPercent(row.mennAndel)}));
  $: leaveTableRows = foreldrepermisjonSummaryRows.map(row => ({...row,
    kvinner:row.måling === 'Andel permisjonsdager' ? formatPercent(row.kvinner) : formatNumber(row.kvinner),
    menn:row.måling === 'Andel permisjonsdager' ? formatPercent(row.menn) : formatNumber(row.menn)}));

  const requiredRoles = ["fastlønn", "overtid", "vakttillegg", "foreldrepermisjon"];
  const fileSelectionWatchers = new Map();
  const uploadCardConfig = [
    {
      role: "fastlønn",
      title: "Fastlønn",
      description: "Bemanning og lønn."
    },
    {
      role: "overtid",
      title: "Overtid",
      description: "Overtid per gruppe."
    },
    {
      role: "vakttillegg",
      title: "Vakttillegg",
      description: "Vakt og beredskap."
    },
    {
      role: "foreldrepermisjon",
      title: "Foreldrepermisjon",
      description: "Uttak og fordeling."
    }
  ];
  const graphColors = {
    darkPurple: "#3028AA",
    orange: "#F15B0A",
    lightPurple: "#345FED",
    yellow: "#F9D649",
    green: "#00A166",
    pink: "#DB2481",
    neutral: "#EFEFEF",
    lightPurpleUu: "#6085FF",
    black: "#000000"
  };
  const chartPalette = [
    graphColors.darkPurple,
    graphColors.orange,
    graphColors.lightPurple,
    graphColors.yellow,
    graphColors.green,
    graphColors.pink,
    graphColors.lightPurpleUu
  ];
  const genderSeriesColors = {
    Kvinner: graphColors.pink,
    Menn: graphColors.darkPurple
  };
  const paySeriesColors = {
    Kvinner: graphColors.pink,
    Menn: graphColors.darkPurple
  };
  const leaveSeriesColors = {
    Kvinner: graphColors.pink,
    Menn: graphColors.darkPurple
  };
  const manualReportStorageKey = "hr-arsrapport-manual-report-inputs";
  const readinessExpandedStorageKey = "hr-arsrapport-readiness-expanded";
  const groupChartColors = {
    "Totalt antall ansatte": graphColors.darkPurple,
    Direktørgruppen: graphColors.orange,
    Seksjonssjefgruppen: graphColors.lightPurple,
    Fagsjefgruppen: graphColors.green,
    Seniorpersonale: graphColors.pink,
    "Øvrige saksbehandlere og andre": graphColors.lightPurpleUu,
    "Mottakere i valgt uttrekk": graphColors.darkPurple
  };
  const norwegianCharacterPattern = /[æøåÆØÅ]/;
  const tabLabels = {
    opplasting: "Datagrunnlag",
    fastlonn: "Fastlønn",
    overtid: "Overtid",
    vakttillegg: "Vakttillegg",
    foreldrepermisjon: "Foreldrepermisjon",
    arsrapport: "Årsrapport"
  };
  let uploadedFiles = [];
  let report = null;
  let loading = false;
  let error = "";
  let dataSourceLabel = "Ingen filer lastet opp";
  let expandedView = null;
  let expandedDialog;
  let expandedCloseButton;
  let focusBeforeDialog;
  let selectedFastlonnYear = "";
  let selectedFastlonnPeriodKey = "12-31";
  let activeTab = "opplasting";
  let sidebarCollapsed = false;
  let readinessExpanded = false;
  let fastlonnPopupColumnMode = "default";
  let overtidPopupColumnMode = "default";
  let vakttilleggPopupColumnMode = "default";
  let foreldrepermisjonPopupColumnMode = "default";
  let fastlonnTopScroll;
  let fastlonnBottomScroll;
  let overtidTopScroll;
  let overtidBottomScroll;
  let vakttilleggTopScroll;
  let vakttilleggBottomScroll;
  let foreldrepermisjonTopScroll;
  let foreldrepermisjonBottomScroll;
  let fastlonnTopScrollWidth = 0;
  let overtidTopScrollWidth = 0;
  let vakttilleggTopScrollWidth = 0;
  let foreldrepermisjonTopScrollWidth = 0;
  let manualReportInputs = createEmptyManualReportInputs();
  let manualReportInputCache = {};
  let loadedManualReportInputKey = "";
  let controlTableRows = {};
  let controlTablePages = {};
  const controlTablePageSizes = [10, 25, 50, 100];
  const fileOverviewPageSizes = [10, 25, 50, 100];
  let fileOverviewRows = 10;
  let fileOverviewPage = 1;

  $: coverage = requiredRoles.map((role) => ({
    role,
    present: uploadedFiles.some((file) => file.role === role && file.source === "rådata" && fileCoversYear(file, selectedFastlonnYear))
  }));
  $: roleUploads = Object.fromEntries(
    requiredRoles.map((role) => [role, uploadedFiles.find((file) => file.role === role && file.source === "rådata" && fileCoversYear(file, selectedFastlonnYear)) ?? null])
  );
  $: missingRoles = coverage.filter((item) => !item.present).map((item) => item.role);
  $: fastlonnSnapshots = uploadedFiles
    .filter((file) => file.role === "fastlønn" && file.source === "rådata")
    .sort((a, b) => compareSnapshotKeys(a.snapshotKey, b.snapshotKey));
  $: availableFastlonnYears = [...new Set(fastlonnSnapshots.map((file) => file.fileYear || file.snapshotKey?.slice(0, 4)).filter(Boolean))]
    .filter((year) => year !== "udat")
    .sort((a, b) => b.localeCompare(a));
  $: selectedFastlonnYearSnapshots = fastlonnSnapshots.filter(
    (file) => (file.fileYear || file.snapshotKey?.slice(0, 4)) === selectedFastlonnYear
  );
  $: if (availableFastlonnYears.length) {
    if (!availableFastlonnYears.includes(selectedFastlonnYear)) {
      selectedFastlonnYear = availableFastlonnYears[0];
    }
  } else if (selectedFastlonnYear) {
    selectedFastlonnYear = "";
  }
  $: fastlonnPeriodChoices = [
    { key: "05-01", label: "01.05" },
    { key: "12-31", label: "31.12" }
  ]
    .map((period) => ({
      ...period,
      exactSnapshot: selectedFastlonnYearSnapshots.find((file) => file.snapshotKey?.slice(5) === period.key) ?? null,
      available: Boolean(selectedFastlonnYearSnapshots.find((file) => file.snapshotKey?.slice(5) === period.key))
    }));
  $: {
    const fallbackPeriod = fastlonnPeriodChoices.find((period) => period.available)?.key ?? "12-31";
    const selectedPeriodAvailable = fastlonnPeriodChoices.some(
      (period) => period.key === selectedFastlonnPeriodKey && period.available
    );
    if (!selectedPeriodAvailable && selectedFastlonnPeriodKey !== fallbackPeriod) {
      selectedFastlonnPeriodKey = fallbackPeriod;
    }
  }
  $: uploadedFilesOverview = uploadedFiles.map((file) => ({
    Fil: file.fileName,
    Type: file.label,
    Kilde: file.source,
    "Lønnsdato": file.snapshotLabel || "—",
    "Filår": file.fileYear || "—",
    Rader: file.rowCount,
    Kolonner: file.columnCount,
    "Kolonner i fil": file.header.join(" | ")
  }));
  $: fileOverviewPageData = getPagedRows(uploadedFilesOverview, fileOverviewPage, fileOverviewRows);
  $: selectedFastlonnPeriodChoice = fastlonnPeriodChoices.find((period) => period.key === selectedFastlonnPeriodKey) ?? null;
  $: activeManualReportInputKey = selectedFastlonnYear || "default";
  $: if (typeof window !== "undefined" && activeManualReportInputKey !== loadedManualReportInputKey) {
    manualReportInputs = normalizeManualReportInputs(manualReportInputCache?.[activeManualReportInputKey]);
    loadedManualReportInputKey = activeManualReportInputKey;
  }
  $: activeTabLabel = tabLabels[activeTab] || "HR Årsrapport";
  $: if (typeof window !== "undefined" && expandedView?.id === "fastlonn-employee-preview-table" && fastlonnPopupColumnMode === "all") {
    syncPopupScrollbars("fastlonn");
  }
  $: if (typeof window !== "undefined" && expandedView?.id === "overtid-table" && overtidPopupColumnMode === "all") {
    syncPopupScrollbars("overtid");
  }
  $: if (typeof window !== "undefined" && expandedView?.id === "vakttillegg-table" && vakttilleggPopupColumnMode === "all") {
    syncPopupScrollbars("vakttillegg");
  }
  $: if (typeof window !== "undefined" && expandedView?.id === "foreldrepermisjon-table" && foreldrepermisjonPopupColumnMode === "all") {
    syncPopupScrollbars("foreldrepermisjon");
  }
  $: report = uploadedFiles.length ? buildReport(uploadedFiles, selectedFastlonnYear, selectedFastlonnPeriodKey, currentGroupAssignments) : null;

  $: overtidSourceFiles = uploadedFiles.filter(
    (file) => file.role === "overtid" && file.source === "rådata" && fileCoversYear(file, selectedFastlonnYear)
  );
  $: vakttilleggSourceFiles = uploadedFiles.filter(
    (file) => file.role === "vakttillegg" && file.source === "rådata" && fileCoversYear(file, selectedFastlonnYear)
  );
  
  $: overtidSourceRows = filterVariablePayRowsByYear(buildRowsForRole(uploadedFiles, "overtid"), selectedFastlonnYear);
  $: overtidSourceHeaders = (() => {
    const fileHeaders = [...new Set(overtidSourceFiles.flatMap((file) => file.header || []))]
      .map((header) => String(header).trim())
      .filter((header) => header && !header.startsWith("__"));
    if (fileHeaders.length) return fileHeaders;
    return Object.keys(overtidSourceRows[0] || {}).filter((header) => !header.startsWith("__"));
  })();
  $: vakttilleggSourceRows = filterVariablePayRowsByYear(buildRowsForRole(uploadedFiles, "vakttillegg"), selectedFastlonnYear);
  $: vakttilleggSourceHeaders = (() => {
    const fileHeaders = [...new Set(vakttilleggSourceFiles.flatMap((file) => file.header || []))]
      .map((header) => String(header).trim())
      .filter((header) => header && !header.startsWith("__"));
    if (fileHeaders.length) return fileHeaders;
    return Object.keys(vakttilleggSourceRows[0] || {}).filter((header) => !header.startsWith("__"));
  })();
  $: foreldrepermisjonSourceRows = filterParentalLeaveRowsByYear(buildRowsForRole(uploadedFiles, "foreldrepermisjon"), selectedFastlonnYear);
  
  $: overtidParticipants = report?.fastlonn ? buildVariablePayParticipants(overtidSourceRows, report.fastlonn.employeeIndex) : [];
  $: vakttilleggParticipants = report?.fastlonn ? buildVariablePayParticipants(vakttilleggSourceRows, report.fastlonn.employeeIndex) : [];
  $: overtidReconciliation = report?.fastlonn ? buildVariablePayReconciliation(overtidSourceRows, report.fastlonn.employeeIndex) : null;
  $: vakttilleggReconciliation = report?.fastlonn ? buildVariablePayReconciliation(vakttilleggSourceRows, report.fastlonn.employeeIndex) : null;
  $: foreldrepermisjonEmployees = buildParentalLeaveEmployees(foreldrepermisjonSourceRows, selectedFastlonnYear);
  
  $: fastlonnEmployeePreview = report?.fastlonn?.employees
    ? [...report.fastlonn.employees]
        .sort((left, right) => left.group.localeCompare(right.group, "nb") || left.name.localeCompare(right.name, "nb"))
        .slice(0, 20)
    : [];
  
  $: genderBalanceRows = report?.fastlonn?.genderBalance ?? [];
  $: salaryComparisonRows = report?.fastlonn?.fastlonn ?? [];

  $: genderBalanceChartData = genderBalanceRows.flatMap((row) => [
    { gruppe: row.group, kjønn: "Kvinner", antall: row.women },
    { gruppe: row.group, kjønn: "Menn", antall: row.men }
  ]);
  
  $: salaryComparisonChartData = salaryComparisonRows.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenAvg },
    { gruppe: row.group, serie: "Menn", verdi: row.menAvg }
  ]);
  
  $: employmentAnalysisRows = report?.employment
    ? [
        {
          kategori: "Midlertidig ansatte",
          kvinnerAntall: report.employment.temporary.women,
          kvinnerAndel: report.employment.totalWomen ? (report.employment.temporary.women / report.employment.totalWomen) * 100 : 0,
          mennAntall: report.employment.temporary.men,
          mennAndel: report.employment.totalMen ? (report.employment.temporary.men / report.employment.totalMen) * 100 : 0
        },
        {
          kategori: "Faktisk deltid",
          kvinnerAntall: report.employment.partTime.women,
          kvinnerAndel: report.employment.totalWomen ? (report.employment.partTime.women / report.employment.totalWomen) * 100 : 0,
          mennAntall: report.employment.partTime.men,
          mennAndel: report.employment.totalMen ? (report.employment.partTime.men / report.employment.totalMen) * 100 : 0
        }
      ]
    : [];
  $: overtidAnalysisRows = report?.overtid ?? [];
  $: vakttilleggAnalysisRows = report?.vakttillegg ?? [];
  $: overtidChartData = overtidAnalysisRows.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenAvg },
    { gruppe: row.group, serie: "Menn", verdi: row.menAvg }
  ]);

  $: vakttilleggChartData = vakttilleggAnalysisRows.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenAvg },
    { gruppe: row.group, serie: "Menn", verdi: row.menAvg }
  ]);

  $: foreldrepermisjonSummaryRows = report?.foreldrepermisjon
    ? [
        {
          måling: "Ansatte med uttak",
          kvinner: report.foreldrepermisjon.womenCount,
          menn: report.foreldrepermisjon.menCount,
          totalt: report.foreldrepermisjon.totalEmployees
        },
        {
          måling: "Gjennomsnittlige uker",
          kvinner: report.foreldrepermisjon.womenAvgWeeks,
          menn: report.foreldrepermisjon.menAvgWeeks,
          totalt:
            report.foreldrepermisjon.totalEmployees
              ? ((report.foreldrepermisjon.womenAvgWeeks * report.foreldrepermisjon.womenCount) +
                  (report.foreldrepermisjon.menAvgWeeks * report.foreldrepermisjon.menCount)) /
                report.foreldrepermisjon.totalEmployees
              : 0
        },
        {
          måling: "Andel permisjonsdager",
          kvinner: report.foreldrepermisjon.womenShareDays,
          menn: report.foreldrepermisjon.menShareDays,
          totalt: report.foreldrepermisjon.totalDays ? 100 : null
        }
      ]
    : [];
  $: foreldrepermisjonWeeksChartData = report?.foreldrepermisjon
    ? [
        { måling: "Gjennomsnittlige uker", kjønn: "Kvinner", verdi: report.foreldrepermisjon.womenAvgWeeks },
        { måling: "Gjennomsnittlige uker", kjønn: "Menn", verdi: report.foreldrepermisjon.menAvgWeeks }
      ]
    : [];

  $: activeSourceSummaryRows = requiredRoles.map((role) => {
    const files = uploadedFiles.filter((file) => file.role === role && file.source === "rådata");
    const latest = files[0] ?? null;
    return {
      tema:
        role === "fastlønn"
          ? "Fastlønn"
          : role === "overtid"
            ? "Overtid"
            : role === "vakttillegg"
              ? "Vakttillegg"
              : "Foreldrepermisjon",
      filer: files.length,
      fil: latest?.fileName || "Mangler",
      år: latest?.fileYear || "—",
      uttrekk: latest?.snapshotLabel || "—",
      rader: files.reduce((sum, file) => sum + (file.rowCount || 0), 0)
    };
  });

  $: globalContextItems = [
    { label: "År", value: selectedFastlonnYear || "Ikke valgt" },
    { label: "Uttrekk", value: selectedFastlonnPeriodChoice?.label || "—" },
    { label: "Datakilde", value: dataSourceLabel },
    { label: "Rådatafiler", value: formatInteger(activeSourceSummaryRows.reduce((sum, row) => sum + row.filer, 0)) }
  ];

  $: datasetReadinessIssues = (() => {
    if (!uploadedFiles.length) return [];
    const issues = missingRoles.length ? [{ id: 'missing-files', title: 'Påkrevde filer mangler', detail: 'Last opp ' + missingRoles.map(roleLabel).join(', ') + '.', action: 'Bruk rådatauttrekkene for rapportåret.' }] : [];
    for (const [index, note] of (report?.notes || []).entries()) {
      if (note.includes('er eksplisitt utelatt gjennom lokale gruppevalg')) continue;
      issues.push({ id: 'report-note-' + index, title: note, detail: '', action: note.includes('Uavklart rapportgruppe') ? 'Velg rapportgruppe under Datagrunnlag.' : '' });
    }
    for (const [id, label, reconciliation] of [['overtime', 'overtid', overtidReconciliation], ['guard', 'vakttillegg', vakttilleggReconciliation]]) {
      if (reconciliation?.unmatchedPeople.length) {
        const issue = issues.find((issue) => issue.title.includes('ansatte med ' + label + ' er ikke'));
        if (issue) issue.people = reconciliation.unmatchedPeople;
      }
    }
    return issues;
  })();
  $: manualInputTasks = [
    { id: "new-hires", label: "Antall nyansatte", field: "nyansatteCount", description: "Brukes i årsrapportens bemanningsomtale." },
    { id: "sick-total", label: "Samlet sykefravær", field: "samletSykefravaer", description: "Oppgis i prosent." },
    { id: "sick-women", label: "Legemeldt sykefravær kvinner", field: "sykefravaerKvinner", description: "Oppgis i prosent." },
    { id: "sick-men", label: "Legemeldt sykefravær menn", field: "sykefravaerMenn", description: "Oppgis i prosent." }
  ].map((task) => ({
    ...task,
    complete: !validateManualField(task.field, manualReportInputs[task.field])
  }));
  $: hasInvalidManualInputs = Object.entries(manualReportInputs).some(([field, value]) => String(value).trim() && validateManualField(field, value));
  $: incompleteManualInputCount = manualInputTasks.filter((task) => !task.complete).length;
  const manualAssessmentTasks = [
    "Fyll inn historikken for rekruttering og inkludering i Word-utkastet.",
    "Beskriv samarbeid med utdanningsinstitusjoner og arbeidet med aktivitets- og redegjørelsesplikten (ARP).",
    "Vurder lønnsforskjeller og frivillig eller ufrivillig deltid ut fra faglig kjennskap til virksomheten."
  ];

  function roleLabel(role) {
    return ({
      fastlønn: "fastlønn",
      overtid: "overtid",
      vakttillegg: "vakttillegg",
      foreldrepermisjon: "foreldrepermisjon"
    })[role] || role;
  }

  function countDistinctValues(values, normalizer = (value) => value) {
    return new Set(values.map((value) => normalizer(value)).filter(Boolean)).size;
  }

  function fileControlKey(file) {
    return `${file.role}:${file.fileName}:${file.snapshotKey || file.fileYear || "na"}`;
  }

  function getControlTableRows(file) {
    return controlTableRows[fileControlKey(file)] ?? 10;
  }

  function setControlTableRows(file, value) {
    const nextRows = Number(value);
    controlTableRows = {
      ...controlTableRows,
      [fileControlKey(file)]: nextRows
    };
    controlTablePages = {
      ...controlTablePages,
      [fileControlKey(file)]: 1
    };
  }

  function getControlTablePage(file) {
    return controlTablePages[fileControlKey(file)] ?? 1;
  }

  function setControlTablePage(file, value) {
    controlTablePages = {
      ...controlTablePages,
      [fileControlKey(file)]: value
    };
  }

  function fileColumns(file) {
    return file?.header || Object.keys(file?.rows?.[0] || {}).filter((key) => !key.startsWith("__"));
  }

  function getControlTableView(file) {
    return getPagedRows(file?.rows || [], getControlTablePage(file), getControlTableRows(file));
  }

  function getPagedRows(rows, page, pageSize) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const safePageSize = Math.max(1, Number(pageSize) || 10);
    const totalPages = Math.max(1, Math.ceil(safeRows.length / safePageSize));
    const currentPage = Math.min(Math.max(1, Number(page) || 1), totalPages);
    const start = (currentPage - 1) * safePageSize;

    return {
      totalPages,
      currentPage,
      startRow: safeRows.length ? start + 1 : 0,
      endRow: Math.min(start + safePageSize, safeRows.length),
      rows: safeRows.slice(start, start + safePageSize)
    };
  }

  function previousPage(page) {
    return Math.max(1, page - 1);
  }

  function nextPage(page, totalPages) {
    return Math.min(totalPages, page + 1);
  }

  function statusTone(value) {
    return value ? "text-[#0f2747]" : "text-slate-400";
  }

  async function openExpandedView(id, title, note = "") {
    if (id === "fastlonn-employee-preview-table") {
      fastlonnPopupColumnMode = "default";
    }
    if (id === "overtid-table") {
      overtidPopupColumnMode = "default";
    }
    if (id === "vakttillegg-table") {
      vakttilleggPopupColumnMode = "default";
    }
    if (id === "foreldrepermisjon-table") {
      foreldrepermisjonPopupColumnMode = "default";
    }
    focusBeforeDialog = document.activeElement;
    expandedView = { id, title, note };
    await tick();
    expandedCloseButton?.focus();
  }

  async function syncPopupScrollbars(type) {
    await tick();
    const bottomScroll =
      type === "fastlonn"
        ? fastlonnBottomScroll
        : type === "overtid"
          ? overtidBottomScroll
          : type === "vakttillegg"
            ? vakttilleggBottomScroll
            : foreldrepermisjonBottomScroll;
    if (!bottomScroll) return;

    const scrollWidth = bottomScroll.scrollWidth;
    const clientWidth = bottomScroll.clientWidth;
    const trackWidth = Math.max(scrollWidth, clientWidth);

    if (type === "fastlonn") {
      fastlonnTopScrollWidth = trackWidth;
    } else if (type === "overtid") {
      overtidTopScrollWidth = trackWidth;
    } else if (type === "vakttillegg") {
      vakttilleggTopScrollWidth = trackWidth;
    } else {
      foreldrepermisjonTopScrollWidth = trackWidth;
    }
  }

  function handleTopScrollbarScroll(type) {
    const topScroll =
      type === "fastlonn"
        ? fastlonnTopScroll
        : type === "overtid"
          ? overtidTopScroll
          : type === "vakttillegg"
            ? vakttilleggTopScroll
            : foreldrepermisjonTopScroll;
    const bottomScroll =
      type === "fastlonn"
        ? fastlonnBottomScroll
        : type === "overtid"
          ? overtidBottomScroll
          : type === "vakttillegg"
            ? vakttilleggBottomScroll
            : foreldrepermisjonBottomScroll;
    if (!topScroll || !bottomScroll) return;
    bottomScroll.scrollLeft = topScroll.scrollLeft;
  }

  function handleBottomScrollbarScroll(type) {
    const topScroll =
      type === "fastlonn"
        ? fastlonnTopScroll
        : type === "overtid"
          ? overtidTopScroll
          : type === "vakttillegg"
            ? vakttilleggTopScroll
            : foreldrepermisjonTopScroll;
    const bottomScroll =
      type === "fastlonn"
        ? fastlonnBottomScroll
        : type === "overtid"
          ? overtidBottomScroll
          : type === "vakttillegg"
            ? vakttilleggBottomScroll
            : foreldrepermisjonBottomScroll;
    if (!topScroll || !bottomScroll) return;
    topScroll.scrollLeft = bottomScroll.scrollLeft;
  }

  function closeExpandedView() {
    expandedView = null;
    tick().then(() => focusBeforeDialog?.focus?.());
  }

  function handleDialogKeydown(event) {
    if (!expandedView) return;
    if (event.key === "Escape") {
      event.preventDefault();
      closeExpandedView();
      return;
    }
    if (event.key !== "Tab" || !expandedDialog) return;
    const focusable = [...expandedDialog.querySelectorAll('button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])')];
    if (!focusable.length) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }

  function handleExpandKeydown(event, id, title, note = "") {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openExpandedView(id, title, note);
    }
  }

  function formatDataCell(value) {
    if (value === null || value === undefined || value === "") return "—";
    const parsedDate = typeof value === "number" ? null : parseCellDate(value);
    if (parsedDate) {
      return parsedDate.toLocaleDateString("nb-NO", { timeZone: "UTC" });
    }
    return String(value);
  }

  function containsNorwegianCharacters(value) {
    return norwegianCharacterPattern.test(String(value ?? ""));
  }

  function buildReadErrorMessage(files, fallbackMessage, caughtError) {
    return caughtError?.message || fallbackMessage;
  }

  async function applyFiles(files, sourceLabel) {
    const parsed = await readExcelFiles(files);
    resetReplacedAssignments(parsed);
    uploadedFiles = mergeFiles(uploadedFiles, parsed);
    dataSourceLabel = sourceLabel;
  }

  function stopFileSelectionWatcher(input) {
    const timer = fileSelectionWatchers.get(input);
    if (timer) window.clearTimeout(timer);
    fileSelectionWatchers.delete(input);
  }

  function watchFileSelection(input, role = null, attempt = 0) {
    stopFileSelectionWatcher(input);
    const timer = window.setTimeout(() => {
      if (input.files?.length) {
        if (role) {
          handleRoleUpload(role, { currentTarget: input });
        } else {
          handleUpload({ currentTarget: input });
        }
        return;
      }
      if (attempt < 39) {
        watchFileSelection(input, role, attempt + 1);
        return;
      }

      fileSelectionWatchers.delete(input);
    }, 250);
    fileSelectionWatchers.set(input, timer);
  }

  async function handleUpload(event) {
    if (loading || exporting) return;
    const input = event.currentTarget;
    const incoming = consumeSelectedFiles(input);
    if (!incoming.length) return;
    stopFileSelectionWatcher(input);

    loading = true;
    error = "";

    try {
      await applyFiles(
        incoming,
        uploadedFiles.length ? "Kombinert datasett med opplastede filer" : "Manuelt opplastede filer"
      );
    } catch (caughtError) {
      error = buildReadErrorMessage(incoming, "Kunne ikke lese de opplastede filene.", caughtError);
    } finally {
      loading = false;
    }
  }

  async function handleRoleUpload(role, event) {
    if (loading || exporting) return;
    const input = event.currentTarget;
    const incoming = consumeSelectedFiles(input);
    if (!incoming.length) return;
    stopFileSelectionWatcher(input);

    loading = true;
    error = "";

    try {
      const parsed = await readExcelFiles(incoming, role);
      resetReplacedAssignments(parsed);
      const invalidFiles = parsed.filter((file) => file.role !== role);

      if (invalidFiles.length) {
        throw new Error(`Feil filtype for ${uploadCardConfig.find((item) => item.role === role)?.title?.toLowerCase() || role}.`);
      }

      const hadExistingFiles = uploadedFiles.length > 0;
      uploadedFiles = mergeFiles(uploadedFiles, parsed);
      dataSourceLabel = hadExistingFiles ? "Kombinert datasett med opplastede filer" : "Manuelt opplastede filer";
    } catch (caughtError) {
      error = buildReadErrorMessage(incoming, "Kunne ikke lese den opplastede filen.", caughtError);
    } finally {
      loading = false;
    }
  }

  function clearUploadedFiles() {
    if (loading || exporting) return;
    groupAssignmentCache = {};
    expandedView = null;
    controlTableRows = {};
    controlTablePages = {};
    uploadedFiles = [];
    report = null;
    error = "";
    dataSourceLabel = "Ingen filer lastet opp";
    selectedFastlonnYear = "";
    selectedFastlonnPeriodKey = "12-31";
  }

  function selectFastlonnPeriod(periodKey) {
    selectedFastlonnPeriodKey = periodKey;
  }

  function updateManualReportInput(field, value) {
    manualReportInputs = {
      ...manualReportInputs,
      [field]: value
    };
    persistManualReportInputs();
  }

  function clearManualReportInputs() {
    manualReportInputs = createEmptyManualReportInputs();
    persistManualReportInputs();
  }

  function persistManualReportInputs() {
    if (typeof window === "undefined") return;
    manualReportInputCache = {
      ...manualReportInputCache,
      [activeManualReportInputKey]: manualReportInputs
    };
    try {
      window.localStorage.setItem(manualReportStorageKey, JSON.stringify(manualReportInputCache));
    } catch {
      error = "Kunne ikke lagre de manuelle feltene lokalt. Datafilene lagres aldri i nettleseren.";
    }
  }

  function toggleSidebar() {
    sidebarCollapsed = !sidebarCollapsed;
    try { window.localStorage.setItem("hr-arsrapport-sidebar-collapsed", String(sidebarCollapsed)); } catch {}
  }

  function toggleReadinessPanel() {
    readinessExpanded = !readinessExpanded;
    try { window.localStorage.setItem(readinessExpandedStorageKey, String(readinessExpanded)); } catch {}
  }

  onMount(() => {
    mounted = true;
    try {
    sidebarCollapsed = window.localStorage.getItem("hr-arsrapport-sidebar-collapsed") === "true";
    const savedReadinessState = window.localStorage.getItem(readinessExpandedStorageKey);
    readinessExpanded = savedReadinessState === null ? false : savedReadinessState === "true";
      manualReportInputCache = JSON.parse(window.localStorage.getItem(manualReportStorageKey) || "{}");
    } catch {
      manualReportInputCache = {};
      try { window.localStorage.removeItem(manualReportStorageKey); } catch {}
    }
    manualReportInputs = normalizeManualReportInputs(manualReportInputCache[activeManualReportInputKey]);
    loadedManualReportInputKey = activeManualReportInputKey;
    window.addEventListener("keydown", handleDialogKeydown);
    return () => {
      window.removeEventListener("keydown", handleDialogKeydown);
      for (const timer of fileSelectionWatchers.values()) clearTimeout(timer);
      fileSelectionWatchers.clear();
    };
  });
</script>

<svelte:head>
  <title>HR Årsrapport</title>
</svelte:head>

<section class="app-shell">
  <div class:sidebar-collapsed={sidebarCollapsed} class="app-frame" inert={Boolean(expandedView)}>
    <aside class="app-sidebar">
      <div class="app-brand">
        <button
          class="sidebar-toggle"
          type="button"
          aria-label={sidebarCollapsed ? "Åpne meny" : "Skjul tekst i meny"}
          aria-pressed={sidebarCollapsed}
          on:click={toggleSidebar}
        >
          <svg class:sidebar-toggle-icon-collapsed={sidebarCollapsed} class="sidebar-toggle-icon" viewBox="0 0 24 24" aria-hidden="true">
            <path d="M5.5 5.5v13" fill="none" stroke="currentColor" stroke-linecap="round" stroke-width="1.8" />
            <path d="m15 7-4.5 5 4.5 5" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8" />
          </svg>
        </button>
        <p class:visually-hidden={sidebarCollapsed} class="app-brand-title">HR Årsrapport</p>
      </div>
      <div class="workspace-tablist" role="group" aria-label="Arbeidsflater">
        <button class={`workspace-tab ${activeTab === "opplasting" ? "active" : ""}`} aria-pressed={activeTab === "opplasting"} aria-label="Datagrunnlag" on:click={() => (activeTab = "opplasting")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8.5A2.5 2.5 0 0 1 6.5 6H10l2 2h5.5A2.5 2.5 0 0 1 20 10.5v7A2.5 2.5 0 0 1 17.5 20h-11A2.5 2.5 0 0 1 4 17.5Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M12 16v-5m0 0-2 2m2-2 2 2" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Datagrunnlag</span></button>
        <button class={`workspace-tab ${activeTab === "fastlonn" ? "active" : ""}`} aria-pressed={activeTab === "fastlonn"} aria-label="Fastlønn" on:click={() => (activeTab = "fastlonn")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8.5A2.5 2.5 0 0 1 6.5 6h11A2.5 2.5 0 0 1 20 8.5v7A2.5 2.5 0 0 1 17.5 18h-11A2.5 2.5 0 0 1 4 15.5Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M15.5 12h.01M8 12h3.5m-1.75-1.75v3.5" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Fastlønn</span></button>
        <button class={`workspace-tab ${activeTab === "overtid" ? "active" : ""}`} aria-pressed={activeTab === "overtid"} aria-label="Overtid" on:click={() => (activeTab = "overtid")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 7v5l3 2.5M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Overtid</span></button>
        <button class={`workspace-tab ${activeTab === "vakttillegg" ? "active" : ""}`} aria-pressed={activeTab === "vakttillegg"} aria-label="Vakttillegg" on:click={() => (activeTab = "vakttillegg")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M15.5 4.5A7.5 7.5 0 1 0 19 18a7 7 0 1 1-3.5-13.5Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M16.5 9.5h.01" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Vakttillegg</span></button>
        <button class={`workspace-tab ${activeTab === "foreldrepermisjon" ? "active" : ""}`} aria-pressed={activeTab === "foreldrepermisjon"} aria-label="Foreldrepermisjon" on:click={() => (activeTab = "foreldrepermisjon")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 10.5a2.5 2.5 0 1 1 5 0c0 1.3-.6 2.3-2.5 4-1.9-1.7-2.5-2.7-2.5-4Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M5 19a4.5 4.5 0 0 1 4.5-4.5h4A4.5 4.5 0 0 1 18 19M8 8a2 2 0 1 1-4 0 2 2 0 0 1 4 0Zm12 0a2 2 0 1 1-4 0 2 2 0 0 1 4 0Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Foreldrepermisjon</span></button>
        <button class={`workspace-tab ${activeTab === "arsrapport" ? "active" : ""}`} aria-pressed={activeTab === "arsrapport"} aria-label="Årsrapport" on:click={() => (activeTab = "arsrapport")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3h7l5 5v13H7zM14 3v5h5M9 13h6M9 17h6" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Årsrapport</span></button>
      </div>
    </aside>

    <main class="app-content">
      <header class="app-hero">
        <p class="app-eyebrow">{activeTabLabel}</p>
        <h1 class="app-title">HR Årsrapport</h1>
      </header>

      <section class="app-context-strip">
        {#each globalContextItems as item}
          <article class="context-pill">
            <p class="context-pill-label">{item.label}</p>
            <p class="context-pill-value">{item.value}</p>
          </article>
        {/each}
      </section>

      {#if uploadedFiles.length}
        <section class:readiness-panel-collapsed={!readinessExpanded} class="readiness-panel" aria-labelledby="readiness-title">
          <div class="readiness-header">
            <div>
              <p class="panel-eyebrow">Kontroll før rapport</p>
              <h2 id="readiness-title" class="readiness-title">Dette må følges opp</h2>
              {#if readinessExpanded}
                <p class="readiness-lead">Listen oppdateres automatisk når du laster opp nye filer eller fyller inn manuelle tall.</p>
              {/if}
            </div>
            <div class="readiness-header-actions">
              <div class="readiness-counts" aria-label="Status for gjenstående arbeid">
                <span class={`readiness-count ${datasetReadinessIssues.length ? "readiness-count-error" : "readiness-count-ok"}`}>
                  {datasetReadinessIssues.length} dataavvik
                </span>
                <span class={`readiness-count ${incompleteManualInputCount ? "readiness-count-warning" : "readiness-count-ok"}`}>
                  {incompleteManualInputCount} manuelle tall mangler
                </span>
              </div>
              <button
                class="readiness-toggle"
                type="button"
                aria-expanded={readinessExpanded}
                aria-controls="readiness-details"
                on:click={toggleReadinessPanel}
              >
                <span>{readinessExpanded ? "Skjul" : "Vis detaljer"}</span>
                <svg class:readiness-toggle-icon-collapsed={!readinessExpanded} class="readiness-toggle-icon" viewBox="0 0 20 20" aria-hidden="true">
                  <path d="m5.5 12.5 4.5-5 4.5 5" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8" />
                </svg>
              </button>
            </div>
          </div>

          {#if readinessExpanded}
            <div id="readiness-details">
              <div class="readiness-grid">
                <section class="readiness-section" aria-labelledby="dataset-issues-title">
              <div class="readiness-section-head">
                <div>
                  <p class="readiness-kicker">Datagrunnlag</p>
                  <h3 id="dataset-issues-title">Må rettes eller avklares</h3>
                </div>
                <button class="readiness-link" type="button" on:click={() => (activeTab = "opplasting")}>Gå til filer</button>
              </div>

              {#if datasetReadinessIssues.length}
                <ol class="readiness-issue-list">
                  {#each datasetReadinessIssues as issue}
                    <li class="readiness-issue">
                      <div class="readiness-issue-marker" aria-hidden="true">!</div>
                      <div>
                        <p class="readiness-item-title">{issue.title}</p>
                        <p class="readiness-item-detail">{issue.detail}</p>
                        {#if issue.action}<p class="readiness-item-action"><span>Gjør dette:</span> {issue.action}</p>{/if}
                        {#if issue.people?.length}
                          <ul class="readiness-people" aria-label={`Personer som ikke er matchet for ${issue.title}`}>
                            {#each issue.people as person}
                              <li><span>{person.name}</span><strong>{formatCurrency(person.amount)}</strong></li>
                            {/each}
                          </ul>
                        {/if}
                      </div>
                    </li>
                  {/each}
                </ol>
              {:else}
                <div class="readiness-empty">
                  <span aria-hidden="true">✓</span>
                  <p>Ingen kjente dataavvik for valgt år og uttrekk.</p>
                </div>
              {/if}
                </section>

                <section class="readiness-section" aria-labelledby="manual-work-title">
              <div class="readiness-section-head">
                <div>
                  <p class="readiness-kicker">Manuelt arbeid</p>
                  <h3 id="manual-work-title">Tall, tekst og vurderinger</h3>
                </div>
                <button class="readiness-link" type="button" on:click={() => (activeTab = "arsrapport")}>Gå til manuelle felt</button>
              </div>

              <ul class="manual-task-list">
                {#each manualInputTasks as task}
                  <li class:manual-task-complete={task.complete} class="manual-task">
                    <span class="manual-task-status" aria-hidden="true">{task.complete ? "✓" : "•"}</span>
                    <div>
                      <p class="readiness-item-title">{task.label}</p>
                      <p class="readiness-item-detail">{task.complete ? "Fylt ut og klar for eksport." : task.description}</p>
                    </div>
                    <span class="manual-task-badge">{task.complete ? "Fylt ut" : "Mangler"}</span>
                  </li>
                {/each}
              </ul>

              <div class="readiness-assessments">
                <p class="readiness-subtitle">Må vurderes og beskrives i rapporten</p>
                <ul>
                  {#each manualAssessmentTasks as task}
                    <li>{task}</li>
                  {/each}
                </ul>
              </div>
                </section>
              </div>

              <p class="readiness-footer">Rapportutkastet kan lastes ned underveis, men avvikene bør være avklart før publisering.</p>
            </div>
          {/if}
        </section>
      {/if}

      {#if error}
        <div role="alert" class="error-box mb-3 px-3 py-2 text-xs">{error}</div>
      {/if}

      {#if loading}
        <div role="status" class="loading-box mb-3 p-4 text-xs">Leser filer og bygger oversikt...</div>
      {/if}

      {#if activeTab === "opplasting"}
        <div class="space-y-3">
        <div class="content-card">
          <div class="panel-actions">
            <div>
              <p class="panel-eyebrow">Datagrunnlag</p>
              <h2 class="panel-title">Last opp lønnsgrunnlaget</h2>
              <p class="section-note mt-2">Filene behandles bare i nettleseren på denne maskinen. De sendes ikke til en server og lagres ikke mellom økter.</p>
            </div>
            <div class="batch-upload-control">
              <label for="batch-upload" class="upload-control-label">Last opp alle fire samlet</label>
              <input id="batch-upload" class="upload-file-input upload-file-input-batch" type="file" disabled={!mounted || loading || exporting} on:cancel={(event) => stopFileSelectionWatcher(event.currentTarget)} multiple accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" on:pointerdown={(event) => watchFileSelection(event.currentTarget)} on:input={handleUpload} on:change={handleUpload} />
            </div>
          </div>

          <div class="upload-grid gap-3">
            {#each uploadCardConfig as card}
              <article class={`upload-card min-h-[196px] p-4 ${roleUploads[card.role] ? "upload-card-ready" : "upload-card-pending"}`}>
                <div class="upload-card-head">
                  <p class="upload-card-label text-[13px]">{card.title}</p>
                  <span class={`upload-card-state text-xs ${roleUploads[card.role] ? "upload-card-state-ready" : "upload-card-state-pending"}`}>
                    {roleUploads[card.role] ? "Lastet inn" : "Mangler"}
                  </span>
                </div>
                <div class="upload-card-copy mt-2">
                  <p class="upload-card-title text-[15px] leading-6">{card.description}</p>
                </div>
                <div class="upload-card-footer mt-3 gap-2">
                  {#if roleUploads[card.role]}
                    <span class="upload-selected-file">{roleUploads[card.role].fileName}</span>
                  {/if}
                  <input class="upload-file-input" type="file" disabled={!mounted || loading || exporting} on:cancel={(event) => stopFileSelectionWatcher(event.currentTarget)} aria-label={`Velg Excel-fil for ${card.title}`} accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" on:pointerdown={(event) => watchFileSelection(event.currentTarget, card.role)} on:input={(event) => handleRoleUpload(card.role, event)} on:change={(event) => handleRoleUpload(card.role, event)} />
                  <span class="upload-card-hint text-xs">XLSX · maks 15 MB</span>
                </div>
              </article>
            {/each}
          </div>

          <GroupAssignments employees={classificationEmployees} onChange={setGroupAssignment} disabled={!mounted || loading || exporting} />

          <div class="upload-lower">
            <div class="upload-date-area p-4">
              <div class="upload-date-head">
                <div>
                  <p class="upload-card-label text-[13px]">Lønnsdato</p>
                  <p class="upload-date-title text-[15px]">Velg dato for uttrekket.</p>
                </div>
                <span class={`upload-card-state text-xs ${fastlonnSnapshots.length ? "upload-card-state-ready" : "upload-card-state-pending"}`}>
                  {fastlonnSnapshots.length ? "Klar" : "Venter"}
                </span>
              </div>
              {#if fastlonnSnapshots.length}
                <div class="snapshot-controls">
                  <label class="report-manual-field">
                    <span class="report-manual-label">Rapportår</span>
                    <select class="report-manual-input" bind:value={selectedFastlonnYear} disabled={!mounted || loading || exporting}>
                      {#each availableFastlonnYears as year}<option value={year}>{year}</option>{/each}
                    </select>
                  </label>
                  <div class="snapshot-periods" role="group" aria-label="Velg lønnsdato">
                    {#each fastlonnPeriodChoices as period}
                      <button
                        type="button"
                        class={`snapshot-period ${selectedFastlonnPeriodKey === period.key ? "snapshot-period-active" : ""}`}
                        disabled={!period.available || loading || exporting}
                        on:click={() => selectFastlonnPeriod(period.key)}
                      >
                        {period.label}
                      </button>
                    {/each}
                  </div>
                </div>
                {#if selectedFastlonnYear && !selectedFastlonnPeriodChoice?.available}
                  <p class="upload-date-note text-[13px] leading-5">Mangler fastlønnfil for {selectedFastlonnPeriodKey === "05-01" ? "01.05" : "31.12"} i {selectedFastlonnYear}.</p>
                {/if}
                {#if selectedFastlonnPeriodChoice?.available}
                  <p class="upload-date-note text-[13px] leading-5">
                    Bruker lønnsdata{selectedFastlonnYear ? ` for ${selectedFastlonnYear}` : ""} og filtrerer ansatte per {selectedFastlonnPeriodChoice.label}.
                  </p>
                {/if}
              {:else}
                <p class="upload-date-note text-[13px] leading-5">Ingen lønnsdato tilgjengelig ennå.</p>
              {/if}
            </div>

            <div class="upload-meta p-4">
              <button class="app-button app-button-danger upload-meta-action" type="button" on:click={clearUploadedFiles} disabled={!uploadedFiles.length || loading || exporting}>
                Fjern opplastede filer
              </button>

              <div class="upload-coverage">
                {#each coverage as item}
                  <span class={`coverage-chip ${item.present ? "coverage-chip-ok" : "coverage-chip-missing"}`}>
                    {item.role === "fastlønn"
                      ? "Fastlønn"
                      : item.role === "overtid"
                        ? "Overtid"
                        : item.role === "vakttillegg"
                          ? "Vakttillegg"
                          : item.role === "foreldrepermisjon"
                            ? "Foreldrepermisjon"
                            : item.role}
                  </span>
                {/each}
              </div>
            </div>
          </div>

        </div>
        {#if uploadedFiles.length}
          <details class="app-panel">
            <summary class="cursor-pointer list-none px-4 py-3 text-sm font-medium text-[#0f5368]">Vis filer og kontrollgrunnlag</summary>
            <div class="border-t border-slate-200 p-4 pt-3">
              <p class="section-note">Viser filmetadata og kontrollgrunnlag i en strammere rapporttabell med tydelige kolonner og enkel paginering.</p>
              <article class="summary-card mt-4">
                <div class="flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <p class="summary-card-label">Filoversikt</p>
                    <p class="mt-1 text-sm font-medium text-slate-900">Alle opplastede filer i én samlet oversikt</p>
                  </div>
                  <div class="table-toolbar">
                    <p class="table-meta-text">
                      {#if uploadedFilesOverview.length}
                        Viser {fileOverviewPageData.startRow}-{fileOverviewPageData.endRow} av {uploadedFilesOverview.length}
                      {:else}
                        Ingen filer
                      {/if}
                    </p>
                    <label class="table-toolbar-control">
                      <span>Rader</span>
                      <select
                        class="table-select"
                        value={fileOverviewRows}
                        on:change={(event) => {
                          fileOverviewRows = Number(event.currentTarget.value);
                          fileOverviewPage = 1;
                        }}
                      >
                        {#each fileOverviewPageSizes as size}
                          <option value={size}>{size}</option>
                        {/each}
                      </select>
                    </label>
                  </div>
                </div>
                <div
                  class="ledger-table-wrap mt-3 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne filoversikt i større visning"
                  on:click={() => openExpandedView("upload-files-overview", "Filoversikt")}
                  on:keydown={(event) => handleExpandKeydown(event, "upload-files-overview", "Filoversikt")}
                >
                  <table class="ledger-table">
                    <thead>
                      <tr>
                        <th>Fil</th>
                        <th>Type</th>
                        <th>Kilde</th>
                        <th>Lønnsdato</th>
                        <th>Filår</th>
                        <th class="text-right">Rader</th>
                        <th class="text-right">Kolonner</th>
                        <th>Kolonner i fil</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each fileOverviewPageData.rows as item}
                        <tr>
                          <td class="font-medium text-slate-900">{item.Fil}</td>
                          <td>{item.Type}</td>
                          <td>{item.Kilde}</td>
                          <td>{item["Lønnsdato"]}</td>
                          <td>{item["Filår"]}</td>
                          <td class="text-right tabular-nums">{formatInteger(item.Rader)}</td>
                          <td class="text-right tabular-nums">{formatInteger(item.Kolonner)}</td>
                          <td class="text-slate-600">{item["Kolonner i fil"]}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                </div>
                <div class="table-pager">
                  <button
                    class="table-pager-button"
                    type="button"
                    disabled={fileOverviewPageData.currentPage <= 1}
                    on:click={() => (fileOverviewPage = previousPage(fileOverviewPageData.currentPage))}
                  >
                    Forrige
                  </button>
                  <p class="table-meta-text">Side {fileOverviewPageData.currentPage} av {fileOverviewPageData.totalPages}</p>
                  <button
                    class="table-pager-button"
                    type="button"
                    disabled={fileOverviewPageData.currentPage >= fileOverviewPageData.totalPages}
                    on:click={() => (fileOverviewPage = nextPage(fileOverviewPageData.currentPage, fileOverviewPageData.totalPages))}
                  >
                    Neste
                  </button>
                </div>
              </article>

              <div class="mt-4 space-y-3">
                {#each uploadedFiles as file}
                  <article class="summary-card">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                      <div>
                        <p class="summary-card-label">{file.label}</p>
                        <p class="mt-1 text-sm font-medium text-slate-900">{file.fileName}</p>
                      </div>
                      <div class="table-toolbar">
                        <p class="table-meta-text">
                          {#if file.rows?.length}
                            Viser {getControlTableView(file).startRow}-{getControlTableView(file).endRow} av {formatInteger(file.rowCount)} rader
                          {:else}
                            0 rader
                          {/if}
                        </p>
                        <label class="table-toolbar-control">
                          <span>Rader</span>
                          <select
                            class="table-select"
                            value={getControlTableRows(file)}
                            on:change={(event) => setControlTableRows(file, event.currentTarget.value)}
                          >
                            {#each controlTablePageSizes as size}
                              <option value={size}>{size}</option>
                            {/each}
                          </select>
                        </label>
                      </div>
                    </div>

                    {#if file.rows?.length}
                      <div
                        class="ledger-table-wrap mt-3 cursor-zoom-in"
                        role="button"
                        tabindex="0"
                        aria-label={`Åpne ${file.label} i større visning`}
                        on:click={() => openExpandedView(`upload-file-${fileIdentity(file)}`, file.fileName)}
                        on:keydown={(event) => handleExpandKeydown(event, `upload-file-${fileIdentity(file)}`, file.fileName)}
                      >
                        <table class="ledger-table ledger-table-dense">
                          <thead>
                            <tr>
                              <th class="w-14 text-right">#</th>
                              {#each fileColumns(file) as column}
                                <th>{column}</th>
                              {/each}
                            </tr>
                          </thead>
                          <tbody>
                            {#each getControlTableView(file).rows as row, rowIndex}
                              <tr>
                                <td class="text-right tabular-nums text-slate-400">{getControlTableView(file).startRow + rowIndex}</td>
                                {#each fileColumns(file) as column}
                                  <td>{String(row[column] ?? "")}</td>
                                {/each}
                              </tr>
                            {/each}
                          </tbody>
                        </table>
                      </div>
                      <div class="table-pager">
                        <button
                          class="table-pager-button"
                          type="button"
                          disabled={getControlTableView(file).currentPage <= 1}
                          on:click={() => setControlTablePage(file, previousPage(getControlTableView(file).currentPage))}
                        >
                          Forrige
                        </button>
                        <p class="table-meta-text">Side {getControlTableView(file).currentPage} av {getControlTableView(file).totalPages}</p>
                        <button
                          class="table-pager-button"
                          type="button"
                          disabled={getControlTableView(file).currentPage >= getControlTableView(file).totalPages}
                          on:click={() => setControlTablePage(file, nextPage(getControlTableView(file).currentPage, getControlTableView(file).totalPages))}
                        >
                          Neste
                        </button>
                      </div>
                    {:else}
                      <p class="mt-3 text-xs text-slate-500">Ingen data tilgjengelig for denne filen.</p>
                    {/if}
                  </article>
                {/each}
              </div>
            </div>
          </details>
        {/if}
        </div>
      {/if}

      {#if activeTab === "fastlonn"}
        <section class="report-page">
          <div class="page-intro"><h2>Fastlønn og kjønnsbalanse</h2><p>Rapportgrunnlag per {report?.fastlonnSnapshotLabel || 'valgt lønnsdato'}.</p></div>
          {#if report?.fastlonn}
            <article class="report-section">
              <h3>Kjønnsbalanse per stillingsgruppe</h3>
              <ReportTable columns={genderColumns} rows={genderTableRows} caption="Kjønnsbalanse per stillingsgruppe" />
            </article>
            <article class="report-section">
              <h3>Gjennomsnittlig årslønn</h3>
              <p class="report-description">Oppgitt årslønn i kroner. Kolonnen «Kvinner i % av menn» viser forholdet mellom gjennomsnittslønnene, slik rapportmalen ber om.</p>
              <ReportTable columns={salaryColumns} rows={moneyTableRows(salaryComparisonRows)} caption="Gjennomsnittlig årslønn per stillingsgruppe" />
              <p class="report-footnote">Tankestrek betyr at beregningsgrunnlaget mangler. Gruppevalg og avgrensninger kontrolleres under Datagrunnlag.</p>
            </article>
            <article class="report-section">
              <h3>Midlertidige ansatte og deltid</h3>
              <ReportTable columns={employmentColumns} rows={employmentTableRows} caption="Midlertidige ansatte og faktisk deltid" />
              <p class="report-footnote">Andelene beregnes innen hvert kjønn. Deltid betyr stillingsprosent under 100; frivillig og ufrivillig deltid må vurderes manuelt.</p>
            </article>
            <div class="report-charts">
              <article class="report-section">
                <h3>Kjønnsbalanse</h3>
                <BarChart data={genderBalanceChartData.filter(row => row.gruppe !== 'Totalt antall ansatte')} x="gruppe" y="antall" series="kjønn" seriesOrder={["Kvinner", "Menn"]} seriesColors={genderSeriesColors} yFmt="num0" xAxisTitle="Stillingsgruppe" yAxisTitle="Antall ansatte" />
              </article>
              <article class="report-section">
                <h3>Gjennomsnittlig årslønn</h3>
                <BarChart data={salaryComparisonChartData.filter(row => row.gruppe !== 'Totalt antall ansatte')} x="gruppe" y="verdi" series="serie" seriesOrder={["Kvinner", "Menn"]} seriesColors={paySeriesColors} yFmt="#,##0" xAxisTitle="Stillingsgruppe" yAxisTitle="Kroner" />
              </article>
            </div>
            <details class="source-details"><summary>Kontroller ansatte i uttrekket</summary>
              <p>Viser inntil 20 ansatte. Alle rådata er tilgjengelige under Datagrunnlag.</p>
              <ReportTable columns={[{key:'name',label:'Ansatt'},{key:'group',label:'Rapportgruppe'},{key:'gender',label:'Kjønn'},{key:'salary',label:'Årslønn',numeric:true}]} rows={fastlonnEmployeePreview.map(row => ({...row,salary:formatCurrency(row.salary)}))} caption="Ansatte i valgt fastlønnsuttrekk" />
            </details>
          {:else}<p class="content-card">Last opp fastlønn under Datagrunnlag for å vise rapporttabellene.</p>{/if}
        </section>
      {/if}

      {#if activeTab === "overtid" || activeTab === "vakttillegg"}
        {@const isOvertime = activeTab === 'overtid'}
        {@const title = isOvertime ? 'Overtid' : 'Vakttillegg'}
        {@const rows = isOvertime ? overtidAnalysisRows : vakttilleggAnalysisRows}
        {@const chartData = isOvertime ? overtidChartData : vakttilleggChartData}
        {@const participants = isOvertime ? overtidParticipants : vakttilleggParticipants}
        {@const reconciliation = isOvertime ? overtidReconciliation : vakttilleggReconciliation}
        <section class="report-page">
          <div class="page-intro"><h2>{title}</h2><p>Utbetalinger i {selectedFastlonnYear || 'valgt år'}, koblet til fastlønnsuttrekket per {report?.fastlonnSnapshotLabel || 'valgt lønnsdato'}.</p></div>
          {#if rows.length}
            <article class="report-section">
              <h3>Gjennomsnittlig {title.toLowerCase()} per mottaker</h3>
              <p class="report-description">Årsbeløp i kroner blant mottakere med positiv samlet utbetaling som kan kobles til valgt lønnsuttrekk.</p>
              <ReportTable columns={salaryColumns} rows={moneyTableRows(rows)} caption={`${title} per stillingsgruppe`} />
              <p class="report-footnote">Antall er antall mottakere i hver gruppe. Personer uten utbetaling inngår ikke i gjennomsnittet.</p>
            </article>
            <article class="report-section">
              <h3>Gjennomsnitt fordelt på kvinner og menn</h3>
              <div role="button" tabindex="0" class="expandable-chart" aria-label={`Åpne ${title.toLowerCase()} per gruppe i større visning`} on:click={() => openExpandedView(isOvertime ? 'overtid-chart' : 'vakttillegg-chart', `${title} per gruppe`)} on:keydown={(event) => handleExpandKeydown(event, isOvertime ? 'overtid-chart' : 'vakttillegg-chart', `${title} per gruppe`)}>
                <BarChart data={chartData.filter(row => !row.gruppe.startsWith('Mottakere i valgt uttrekk'))} x="gruppe" y="verdi" series="serie" seriesOrder={["Kvinner", "Menn"]} seriesColors={paySeriesColors} yFmt="#,##0" xAxisTitle="Stillingsgruppe" yAxisTitle="Kroner" />
              </div>
            </article>
            <details class="source-details"><summary>Kontroller kobling og utbetalinger</summary>
              <p>{formatInteger(reconciliation?.matchedCount)} av {formatInteger(reconciliation?.sourceCount)} mottakere er koblet til lønnsuttrekket. Ukoblet beløp: {formatCurrency(reconciliation?.unmatchedAmount)}. Se merknadene under Datagrunnlag.</p>
              <div role="button" tabindex="0" aria-label={`Åpne datagrunnlag for ${title.toLowerCase()} i større visning`} on:click={() => openExpandedView(isOvertime ? 'overtid-table' : 'vakttillegg-table', `Datagrunnlag for ${title.toLowerCase()}`)} on:keydown={(event) => handleExpandKeydown(event, isOvertime ? 'overtid-table' : 'vakttillegg-table', `Datagrunnlag for ${title.toLowerCase()}`)}>
                <ReportTable columns={[{key:'name',label:'Ansatt'},{key:'group',label:'Rapportgruppe'},{key:'gender',label:'Kjønn'},{key:'amount',label:'Årsbeløp',numeric:true}]} rows={participants.slice(0,20).map(row => ({...row,amount:formatCurrency(row.amount)}))} caption={`Kontrollutvalg: ${title.toLowerCase()} (inntil 20 mottakere)`} />
              </div>
            </details>
          {:else}<p class="content-card">Last opp fastlønn og {title.toLowerCase()} under Datagrunnlag for å vise rapporttabellen.</p>{/if}
        </section>
      {/if}

      {#if activeTab === "foreldrepermisjon"}
        <section class="report-page">
          <div class="page-intro"><h2>Foreldrepermisjon</h2><p>Permisjonsuttak i {selectedFastlonnYear || 'valgt år'}, fordelt på kvinner og menn.</p></div>
          {#if report?.foreldrepermisjon}
            <article class="report-section">
              <h3>Permisjonsuttak</h3>
              <ReportTable columns={leaveColumns} rows={leaveTableRows} caption="Foreldrepermisjon fordelt på kjønn" />
              <p class="report-footnote">Gjennomsnittet gjelder ansatte med uttak. En uke tilsvarer fem arbeidsdager. Andel permisjonsdager beregnes av samlet uttak.</p>
            </article>
            <article class="report-section">
              <h3>Gjennomsnittlig uttak i uker</h3>
              <BarChart data={foreldrepermisjonWeeksChartData} x="måling" y="verdi" series="kjønn" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="#,##0.0" xAxisTitle="Måling" yAxisTitle="Uker" />
            </article>
            <details class="source-details"><summary>Kontroller permisjonsperiodene</summary>
              <p>Arbeidsdager i kalenderåret, uten helger og norske helligdager. Overlappende perioder telles én gang; dagene vektes ikke med permisjonsgrad.</p>
              <ReportTable columns={[{key:'name',label:'Ansatt'},{key:'gender',label:'Kjønn'},{key:'days',label:'Arbeidsdager',numeric:true},{key:'weeks',label:'Uker',numeric:true}]} rows={foreldrepermisjonEmployees.map(row => ({...row,days:formatInteger(row.days),weeks:formatNumber(row.weeks)}))} caption="Permisjonsuttak per ansatt" />
            </details>
          {:else}<p class="content-card">Last opp foreldrepermisjon under Datagrunnlag for å vise rapporttabellen.</p>{/if}
        </section>
      {/if}

      {#if activeTab === "arsrapport"}
        <section class="space-y-3">
          <div class="content-card">
            <div class="panel-actions">
              <div>
                <p class="panel-eyebrow">Årsrapport</p>
                <h2 class="panel-title">Last ned redigerbar årsrapport</h2>
                <p class="section-note mt-2">Bruk denne fanen til å hente ut rapportutkastet når datagrunnlaget er kontrollert.</p>
              </div>
              <button class="app-button app-button-download" on:click={downloadReport} disabled={!report?.fastlonn || loading || exporting || hasInvalidManualInputs}>{exporting ? "Lager Word-rapport …" : "Last ned rapportutkast (DOCX)"}</button>
            </div>
          </div>
          <div class="content-card">
            <p class="panel-eyebrow">Forhåndsutfylling</p>
            <h2 class="panel-title">Manuelle felt før eksport</h2>
            {#if hasInvalidManualInputs}<p role="alert" class="error-box p-3 mt-2">Antall nyansatte må være et heltall som er 0 eller større. Sykefravær må være en prosent mellom 0 og 100. Rett de markerte feltene før eksport.</p>{/if}
            <p class="section-note mt-2">
              Verdiene under lagres lokalt i nettleseren for {selectedFastlonnYear || "gjeldende utvalg"} og settes rett inn i DOCX-filen ved eksport.
            </p>
            {#if !selectedFastlonnYear}<p class="section-note mt-2">Last opp fastlønn og velg rapportår før du fyller inn feltene.</p>{/if}

            <div class="report-manual-grid mt-4">
              <label class="report-manual-field">
                <span class="report-manual-label">Antall nyansatte</span>
                <input class="report-manual-input" disabled={!selectedFastlonnYear || exporting} type="text" inputmode="numeric" aria-invalid={Boolean(manualReportInputs.nyansatteCount && validateManualField("nyansatteCount", manualReportInputs.nyansatteCount))} maxlength="20" value={manualReportInputs.nyansatteCount} on:input={(event) => updateManualReportInput("nyansatteCount", event.currentTarget.value)} />
              </label>
              <label class="report-manual-field">
                <span class="report-manual-label">Samlet sykefravær %</span>
                <input class="report-manual-input" disabled={!selectedFastlonnYear || exporting} type="text" inputmode="decimal" aria-invalid={Boolean(manualReportInputs.samletSykefravaer && validateManualField("samletSykefravaer", manualReportInputs.samletSykefravaer))} maxlength="20" value={manualReportInputs.samletSykefravaer} on:input={(event) => updateManualReportInput("samletSykefravaer", event.currentTarget.value)} />
              </label>
              <label class="report-manual-field">
                <span class="report-manual-label">Legemeldt sykefravær kvinner %</span>
                <input class="report-manual-input" disabled={!selectedFastlonnYear || exporting} type="text" inputmode="decimal" aria-invalid={Boolean(manualReportInputs.sykefravaerKvinner && validateManualField("sykefravaerKvinner", manualReportInputs.sykefravaerKvinner))} maxlength="20" value={manualReportInputs.sykefravaerKvinner} on:input={(event) => updateManualReportInput("sykefravaerKvinner", event.currentTarget.value)} />
              </label>
              <label class="report-manual-field">
                <span class="report-manual-label">Legemeldt sykefravær menn %</span>
                <input class="report-manual-input" disabled={!selectedFastlonnYear || exporting} type="text" inputmode="decimal" aria-invalid={Boolean(manualReportInputs.sykefravaerMenn && validateManualField("sykefravaerMenn", manualReportInputs.sykefravaerMenn))} maxlength="20" value={manualReportInputs.sykefravaerMenn} on:input={(event) => updateManualReportInput("sykefravaerMenn", event.currentTarget.value)} />
              </label>
            </div>
            <button class="app-button mt-3" type="button" on:click={clearManualReportInputs} disabled={!selectedFastlonnYear || exporting}>Slett lagrede rapportfelt for {selectedFastlonnYear || 'valgt år'}</button>
          </div>
          <div class="content-card report-delivery">
            <h2 class="panel-title">Dette inngår i Word-utkastet</h2>
            <p class="section-note mt-2">Tabeller for kjønnsbalanse, fastlønn, overtid, vakttillegg, midlertidige ansatte, deltid og foreldrepermisjon, samt de manuelle tallene over.</p>
            <p class="section-note mt-2">Historikk for rekruttering og inkludering, samarbeid med utdanningsinstitusjoner, ARP og faglige vurderinger ferdigstilles i Word. Tidligere års vurderinger gjenbrukes ikke automatisk.</p>
          </div>
        </section>
      {/if}
    </main>
  </div>

  {#if expandedView}
    <div class="fixed inset-0 z-[80] flex items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="expanded-dialog-title">
      <button class="absolute inset-0 bg-slate-950/60" type="button" aria-label="Lukk popup" on:click={closeExpandedView}></button>
      <div bind:this={expandedDialog} class="relative z-10 max-h-[92vh] w-[min(96vw,1400px)] overflow-y-auto rounded-2xl border border-slate-200 bg-white shadow-2xl">
        <div class="sticky top-0 z-10 flex items-start justify-between gap-4 border-b border-slate-200 bg-white px-5 py-4">
          <div>
            <p class="text-[11px] font-medium uppercase tracking-[0.18em] text-slate-500">Større visning</p>
            <h3 id="expanded-dialog-title" class="mt-1 text-lg font-semibold text-[#0f2747]">{expandedView.title}</h3>
            {#if expandedView.note}
              <p class="mt-1 text-sm text-slate-600">{expandedView.note}</p>
            {/if}
          </div>
          <button bind:this={expandedCloseButton} class="rounded-lg border border-slate-200 px-3 py-1.5 text-sm font-medium text-slate-700 hover:bg-slate-50" type="button" on:click={closeExpandedView}>
            Lukk
          </button>
        </div>

        <div class="p-5">
          {#if expandedView.id === "upload-files-overview"}
            <div class="ledger-table-wrap">
              <table class="ledger-table">
                <thead>
                  <tr>
                    <th>Fil</th>
                    <th>Type</th>
                    <th>Kilde</th>
                    <th>Lønnsdato</th>
                    <th>Filår</th>
                    <th class="text-right">Rader</th>
                    <th class="text-right">Kolonner</th>
                    <th>Kolonner i fil</th>
                  </tr>
                </thead>
                <tbody>
                  {#each uploadedFilesOverview as item}
                    <tr>
                      <td class="font-medium text-slate-900">{item.Fil}</td>
                      <td>{item.Type}</td>
                      <td>{item.Kilde}</td>
                      <td>{item["Lønnsdato"]}</td>
                      <td>{item["Filår"]}</td>
                      <td class="text-right tabular-nums">{formatInteger(item.Rader)}</td>
                      <td class="text-right tabular-nums">{formatInteger(item.Kolonner)}</td>
                      <td class="text-slate-600">{item["Kolonner i fil"]}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id?.startsWith("upload-file-")}
            {@const expandedFile = uploadedFiles.find((file) => `upload-file-${fileIdentity(file)}` === expandedView.id)}
            {#if expandedFile}
              <div class="ledger-table-wrap">
                <table class="ledger-table ledger-table-dense">
                  <thead>
                    <tr>
                      <th class="w-14 text-right">#</th>
                      {#each fileColumns(expandedFile) as column}
                        <th>{column}</th>
                      {/each}
                    </tr>
                  </thead>
                  <tbody>
                    {#each expandedFile.rows as row, rowIndex}
                      <tr>
                        <td class="text-right tabular-nums text-slate-400">{rowIndex + 1}</td>
                        {#each fileColumns(expandedFile) as column}
                          <td>{String(row[column] ?? "")}</td>
                        {/each}
                      </tr>
                    {/each}
                  </tbody>
                </table>
              </div>
            {/if}
          {:else if expandedView.id === "overtid-table"}
            <div class="space-y-4">
              <div class="snapshot-controls">
                <div class="snapshot-periods">
                  <button
                    type="button"
                    class={`snapshot-period ${overtidPopupColumnMode === "default" ? "snapshot-period-active" : ""}`}
                    on:click={() => (overtidPopupColumnMode = "default")}
                  >
                    Standardvisning
                  </button>
                  <button
                    type="button"
                    class={`snapshot-period ${overtidPopupColumnMode === "all" ? "snapshot-period-active" : ""}`}
                    on:click={() => (overtidPopupColumnMode = "all")}
                  >
                    Alle kolonner
                  </button>
                </div>
              </div>
              {#if overtidPopupColumnMode === "all"}
                <div class="popup-top-scroll" bind:this={overtidTopScroll} on:scroll={() => handleTopScrollbarScroll("overtid")}>
                  <div style={`width: ${overtidTopScrollWidth}px; height: 1px;`}></div>
                </div>
              {/if}
              <div class="overflow-x-auto max-w-full" bind:this={overtidBottomScroll} on:scroll={() => handleBottomScrollbarScroll("overtid")}>
                {#if overtidPopupColumnMode === "all"}
                  <table class="min-w-max border-collapse text-left text-sm whitespace-nowrap">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        {#each overtidSourceHeaders as header}
                          <th class="border px-3 py-2 font-semibold">{header}</th>
                        {/each}
                      </tr>
                    </thead>
                    <tbody>
                      {#each overtidSourceRows as row}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          {#each overtidSourceHeaders as header}
                            <td class="border px-3 py-2 align-top">{formatDataCell(row[header])}</td>
                          {/each}
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {:else}
                  <table class="min-w-full border-collapse text-left text-sm">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        <th class="border px-3 py-2 font-semibold">Ansatt</th>
                        <th class="border px-3 py-2 font-semibold">Gruppe</th>
                        <th class="border px-3 py-2 font-semibold">Kjønn</th>
                        <th class="border px-3 py-2 font-semibold">Sum overtid</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each overtidParticipants as employee}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          <td class="border px-3 py-2 align-top">{employee.name}</td>
                          <td class="border px-3 py-2 align-top">{employee.group}</td>
                          <td class="border px-3 py-2 align-top">{employee.gender}</td>
                          <td class="border px-3 py-2 align-top">{formatCurrency(employee.amount)}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {/if}
              </div>
            </div>
          {:else if expandedView.id === "overtid-chart"}
            <BarChart data={overtidChartData} x="gruppe" y="verdi" series="serie" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={paySeriesColors} yFmt="#,##0" xAxisTitle="Gruppe" yAxisTitle="Kroner" swapXY={true} />
          {:else if expandedView.id === "vakttillegg-table"}
            <div class="space-y-4">
              <div class="snapshot-controls">
                <div class="snapshot-periods">
                  <button
                    type="button"
                    class={`snapshot-period ${vakttilleggPopupColumnMode === "default" ? "snapshot-period-active" : ""}`}
                    on:click={() => (vakttilleggPopupColumnMode = "default")}
                  >
                    Standardvisning
                  </button>
                  <button
                    type="button"
                    class={`snapshot-period ${vakttilleggPopupColumnMode === "all" ? "snapshot-period-active" : ""}`}
                    on:click={() => (vakttilleggPopupColumnMode = "all")}
                  >
                    Alle kolonner
                  </button>
                </div>
              </div>
              {#if vakttilleggPopupColumnMode === "all"}
                <div class="popup-top-scroll" bind:this={vakttilleggTopScroll} on:scroll={() => handleTopScrollbarScroll("vakttillegg")}>
                  <div style={`width: ${vakttilleggTopScrollWidth}px; height: 1px;`}></div>
                </div>
              {/if}
              <div class="overflow-x-auto max-w-full" bind:this={vakttilleggBottomScroll} on:scroll={() => handleBottomScrollbarScroll("vakttillegg")}>
                {#if vakttilleggPopupColumnMode === "all"}
                  <table class="min-w-max border-collapse text-left text-sm whitespace-nowrap">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        {#each vakttilleggSourceHeaders as header}
                          <th class="border px-3 py-2 font-semibold">{header}</th>
                        {/each}
                      </tr>
                    </thead>
                    <tbody>
                      {#each vakttilleggSourceRows as row}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          {#each vakttilleggSourceHeaders as header}
                            <td class="border px-3 py-2 align-top">{formatDataCell(row[header])}</td>
                          {/each}
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {:else}
                  <table class="min-w-full border-collapse text-left text-sm">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        <th class="border px-3 py-2 font-semibold">Ansatt</th>
                        <th class="border px-3 py-2 font-semibold">Gruppe</th>
                        <th class="border px-3 py-2 font-semibold">Kjønn</th>
                        <th class="border px-3 py-2 font-semibold">Sum vakttillegg</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each vakttilleggParticipants as employee}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          <td class="border px-3 py-2 align-top">{employee.name}</td>
                          <td class="border px-3 py-2 align-top">{employee.group}</td>
                          <td class="border px-3 py-2 align-top">{employee.gender}</td>
                          <td class="border px-3 py-2 align-top">{formatCurrency(employee.amount)}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {/if}
              </div>
            </div>
          {:else if expandedView.id === "vakttillegg-chart"}
            <BarChart data={vakttilleggChartData} x="gruppe" y="verdi" series="serie" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={paySeriesColors} yFmt="#,##0" xAxisTitle="Gruppe" yAxisTitle="Kroner" swapXY={true} />
          {/if}
        </div>
      </div>
    </div>
  {/if}
</section>
