<script>
  import { BarChart, BoxPlot, Heatmap, LineChart, ScatterPlot } from "@evidence-dev/core-components";
  import { onMount, tick } from "svelte";
  import { base } from "$app/paths";
  import * as XLSX from "xlsx";
  import {
    AlignmentType,
    BorderStyle,
    Document,
    Packer,
    Paragraph,
    ShadingType,
    Table,
    TableCell,
    TableRow,
    TextRun,
    WidthType,
    HeadingLevel
  } from "docx";

  const requiredRoles = ["fastlønn", "overtid", "vakttillegg", "foreldrepermisjon"];
  const directorGroupUnderdirectors = new Set(["Meltevik|Stine", "Thune|Rita Lund"]);
  const defaultDataFiles = [
    "Utdrag fra SAP foreldrepermisjoner 2025 - Raadata.xlsx",
    "Utdrag fra SAP ordinaer fastlonn per 311225 - Raadata.xlsx",
    "Utdrag fra SAP overtid i 2025 - Raadata.xlsx",
    "Utdrag fra SAP vakttillegg i 2025 - Raadata.xlsx"
  ];
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
  const groupOrder = [
    "Totalt antall ansatte",
    "Direktørgruppen",
    "Seksjonssjefgruppen",
    "Fagsjefgruppen",
    "Seniorpersonale",
    "Øvrige saksbehandlere og andre"
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
  const groupChartColors = {
    "Totalt antall ansatte": graphColors.darkPurple,
    Direktørgruppen: graphColors.orange,
    Seksjonssjefgruppen: graphColors.lightPurple,
    Fagsjefgruppen: graphColors.green,
    Seniorpersonale: graphColors.pink,
    "Øvrige saksbehandlere og andre": graphColors.lightPurpleUu,
    "Snitt alle ansatte": graphColors.darkPurple
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
  let dataSourceLabel = "Standardfiler fra default_data";
  let expandedView = null;
  let selectedFastlonnYear = "";
  let selectedFastlonnPeriodKey = "12-31";
  let activeTab = "opplasting";
  let sidebarCollapsed = false;
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

  const numberFormatter = new Intl.NumberFormat("nb-NO", {
    maximumFractionDigits: 1
  });
  const integerFormatter = new Intl.NumberFormat("nb-NO", {
    maximumFractionDigits: 0
  });
  const percentFormatter = new Intl.NumberFormat("nb-NO", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  const percentValueFormatter = new Intl.NumberFormat("nb-NO", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });

  $: coverage = requiredRoles.map((role) => ({
    role,
    present: uploadedFiles.some((file) => file.role === role && file.source === "rådata")
  }));
  $: roleUploads = Object.fromEntries(
    requiredRoles.map((role) => [role, uploadedFiles.find((file) => file.role === role && file.source === "rådata") ?? null])
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
      available: Boolean(
        selectedFastlonnYearSnapshots.find((file) => {
          const snapshotDate = parseSnapshotKeyDate(file.snapshotKey);
          const targetDate = periodDateForYear(selectedFastlonnYear, period.key);
          return snapshotDate && targetDate && snapshotDate.getTime() >= targetDate.getTime();
        })
      )
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
    manualReportInputs = normalizeManualReportInputs(manualReportInputCache[activeManualReportInputKey]);
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
  $: if (uploadedFiles.length) {
    report = buildReport(uploadedFiles, selectedFastlonnYear, selectedFastlonnPeriodKey);
  }
  $: fastlonnSelection = getFastlonnFilesForSelection(uploadedFiles, selectedFastlonnYear, selectedFastlonnPeriodKey);
  $: fastlonnSourceFiles = fastlonnSelection.fastlonnFilesForPeriod;
  $: fastlonnSourceRows = buildFastlonnRows(fastlonnSourceFiles);
  $: fastlonnSourceDateColumns = detectEmploymentDateColumns(fastlonnSourceRows);
  $: fastlonnPreviewRows = fastlonnSourceRows.filter((row) =>
    isActiveOnDate(row, fastlonnSelection.targetSnapshotDate, fastlonnSourceDateColumns)
  );
  $: fastlonnSourceHeaders = (() => {
    const fileHeaders = [...new Set(fastlonnSourceFiles.flatMap((file) => file.header || []))]
      .map((header) => String(header).trim())
      .filter((header) => header && !header.startsWith("__"));
    if (fileHeaders.length) return fileHeaders;
    return Object.keys(fastlonnPreviewRows[0] || {}).filter((header) => !header.startsWith("__"));
  })();
  $: overtidSourceFiles = uploadedFiles.filter(
    (file) => file.role === "overtid" && file.source === "rådata" && (!selectedFastlonnYear || file.fileYear === selectedFastlonnYear)
  );
  $: vakttilleggSourceFiles = uploadedFiles.filter(
    (file) => file.role === "vakttillegg" && file.source === "rådata" && (!selectedFastlonnYear || file.fileYear === selectedFastlonnYear)
  );
  $: foreldrepermisjonSourceFiles = uploadedFiles.filter(
    (file) => file.role === "foreldrepermisjon" && file.source === "rådata" && (!selectedFastlonnYear || file.fileYear === selectedFastlonnYear)
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
  $: foreldrepermisjonSourceHeaders = (() => {
    const fileHeaders = [...new Set(foreldrepermisjonSourceFiles.flatMap((file) => file.header || []))]
      .map((header) => String(header).trim())
      .filter((header) => header && !header.startsWith("__"));
    if (fileHeaders.length) return fileHeaders;
    return Object.keys(foreldrepermisjonSourceRows[0] || {}).filter((header) => !header.startsWith("__"));
  })();
  $: overtidParticipants = report?.fastlonn ? buildVariablePayParticipants(overtidSourceRows, report.fastlonn.employeeIndex) : [];
  $: vakttilleggParticipants = report?.fastlonn ? buildVariablePayParticipants(vakttilleggSourceRows, report.fastlonn.employeeIndex) : [];
  $: foreldrepermisjonEmployees = buildParentalLeaveEmployees(foreldrepermisjonSourceRows);
  $: foreldrepermisjonGroupedEmployees = report?.fastlonn ? buildParentalLeaveGroupedEmployees(foreldrepermisjonSourceRows, report.fastlonn.employeeIndex) : [];
  $: fastlonnEmployeePreview = report?.fastlonn?.employees
    ? [...report.fastlonn.employees]
        .sort((left, right) => left.group.localeCompare(right.group, "nb") || left.name.localeCompare(right.name, "nb"))
        .slice(0, 20)
    : [];
  $: salarySpreadRows = report?.fastlonn
    ? groupOrder
        .map((groupName) => {
          const salaries = report.fastlonn.employees
            .filter((employee) =>
              groupName === "Totalt antall ansatte" ? employee.salary > 0 : employee.group === groupName && employee.salary > 0
            )
            .map((employee) => employee.salary)
            .sort((a, b) => a - b);

          if (!salaries.length) return null;

          return {
            gruppe: groupName,
            min: salaries[0],
            q1: quantileSorted(salaries, 0.25),
            median: quantileSorted(salaries, 0.5),
            q3: quantileSorted(salaries, 0.75),
            max: salaries[salaries.length - 1],
            spread: salaries[salaries.length - 1] - salaries[0]
          };
        })
        .filter(Boolean)
    : [];
  $: genderBalanceRows = report?.fastlonn?.genderBalance ?? [];
  $: salaryComparisonRows = report?.fastlonn?.fastlonn ?? [];
  $: fastlonnRepresentationTableRows = genderBalanceRows.map((row) => ({
    gruppe: row.group,
    kvinner: row.women,
    menn: row.men,
    kvinneandel: formatPercent(row.womenShare),
    mannandel: formatPercent(row.menShare),
    totalt: row.total
  }));
  $: fastlonnSalaryTableRows = salaryComparisonRows.map((row) => ({
    gruppe: row.group,
    ansatte: row.n,
    kvinner: formatCurrency(row.womenAvg),
    menn: formatCurrency(row.menAvg),
    totalt: formatCurrency(row.totalAvg),
    lønnsforhold: row.menAvg > 0 ? formatPercent(row.womenPctOfMen) : "—"
  }));
  $: fastlonnSpreadTableRows = salarySpreadRows.map((row) => ({
    gruppe: row.gruppe,
    min: formatCurrency(row.min),
    q1: formatCurrency(row.q1),
    median: formatCurrency(row.median),
    q3: formatCurrency(row.q3),
    maks: formatCurrency(row.max)
  }));
  $: fastlonnEmploymentTableRows = employmentAnalysisRows.map((row) => ({
    kategori: row.kategori,
    kvinner: `${formatInteger(row.kvinnerAntall)} (${formatPercent(row.kvinnerAndel)})`,
    menn: `${formatInteger(row.mennAntall)} (${formatPercent(row.mennAndel)})`
  }));
  $: fastlonnSourceSummaryRows = fastlonnSourceFiles.map((file) => ({
    fil: file.fileName,
    snapshot: file.snapshotLabel || "—",
    år: file.fileYear || "—",
    rader: formatInteger(file.rowCount)
  }));
  $: genderBalanceChartData = genderBalanceRows.flatMap((row) => [
    { gruppe: row.group, kjønn: "Kvinner", antall: row.women },
    { gruppe: row.group, kjønn: "Menn", antall: row.men }
  ]);
  $: genderBalanceHeatmapData = genderBalanceRows.flatMap((row) => [
    { gruppe: row.group, kjønn: "Kvinner", andel: row.womenShare / 100 },
    { gruppe: row.group, kjønn: "Menn", andel: row.menShare / 100 }
  ]);
  $: salaryComparisonChartData = salaryComparisonRows.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenAvg },
    { gruppe: row.group, serie: "Menn", verdi: row.menAvg }
  ]);
  $: salaryScatterData = salaryComparisonRows
    .filter((row) => row.group !== "Totalt antall ansatte" && row.n > 0 && row.womenAvg > 0 && row.menAvg > 0)
    .map((row) => ({
      gruppe: row.group,
      kvinner: row.womenAvg,
      menn: row.menAvg,
      antall: row.n
    }));
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
  $: overtidGroupMetrics = report?.fastlonn ? buildVariablePayGroupMetrics(overtidParticipants, report.fastlonn.employees) : [];
  $: overtidParticipationChartData = overtidGroupMetrics.map((row) => ({
    gruppe: row.group,
    andel: row.participationShare / 100
  }));
  $: overtidParticipationGenderChartData = overtidGroupMetrics.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenParticipationShare },
    { gruppe: row.group, serie: "Menn", verdi: row.menParticipationShare }
  ]);
  $: overtidGenderShareChartData = overtidGroupMetrics.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenParticipants },
    { gruppe: row.group, serie: "Menn", verdi: row.menParticipants }
  ]);
  $: overtidTotalChartData = overtidGroupMetrics.map((row) => ({
    gruppe: row.group,
    verdi: row.totalAmount,
    verdi_tusen: row.totalAmount / 1000
  }));
  $: overtidTopGroupsChartData = [...overtidGroupMetrics]
    .filter((row) => row.group !== "Snitt alle ansatte" && row.totalAmount > 0)
    .sort((left, right) => right.totalAmount - left.totalAmount)
    .slice(0, 5)
    .map((row) => ({
      gruppe: row.group,
      verdi: row.totalAmount
    }));
  $: overtidSpreadRows = buildVariablePaySpreadRows(overtidParticipants);
  $: overtidMonthlyRows = report?.fastlonn ? buildVariablePayMonthlyRows(overtidSourceRows, report.fastlonn.employeeIndex) : [];
  $: overtidMonthlyChartData = overtidMonthlyRows.flatMap((row) => [
    { måned: row.måned, serie: "Kvinner", verdi: row.kvinner },
    { måned: row.måned, serie: "Menn", verdi: row.menn }
  ]);
  $: overtidMonthlyTotalChartData = overtidMonthlyRows.map((row) => ({
    måned: row.måned,
    verdi: row.totalt
  }));
  $: overtidHeatmapData = report?.fastlonn ? buildVariablePayMonthHeatmap(overtidSourceRows, report.fastlonn.employeeIndex) : [];
  $: overtidLoadScatterData = overtidGroupMetrics
    .filter((row) => row.group !== "Snitt alle ansatte" && row.participants > 0 && row.totalAmount > 0)
    .map((row) => ({
      gruppe: row.group,
      ansatte: row.participants,
      total: row.totalAmount,
      snitt: row.totalAvgPerParticipant
    }));
  $: vakttilleggChartData = vakttilleggAnalysisRows.flatMap((row) => [
    { gruppe: row.group, serie: "Kvinner", verdi: row.womenAvg },
    { gruppe: row.group, serie: "Menn", verdi: row.menAvg }
  ]);
  $: vakttilleggGroupMetrics = report?.fastlonn ? buildVariablePayGroupMetrics(vakttilleggParticipants, report.fastlonn.employees) : [];
  $: vakttilleggParticipationChartData = vakttilleggGroupMetrics.map((row) => ({
    gruppe: row.group,
    andel: row.participationShare / 100
  }));
  $: vakttilleggTotalChartData = vakttilleggGroupMetrics.map((row) => ({
    gruppe: row.group,
    verdi: row.totalAmount
  }));
  $: vakttilleggSpreadRows = buildVariablePaySpreadRows(vakttilleggParticipants);
  $: vakttilleggMonthlyRows = report?.fastlonn ? buildVariablePayMonthlyRows(vakttilleggSourceRows, report.fastlonn.employeeIndex) : [];
  $: vakttilleggMonthlyTotalChartData = vakttilleggMonthlyRows.map((row) => ({
    måned: row.måned,
    verdi: row.totalt
  }));
  $: salarySpreadSeriesColors = Object.fromEntries(salarySpreadRows.map((row) => [row.gruppe, groupChartColors[row.gruppe] || graphColors.black]));
  $: salaryScatterSeriesColors = Object.fromEntries(salaryScatterData.map((row) => [row.gruppe, groupChartColors[row.gruppe] || graphColors.black]));
  $: vakttilleggSpreadSeriesColors = Object.fromEntries(
    vakttilleggSpreadRows.map((row) => [
      row.gruppe,
      row.gruppe === "Ledere med personalansvar" ? graphColors.orange : groupChartColors[row.gruppe] || graphColors.black
    ])
  );
  $: overtidTotalChartPalette = overtidTotalChartData.map((row) => groupChartColors[row.gruppe] || graphColors.black);
  $: overtidParticipationChartPalette = overtidParticipationChartData.map((row) => groupChartColors[row.gruppe] || graphColors.black);
  $: vakttilleggTotalChartPalette = vakttilleggTotalChartData.map((row) => groupChartColors[row.gruppe] || graphColors.black);
  $: vakttilleggParticipationChartPalette = vakttilleggParticipationChartData.map((row) => groupChartColors[row.gruppe] || graphColors.black);
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
          totalt: 100
        }
      ]
    : [];
  $: foreldrepermisjonWeeksChartData = report?.foreldrepermisjon
    ? [
        { måling: "Gjennomsnittlige uker", kjønn: "Kvinner", verdi: report.foreldrepermisjon.womenAvgWeeks },
        { måling: "Gjennomsnittlige uker", kjønn: "Menn", verdi: report.foreldrepermisjon.menAvgWeeks }
      ]
    : [];
  $: foreldrepermisjonCountChartData = report?.foreldrepermisjon
    ? [
        { måling: "Ansatte med uttak", kjønn: "Kvinner", verdi: report.foreldrepermisjon.womenCount },
        { måling: "Ansatte med uttak", kjønn: "Menn", verdi: report.foreldrepermisjon.menCount }
      ]
    : [];
  $: foreldrepermisjonGroupCountChartData = buildParentalLeaveGroupCountChartData(foreldrepermisjonGroupedEmployees);
  $: foreldrepermisjonGroupParticipationChartData = report?.fastlonn
    ? buildParentalLeaveGroupParticipationData(foreldrepermisjonGroupedEmployees, report.fastlonn.employees)
    : [];
  $: foreldrepermisjonGroupParticipationHeatmapData = foreldrepermisjonGroupParticipationChartData.map((row) => ({
    ...row,
    måling: "Andel"
  }));
  $: foreldrepermisjonSpreadRows = buildParentalLeaveSpreadRows(foreldrepermisjonGroupedEmployees);
  $: foreldrepermisjonShareChartData = report?.foreldrepermisjon
    ? [
        { måling: "Andel permisjonsdager", kjønn: "Kvinner", verdi: report.foreldrepermisjon.womenShareDays },
        { måling: "Andel permisjonsdager", kjønn: "Menn", verdi: report.foreldrepermisjon.menShareDays }
      ]
    : [];
  $: employmentHeatmapData = employmentAnalysisRows.flatMap((row) => [
    { måling: row.kategori, kjønn: "Kvinner", andel: row.kvinnerAndel / 100 },
    { måling: row.kategori, kjønn: "Menn", andel: row.mennAndel / 100 }
  ]);
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
  $: overtidSourceDistinctEmployees = countDistinctValues(overtidSourceRows.map((row) => row["Etternavn, fornavn"]), normalizePersonName);
  $: vakttilleggSourceDistinctEmployees = countDistinctValues(vakttilleggSourceRows.map((row) => row["Etternavn, fornavn"]), normalizePersonName);
  $: selectedParentalYearLabel = selectedFastlonnYear || "—";
  $: globalContextItems = [
    { label: "År", value: selectedFastlonnYear || "Ikke valgt" },
    { label: "Uttrekk", value: selectedFastlonnPeriodChoice?.label || "—" },
    { label: "Datakilde", value: dataSourceLabel },
    { label: "Rådatafiler", value: formatInteger(activeSourceSummaryRows.reduce((sum, row) => sum + row.filer, 0)) }
  ];
  $: overallGenderBalance = report?.fastlonn?.genderBalance?.find((row) => row.group === "Totalt antall ansatte") ?? null;
  $: salaryGapRow = report?.fastlonn?.fastlonn
    ?.filter((row) => row.group !== "Totalt antall ansatte" && row.n > 0 && row.womenAvg > 0 && row.menAvg > 0)
    ?.sort((a, b) => Math.abs(100 - b.womenPctOfMen) - Math.abs(100 - a.womenPctOfMen))[0] ?? null;
  $: representationSkewRow = report?.fastlonn?.genderBalance
    ?.filter((row) => row.group !== "Totalt antall ansatte" && row.total > 0)
    ?.sort((a, b) => Math.abs(b.womenShare - 50) - Math.abs(a.womenShare - 50))[0] ?? null;
  $: workPatternGap = report?.employment
    ? [
        {
          label: "Midlertidig ansatte",
          womenShare: report.employment.totalWomen ? (report.employment.temporary.women / report.employment.totalWomen) * 100 : 0,
          menShare: report.employment.totalMen ? (report.employment.temporary.men / report.employment.totalMen) * 100 : 0
        },
        {
          label: "Faktisk deltid",
          womenShare: report.employment.totalWomen ? (report.employment.partTime.women / report.employment.totalWomen) * 100 : 0,
          menShare: report.employment.totalMen ? (report.employment.partTime.men / report.employment.totalMen) * 100 : 0
        }
      ].sort((a, b) => Math.abs(b.womenShare - b.menShare) - Math.abs(a.womenShare - a.menShare))[0]
    : null;
  $: priorityCards = [
    overallGenderBalance
      ? {
          eyebrow: "Representasjon",
          metric: `${formatPercent(overallGenderBalance.womenShare)} kvinner`,
          title: "Totalt i virksomheten",
          takeaway: representationSkewRow
            ? `${representationSkewRow.group} er skjevest med ${formatPercent(representationSkewRow.womenShare)} kvinner og ${formatPercent(representationSkewRow.menShare)} menn.`
            : "Ingen tydelig skjev gruppe funnet.",
          tone: "fastlonn-kpi-card-representation"
        }
      : null,
    salaryGapRow
      ? {
          eyebrow: "Lønn",
          metric: `${formatPercent(salaryGapRow.womenPctOfMen)} av menns lønn`,
          title: salaryGapRow.group,
          takeaway: "Største gapet bør forklares med roller, ansvar, ansiennitet eller lønnspraksis.",
          tone: "fastlonn-kpi-card-salary"
        }
      : null,
    workPatternGap
      ? {
          eyebrow: "Arbeidsmønster",
          metric: `${formatPercent(Math.abs(workPatternGap.womenShare - workPatternGap.menShare))}poeng`,
          title: workPatternGap.label,
          takeaway: "Største forskjell mellom kvinner og menn i arbeidsmønster.",
          tone: "fastlonn-kpi-card-work-pattern"
        }
      : null,
    report?.foreldrepermisjon
      ? {
          eyebrow: "Foreldrepermisjon",
          metric: `${formatNumber(report.foreldrepermisjon.womenAvgWeeks)} mot ${formatNumber(report.foreldrepermisjon.menAvgWeeks)} uker`,
          title: "Snittuttak kvinner og menn",
          takeaway: `${formatPercent(report.foreldrepermisjon.womenShareDays)} av permisjonsdagene gjelder kvinner.`,
          tone: "fastlonn-kpi-card-parental-leave"
        }
      : null
  ].filter(Boolean);
  $: overviewInsights = [];
  $: if (report?.fastlonn?.genderBalance?.length) {
    const total = report.fastlonn.genderBalance.find((row) => row.group === "Totalt antall ansatte");
    const salaryGap = report.fastlonn.fastlonn
      .filter((row) => row.group !== "Totalt antall ansatte" && row.n > 0)
      .sort((a, b) => Math.abs(100 - b.womenPctOfMen) - Math.abs(100 - a.womenPctOfMen))[0];
    const nextInsights = [];

    if (total) {
      nextInsights.push(
        `${formatPercent(total.womenShare)} av de ansatte er kvinner, og ${formatPercent(total.menShare)} er menn.`
      );
    }

    if (salaryGap) {
      nextInsights.push(
        `Størst lønnsforskjell finnes i ${salaryGap.group}, der kvinners gjennomsnittslønn er ${formatPercent(salaryGap.womenPctOfMen)} av menns.`
      );
    }

    if (report?.employment) {
      const womenPartTime = report.employment.totalWomen
        ? (report.employment.partTime.women / report.employment.totalWomen) * 100
        : 0;
      const menPartTime = report.employment.totalMen
        ? (report.employment.partTime.men / report.employment.totalMen) * 100
        : 0;
      nextInsights.push(
        `Faktisk deltid er registrert for ${formatPercent(womenPartTime)} av kvinnene og ${formatPercent(menPartTime)} av mennene.`
      );
    }

    if (report?.foreldrepermisjon) {
      nextInsights.push(
        `Foreldrepermisjon utgjør i snitt ${formatNumber(report.foreldrepermisjon.womenAvgWeeks)} uker for kvinner og ${formatNumber(report.foreldrepermisjon.menAvgWeeks)} uker for menn.`
      );
    }

    overviewInsights = nextInsights;
  }
  $: overtidTopGroup = [...overtidGroupMetrics]
    .filter((row) => row.group !== "Snitt alle ansatte" && row.totalAmount > 0)
    .sort((left, right) => right.totalAmount - left.totalAmount)[0] ?? null;
  $: vakttilleggTopGroup = [...vakttilleggGroupMetrics]
    .filter((row) => row.group !== "Snitt alle ansatte" && row.totalAmount > 0)
    .sort((left, right) => right.totalAmount - left.totalAmount)[0] ?? null;
  $: parentalLeaveTopGroup = [...foreldrepermisjonGroupParticipationChartData]
    .sort((left, right) => right.andel - left.andel)[0] ?? null;
  $: dataQualityChecks = [
    {
      label: "Nødvendige rådatafiler",
      status: missingRoles.length ? "Mangler" : "Klar",
      detail: missingRoles.length ? `Mangler ${missingRoles.join(", ")}.` : "Alle fire rådataområder er lastet inn.",
      tone: missingRoles.length ? "warning" : "ok"
    },
    {
      label: "Fastlønn for valgt uttrekk",
      status: fastlonnSourceFiles.length ? "Klar" : "Svak",
      detail: fastlonnSourceFiles.length
        ? `Bruker ${formatInteger(fastlonnSourceFiles.length)} fastlønnfil(er) for ${selectedFastlonnPeriodChoice?.label || "valgt uttrekk"}.`
        : "Fant ingen fastlønnfil som dekker valgt år og uttrekksdato.",
      tone: fastlonnSourceFiles.length ? "ok" : "warning"
    },
    {
      label: "Match overtid mot fastlønn",
      status: !overtidSourceRows.length ? "Ingen data" : overtidParticipants.length === overtidSourceDistinctEmployees ? "Klar" : "Svak",
      detail: !overtidSourceRows.length
        ? "Ingen overtidsrader i valgt år."
        : `${formatInteger(overtidParticipants.length)} av ${formatInteger(overtidSourceDistinctEmployees)} ansatte med overtid er matchet mot fastlønn.`,
      tone: !overtidSourceRows.length ? "neutral" : overtidParticipants.length === overtidSourceDistinctEmployees ? "ok" : "warning"
    },
    {
      label: "Match vakttillegg mot fastlønn",
      status: !vakttilleggSourceRows.length ? "Ingen data" : vakttilleggParticipants.length === vakttilleggSourceDistinctEmployees ? "Klar" : "Svak",
      detail: !vakttilleggSourceRows.length
        ? "Ingen vakttilleggsrader i valgt år."
        : `${formatInteger(vakttilleggParticipants.length)} av ${formatInteger(vakttilleggSourceDistinctEmployees)} ansatte med vakttillegg er matchet mot fastlønn.`,
      tone: !vakttilleggSourceRows.length ? "neutral" : vakttilleggParticipants.length === vakttilleggSourceDistinctEmployees ? "ok" : "warning"
    },
    {
      label: "Foreldrepermisjon mot fastlønn",
      status: !foreldrepermisjonEmployees.length ? "Ingen data" : foreldrepermisjonGroupedEmployees.length === foreldrepermisjonEmployees.length ? "Klar" : "Svak",
      detail: !foreldrepermisjonEmployees.length
        ? "Ingen permisjonsrader i valgt år."
        : `${formatInteger(foreldrepermisjonGroupedEmployees.length)} av ${formatInteger(foreldrepermisjonEmployees.length)} ansatte med permisjonsuttak er koblet til fastlønn.`,
      tone: !foreldrepermisjonEmployees.length ? "neutral" : foreldrepermisjonGroupedEmployees.length === foreldrepermisjonEmployees.length ? "ok" : "warning"
    }
  ];
  $: reportHighlights = [
    overallGenderBalance
      ? {
          title: "Representasjon",
          metric: `${formatPercent(overallGenderBalance.womenShare)} kvinner totalt`,
          detail: representationSkewRow
            ? `${representationSkewRow.group} er mest skjevfordelt.`
            : "Ingen tydelig skjevhet mellom grupper."
        }
      : null,
    salaryGapRow
      ? {
          title: "Største lønnsgap",
          metric: salaryGapRow.group,
          detail: `Kvinners lønn er ${formatPercent(salaryGapRow.womenPctOfMen)} av menns i denne gruppen.`
        }
      : null,
    overtidTopGroup
      ? {
          title: "Overtid",
          metric: overtidTopGroup.group,
          detail: `${formatCurrency(overtidTopGroup.totalAmount)} i samlet overtid.`
        }
      : null,
    vakttilleggTopGroup
      ? {
          title: "Vakttillegg",
          metric: vakttilleggTopGroup.group,
          detail: `${formatCurrency(vakttilleggTopGroup.totalAmount)} i samlet vakttillegg.`
        }
      : null,
    parentalLeaveTopGroup
      ? {
          title: "Foreldrepermisjon",
          metric: parentalLeaveTopGroup.gruppe,
          detail: `${formatPercent(parentalLeaveTopGroup.andel * 100)} av gruppen har registrert uttak.`
        }
      : null,
    workPatternGap
      ? {
          title: "Arbeidsmønster",
          metric: workPatternGap.label,
          detail: `${formatPercent(Math.abs(workPatternGap.womenShare - workPatternGap.menShare))}poeng forskjell mellom kvinner og menn.`
        }
      : null
  ].filter(Boolean);
  $: reportActionItems = [
    salaryGapRow ? `Forklar lønnsgapet i ${salaryGapRow.group} med ansvar, rolle, ansiennitet eller lønnspraksis.` : null,
    overtidTopGroup ? `Beskriv hvorfor ${overtidTopGroup.group} har høyest samlet overtid og om dette er forventet drift.` : null,
    vakttilleggTopGroup ? `Forklar hvorfor ${vakttilleggTopGroup.group} bærer mest vakttillegg og hvilke beredskapsordninger som ligger bak.` : null,
    parentalLeaveTopGroup ? `Vurder om uttaket i ${parentalLeaveTopGroup.gruppe} krever forklaring knyttet til bemanning, alder eller organisering.` : null,
    dataQualityChecks.some((check) => check.tone === "warning")
      ? "Avklar datakvalitetsavvikene før rapportteksten ferdigstilles."
      : "Datagrunnlaget ser sammenhengende ut for valgt år og uttrekk."
  ].filter(Boolean);

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function formatNumber(value) {
    return numberFormatter.format(Number(value ?? 0));
  }

  function formatPercent(value) {
    return `${percentFormatter.format(Number(value ?? 0))} %`;
  }

  function formatPercentValue(value) {
    return percentValueFormatter.format(Number(value ?? 0));
  }

  function formatInteger(value) {
    return integerFormatter.format(Number(value ?? 0));
  }

  function formatCurrency(value) {
    return `${integerFormatter.format(Number(value ?? 0))} kr`;
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
    return Object.keys(file?.rows?.[0] || {});
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

  function openExpandedView(id, title, note = "") {
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
    expandedView = { id, title, note };
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
  }

  function handleExpandKeydown(event, id, title, note = "") {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openExpandedView(id, title, note);
    }
  }

  function mean(values) {
    if (!values.length) return 0;
    return values.reduce((sum, value) => sum + value, 0) / values.length;
  }

  function quantileSorted(values, quantile) {
    if (!values.length) return 0;
    const index = (values.length - 1) * quantile;
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    if (lower === upper) return values[lower];
    const weight = index - lower;
    return values[lower] * (1 - weight) + values[upper] * weight;
  }

  function normalizePersonName(value) {
    return normalizeText(value)
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function formatDataCell(value) {
    if (value === null || value === undefined || value === "") return "—";
    const parsedDate = parseCellDate(value);
    if (parsedDate) {
      return parsedDate.toLocaleDateString("nb-NO", { timeZone: "UTC" });
    }
    return String(value);
  }

  function parseCellDate(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number" && value > 20000 && value < 70000) {
      return new Date(Date.UTC(1899, 11, 30 + Number(value)));
    }
    const text = String(value).trim();
    if (!text) return null;
    const dottedMatch = text.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$/);
    if (dottedMatch) {
      const [, day, month, rawYear] = dottedMatch;
      const year = rawYear.length === 2 ? `20${rawYear}` : rawYear;
      return new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
    }
    const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (!isoMatch) return null;
    const isoDate = new Date(text);
    return Number.isNaN(isoDate.getTime()) ? null : isoDate;
  }

  function detectEmploymentDateColumns(rows) {
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

  function isActiveOnDate(record, snapshotDate, dateColumns) {
    if (!snapshotDate) return true;
    const startDate = dateColumns.startColumn ? parseCellDate(record[dateColumns.startColumn]) : null;
    const endDate = dateColumns.endColumn ? parseCellDate(record[dateColumns.endColumn]) : null;
    if (startDate && startDate > snapshotDate) return false;
    if (endDate && endDate < snapshotDate) return false;
    return true;
  }

  function parseSnapshotMeta(fileName) {
    const normalizedName = normalizeText(fileName);
    const compactName = normalizedName.replace(/\s+/g, "");
    const dottedMatch = normalizedName.match(/per(?:\s+den)?\s*(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?/);
    const compactMatch = compactName.match(/per(?:den)?(\d{2})(\d{2})(\d{2,4})/);
    const yearMatch = normalizedName.match(/\b(20\d{2})\b/);
    const fileYear = yearMatch?.[1] ?? "";

    if (!dottedMatch && !compactMatch) {
      return {
        snapshotKey: "udatert",
        snapshotLabel: "Udatert uttrekk",
        fileYear
      };
    }

    const [, rawDay, rawMonth, rawYearFromDotted] = dottedMatch || [];
    const [, compactDay, compactMonth, rawYearFromCompact] = compactMatch || [];
    const day = rawDay || compactDay;
    const month = rawMonth || compactMonth;
    const rawYear = rawYearFromDotted || rawYearFromCompact || "";
    const fullYear = rawYear.length === 2 ? `20${rawYear}` : rawYear;

    if (!day || !month || !fullYear) {
      return {
        snapshotKey: "udatert",
        snapshotLabel: "Udatert uttrekk",
        fileYear
      };
    }

    const paddedDay = String(day).padStart(2, "0");
    const paddedMonth = String(month).padStart(2, "0");

    return {
      snapshotKey: `${fullYear}-${paddedMonth}-${paddedDay}`,
      snapshotLabel: `${Number(day)}.${Number(month)}.${fullYear}`,
      fileYear: fullYear
    };
  }

  function detectFileMeta(fileName, header) {
    const lower = normalizeText(fileName);
    const normalizedHeader = header.map((item) => normalizeText(item));
    const isBearbeidet = lower.includes("bearbeidet");
    const snapshotMeta = parseSnapshotMeta(fileName);

    if (lower.includes("utregningsskjema")) {
      return { role: "utregningsskjema", source: "referanse", label: "Utregningsskjema" };
    }
    if (lower.includes("fastlonn") || lower.includes("fastlønn")) {
      return {
        role: "fastlønn",
        source: isBearbeidet ? "bearbeidet" : "rådata",
        label: "Fastlønn",
        snapshotKey: snapshotMeta.snapshotKey,
        snapshotLabel: snapshotMeta.snapshotLabel
      };
    }
    if (lower.includes("overtid")) {
      return { role: "overtid", source: isBearbeidet ? "bearbeidet" : "rådata", label: "Overtid" };
    }
    if (lower.includes("vakttillegg")) {
      return { role: "vakttillegg", source: isBearbeidet ? "bearbeidet" : "rådata", label: "Vakttillegg" };
    }
    if (lower.includes("foreldrepermis")) {
      return { role: "foreldrepermisjon", source: isBearbeidet ? "bearbeidet" : "rådata", label: "Foreldrepermisjon" };
    }
    if (
      normalizedHeader.includes("stillingsgruppe betegnelse") &&
      normalizedHeader.includes("nokkel for kjonn") &&
      normalizedHeader.includes("107a - individuell lonn arsbe")
    ) {
      return {
        role: "fastlønn",
        source: "rådata",
        label: "Fastlønn",
        snapshotKey: snapshotMeta.snapshotKey,
        snapshotLabel: snapshotMeta.snapshotLabel
      };
    }
    if (
      normalizedHeader.includes("frav.type") &&
      normalizedHeader.includes("frav.tekst") &&
      normalizedHeader.includes("frav.dager")
    ) {
      return { role: "foreldrepermisjon", source: isBearbeidet ? "bearbeidet" : "rådata", label: "Foreldrepermisjon" };
    }
    return { role: "ukjent", source: "ukjent", label: "Ukjent fil" };
  }

  function compareSnapshotKeys(left, right) {
    if (left === right) return 0;
    if (!left || left === "udatert") return 1;
    if (!right || right === "udatert") return -1;
    return left.localeCompare(right);
  }

  function parseSnapshotKeyDate(snapshotKey) {
    if (!snapshotKey || snapshotKey === "udatert") return null;
    const date = new Date(`${snapshotKey}T00:00:00Z`);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function periodDateForYear(year, periodKey) {
    if (!year || !periodKey) return null;
    return new Date(Date.UTC(Number(year), Number(periodKey.slice(0, 2)) - 1, Number(periodKey.slice(3, 5))));
  }

  function yearFromDate(date) {
    return date ? String(date.getUTCFullYear()) : "";
  }

  function deriveFileYear(role, rows, fallbackYear) {
    if (fallbackYear) return fallbackYear;

    if (role === "overtid" || role === "vakttillegg") {
      const payoutDate = rows.map((row) => parseCellDate(row["Utbetalingsdato"])).find(Boolean);
      return yearFromDate(payoutDate);
    }

    if (role === "foreldrepermisjon") {
      const startDate = rows.map((row) => parseCellDate(row["Start"])).find(Boolean);
      const endDate = rows.map((row) => parseCellDate(row["Slutt"])).find(Boolean);
      return yearFromDate(startDate || endDate);
    }

    return "";
  }

  function employeeGroup(record) {
    const title = String(record["Stillingsgruppe betegnelse"] || "");
    const [code] = title.split(" ");
    if (code === "9106" || code === "1060") return "Direktørgruppen";
    if (code === "1059") {
      const personKey = `${record["Etternavn"]}|${record["Fornavn"]}`;
      return directorGroupUnderdirectors.has(personKey) ? "Direktørgruppen" : "Fagsjefgruppen";
    }
    if (code === "1211") return "Seksjonssjefgruppen";
    if (["1220", "1538", "1072"].includes(code)) return "Fagsjefgruppen";
    if (["1088", "1181", "1114", "1364"].includes(code)) return "Seniorpersonale";
    if (["1087", "1434", "1408", "1065", "1124", "1184"].includes(code)) {
      return "Øvrige saksbehandlere og andre";
    }
    return null;
  }

  function annualSalary(record) {
    const individual = Number(record["107A - Individuell lønn årsbe"] || 0);
    const leader = Number(record["1006-Årslønn lederlønnstab."] || 0);
    return individual || leader;
  }

  function compareFastlonnCandidates(left, right, snapshotDate) {
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

  function buildEmployeeIndex(rows, snapshotDate) {
    const chosenByPerson = new Map();
    const dateColumns = detectEmploymentDateColumns(rows);
    const candidates = rows
      .map((record) => {
        const aliases = [
          `${record["Etternavn"] || ""} ${record["Fornavn"] || ""}`,
          `${record["Fornavn"] || ""} ${record["Etternavn"] || ""}`,
          record["Ansattnr - navn"]
        ]
          .map((alias) => normalizePersonName(alias))
          .filter(Boolean);
        const personKey = aliases[0] || "";
        const employee = {
          personKey,
          name: `${record["Etternavn"] || ""} ${record["Fornavn"] || ""}`.trim(),
          gender: String(record["Nøkkel for kjønn"] || record["Kjønn"] || ""),
          group: employeeGroup(record),
          salary: annualSalary(record),
          employmentGroup: String(record["Medarbeidergruppe"] || ""),
          partTimePercent: Number(record["Deltids-% 0007"] || 100)
        };
        return {
          employee,
          aliases,
          snapshotDate: parseSnapshotKeyDate(record.__snapshotKey || ""),
          record
        };
      })
      .filter(({ employee, record }) => employee.personKey && employee.group && employee.gender && isActiveOnDate(record, snapshotDate, dateColumns));

    for (const candidate of candidates) {
      const current = chosenByPerson.get(candidate.employee.personKey);
      if (!current || compareFastlonnCandidates(candidate, current, snapshotDate) < 0) {
        chosenByPerson.set(candidate.employee.personKey, candidate);
      }
    }

    const byName = new Map();
    const employees = [...chosenByPerson.values()].map((candidate) => {
      for (const alias of candidate.aliases) {
        byName.set(alias, candidate.employee);
      }
      return candidate.employee;
    });

    return { employees, byName };
  }

  function summarizeEmployment(employees) {
    const women = employees.filter((employee) => employee.gender === "Kvinne");
    const men = employees.filter((employee) => employee.gender === "Mann");
    const isTemporary = (employee) => ["Midlert./engasj. tjm", "Vikarer"].includes(employee.employmentGroup);
    const isPartTime = (employee) => employee.partTimePercent < 100;

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

  function variablePayGroup(group) {
    if (group === "Direktørgruppen" || group === "Seksjonssjefgruppen") {
      return "Ledere med personalansvar";
    }
    return group;
  }

  const variablePayOrder = [
    "Snitt alle ansatte",
    "Ledere med personalansvar",
    "Fagsjefgruppen",
    "Seniorpersonale",
    "Øvrige saksbehandlere og andre"
  ];
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Des"];

  function averageAmount(records) {
    return records.length ? records.reduce((sum, record) => sum + record.amount, 0) / records.length : 0;
  }

  function amountSum(records) {
    return records.reduce((sum, record) => sum + (Number(record.amount) || 0), 0);
  }

  function monthKeyFromDate(date) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  function formatMonthLabel(monthKey) {
    const [year, month] = String(monthKey).split("-");
    const monthIndex = Number(month) - 1;
    return `${monthNames[monthIndex] || month} ${year}`;
  }

  function summarizeVariablePay(rows, employeeIndex) {
    const totalsByPerson = new Map();

    for (const row of rows) {
      if (!row["Etternavn, fornavn"]) continue;

      const employee = employeeIndex.byName.get(normalizePersonName(row["Etternavn, fornavn"]));
      if (!employee?.group) continue;

      const key = normalizePersonName(row["Etternavn, fornavn"]);
      const current = totalsByPerson.get(key) || {
        ...employee,
        amount: 0
      };

      current.amount += Number(row["Beløp"] || 0);
      totalsByPerson.set(key, current);
    }

    const participants = [...totalsByPerson.values()].filter((employee) => employee.amount > 0);
    const order = [
      "Snitt alle ansatte",
      "Ledere med personalansvar",
      "Fagsjefgruppen",
      "Seniorpersonale",
      "Øvrige saksbehandlere og andre"
    ];

    return order.map((groupName) => {
      const subset =
        groupName === "Snitt alle ansatte"
          ? participants
          : participants.filter((employee) => variablePayGroup(employee.group) === groupName);
      const women = subset.filter((employee) => employee.gender === "Kvinne");
      const men = subset.filter((employee) => employee.gender === "Mann");
      const womenAvg = averageAmount(women);
      const menAvg = averageAmount(men);

      return {
        group: groupName,
        n: subset.length,
        womenAvg,
        menAvg,
        totalAvg: averageAmount(subset),
        womenPctOfMen: menAvg ? (womenAvg / menAvg) * 100 : 0
      };
    });
  }

  function buildVariablePayGroupMetrics(participants, employees) {
    const eligibleEmployees = employees.filter((employee) => employee.gender === "Kvinne" || employee.gender === "Mann");

    return variablePayOrder.map((groupName) => {
      const employeeSubset =
        groupName === "Snitt alle ansatte"
          ? eligibleEmployees
          : eligibleEmployees.filter((employee) => variablePayGroup(employee.group) === groupName);
      const participantSubset =
        groupName === "Snitt alle ansatte"
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
    });
  }

  function buildVariablePaySpreadRows(participants) {
    return variablePayOrder
      .filter((groupName) => groupName !== "Snitt alle ansatte")
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

  function buildVariablePayMonthlyRows(rows, employeeIndex) {
    const totals = new Map();

    for (const row of rows) {
      const personName = row["Etternavn, fornavn"];
      const payoutDate = parseCellDate(row["Utbetalingsdato"]);
      if (!personName || !payoutDate) continue;

      const employee = employeeIndex.byName.get(normalizePersonName(personName));
      if (!employee || (employee.gender !== "Kvinne" && employee.gender !== "Mann")) continue;

      const key = monthKeyFromDate(payoutDate);
      const current = totals.get(key) || {
        måned: formatMonthLabel(key),
        sortKey: key,
        kvinner: 0,
        menn: 0,
        totalt: 0
      };
      const amount = Number(row["Beløp"] || 0);

      if (employee.gender === "Kvinne") {
        current.kvinner += amount;
      } else if (employee.gender === "Mann") {
        current.menn += amount;
      }

      current.totalt += amount;
      totals.set(key, current);
    }

    return [...totals.values()].sort((left, right) => left.sortKey.localeCompare(right.sortKey));
  }

  function buildVariablePayMonthHeatmap(rows, employeeIndex) {
    const totals = new Map();

    for (const row of rows) {
      const personName = row["Etternavn, fornavn"];
      const payoutDate = parseCellDate(row["Utbetalingsdato"]);
      if (!personName || !payoutDate) continue;

      const employee = employeeIndex.byName.get(normalizePersonName(personName));
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

      current.verdi += Number(row["Beløp"] || 0);
      totals.set(compoundKey, current);
    }

    return [...totals.values()].sort((left, right) => {
      const groupDiff = variablePayOrder.indexOf(left.gruppe) - variablePayOrder.indexOf(right.gruppe);
      if (groupDiff !== 0) return groupDiff;
      return left.sortKey.localeCompare(right.sortKey);
    });
  }

  function summarizeParentalLeave(rows) {
    const byPerson = new Map();

    for (const row of rows) {
      if (!row["Kjønn"] || !row["Fornavn"] || !row["Etternavn"]) continue;

      const key = `${row["Etternavn"]}|${row["Fornavn"]}|${row["Kjønn"]}`;
      const current = byPerson.get(key) || {
        gender: row["Kjønn"],
        days: 0,
        weightedDays: 0
      };

      const days = Number(row["Frav.dager"] || 0);
      const workAbility = Number(row["Arbeidsførhet"] || 0);

      current.days += days;
      current.weightedDays += days * (100 - workAbility) / 100;
      byPerson.set(key, current);
    }

    const employees = [...byPerson.values()];
    const women = employees.filter((employee) => employee.gender === "Kvinne");
    const men = employees.filter((employee) => employee.gender === "Mann");
    const totalWeightedDays = employees.reduce((sum, employee) => sum + employee.weightedDays, 0);
    const totalEmployees = employees.length;

    return {
      totalEmployees,
      womenCount: women.length,
      menCount: men.length,
      womenAvgWeeks: women.length ? women.reduce((sum, employee) => sum + employee.weightedDays, 0) / women.length / 5 : 0,
      menAvgWeeks: men.length ? men.reduce((sum, employee) => sum + employee.weightedDays, 0) / men.length / 5 : 0,
      womenShareDays: totalWeightedDays
        ? women.reduce((sum, employee) => sum + employee.weightedDays, 0) / totalWeightedDays * 100
        : 0,
      menShareDays: totalWeightedDays
        ? men.reduce((sum, employee) => sum + employee.weightedDays, 0) / totalWeightedDays * 100
        : 0
    };
  }

  function summarizeFastlonn(rows, snapshotDate) {
    const employeeIndex = buildEmployeeIndex(rows, snapshotDate);
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
        womenPctOfMen: menAvg ? (womenAvg / menAvg) * 100 : 0
      };
    });

    return { employees, employeeIndex, genderBalance, fastlonn };
  }

  function buildFastlonnRows(files) {
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

  function fileMatchesYear(file, year) {
    return (file.fileYear || file.snapshotKey?.slice(0, 4)) === year;
  }

  function buildRowsForRole(files, role) {
    return files
      .filter((file) => file.role === role && file.source === "rådata")
      .flatMap((file) => file.rows);
  }

  function getFastlonnFilesForSelection(files, selectedYear, selectedPeriodKey) {
    const availableFastlonnFiles = files.filter((file) => file.role === "fastlønn" && file.source === "rådata");
    const fastlonnFilesForYear = selectedYear
      ? availableFastlonnFiles.filter((file) => fileMatchesYear(file, selectedYear))
      : availableFastlonnFiles;
    const targetSnapshotDate = periodDateForYear(selectedYear, selectedPeriodKey);

    const fastlonnFilesForPeriod = targetSnapshotDate
      ? fastlonnFilesForYear.filter((file) => {
          const snapshotDate = parseSnapshotKeyDate(file.snapshotKey);
          return snapshotDate && snapshotDate.getTime() >= targetSnapshotDate.getTime();
        })
      : fastlonnFilesForYear;

    return {
      availableFastlonnFiles,
      fastlonnFilesForYear,
      fastlonnFilesForPeriod,
      targetSnapshotDate
    };
  }

  function filterVariablePayRowsByYear(rows, year) {
    if (!year) return rows;
    return rows.filter((row) => yearFromDate(parseCellDate(row["Utbetalingsdato"])) === year);
  }

  function filterParentalLeaveRowsByYear(rows, year) {
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

  function buildVariablePayParticipants(rows, employeeIndex) {
    const totalsByPerson = new Map();

    for (const row of rows) {
      if (!row["Etternavn, fornavn"]) continue;

      const employee = employeeIndex.byName.get(normalizePersonName(row["Etternavn, fornavn"]));
      if (!employee?.group) continue;

      const key = normalizePersonName(row["Etternavn, fornavn"]);
      const current = totalsByPerson.get(key) || {
        name: row["Etternavn, fornavn"],
        gender: employee.gender,
        group: employee.group,
        amount: 0
      };

      current.amount += Number(row["Beløp"] || 0);
      totalsByPerson.set(key, current);
    }

    return [...totalsByPerson.values()].sort((left, right) => right.amount - left.amount || left.name.localeCompare(right.name, "nb"));
  }

  function summarizeVariablePayParticipants(participants) {
    const women = participants.filter((participant) => participant.gender === "Kvinne").length;
    const men = participants.filter((participant) => participant.gender === "Mann").length;
    return {
      total: participants.length,
      women,
      men
    };
  }

  function buildParentalLeaveEmployees(rows) {
    const byPerson = new Map();

    for (const row of rows) {
      if (!row["Kjønn"] || !row["Fornavn"] || !row["Etternavn"]) continue;

      const key = `${row["Etternavn"]}|${row["Fornavn"]}|${row["Kjønn"]}`;
      const current = byPerson.get(key) || {
        name: `${row["Etternavn"]}, ${row["Fornavn"]}`,
        gender: row["Kjønn"],
        days: 0,
        weightedDays: 0
      };

      const days = Number(row["Frav.dager"] || 0);
      const workAbility = Number(row["Arbeidsførhet"] || 0);

      current.days += days;
      current.weightedDays += days * (100 - workAbility) / 100;
      byPerson.set(key, current);
    }

    return [...byPerson.values()]
      .map((employee) => ({
        ...employee,
        weeks: employee.weightedDays / 5
      }))
      .sort((left, right) => right.weightedDays - left.weightedDays || left.name.localeCompare(right.name, "nb"));
  }

  function buildParentalLeaveGroupedEmployees(rows, employeeIndex) {
    const byPerson = new Map();

    for (const row of rows) {
      if (!row["Kjønn"] || !row["Fornavn"] || !row["Etternavn"]) continue;

      const fullName = `${row["Etternavn"] || ""} ${row["Fornavn"] || ""}`.trim();
      const employee = employeeIndex.byName.get(normalizePersonName(fullName));
      if (!employee?.group) continue;

      const key = `${row["Etternavn"]}|${row["Fornavn"]}|${row["Kjønn"]}`;
      const current = byPerson.get(key) || {
        name: `${row["Etternavn"]}, ${row["Fornavn"]}`,
        gender: row["Kjønn"],
        group: employee.group,
        days: 0,
        weightedDays: 0
      };

      const days = Number(row["Frav.dager"] || 0);
      const workAbility = Number(row["Arbeidsførhet"] || 0);

      current.days += days;
      current.weightedDays += days * (100 - workAbility) / 100;
      byPerson.set(key, current);
    }

    return [...byPerson.values()]
      .map((employee) => ({
        ...employee,
        weeks: employee.weightedDays / 5
      }))
      .sort((left, right) => right.weightedDays - left.weightedDays || left.name.localeCompare(right.name, "nb"));
  }

  function buildParentalLeaveGroupCountChartData(employees) {
    return groupOrder
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

  function buildParentalLeaveGroupParticipationData(participants, employees) {
    return groupOrder
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

  function buildParentalLeaveSpreadRows(employees) {
    return groupOrder
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

  function buildReport(files, selectedYear, selectedPeriodKey) {
    const { fastlonnFilesForPeriod, targetSnapshotDate } = getFastlonnFilesForSelection(files, selectedYear, selectedPeriodKey);
    const fastlonnRows = buildFastlonnRows(fastlonnFilesForPeriod);
    const overtidRows = filterVariablePayRowsByYear(buildRowsForRole(files, "overtid"), selectedYear);
    const vakttilleggRows = filterVariablePayRowsByYear(buildRowsForRole(files, "vakttillegg"), selectedYear);
    const foreldrepermisjonRows = filterParentalLeaveRowsByYear(buildRowsForRole(files, "foreldrepermisjon"), selectedYear);
    const fastlonn = fastlonnRows.length ? summarizeFastlonn(fastlonnRows, targetSnapshotDate) : null;
    const employment = fastlonn ? summarizeEmployment(fastlonn.employees) : null;
    const overtid = overtidRows.length && fastlonn ? summarizeVariablePay(overtidRows, fastlonn.employeeIndex) : null;
    const vakttillegg = vakttilleggRows.length && fastlonn ? summarizeVariablePay(vakttilleggRows, fastlonn.employeeIndex) : null;
    const foreldrepermisjon = foreldrepermisjonRows.length ? summarizeParentalLeave(foreldrepermisjonRows) : null;
    const overtidParticipants = fastlonn ? buildVariablePayParticipants(overtidRows, fastlonn.employeeIndex) : [];
    const vakttilleggParticipants = fastlonn ? buildVariablePayParticipants(vakttilleggRows, fastlonn.employeeIndex) : [];

    const notes = [];
    if (files.some((file) => file.source === "bearbeidet")) {
      notes.push("Bearbeidede filer er lastet opp. Målet er at løsningen skal fungere med rådata alene.");
    }
    if (files.some((file) => file.source === "referanse")) {
      notes.push("Utregningsskjema er lastet opp som referanse. Det skal ikke være nødvendig i den ferdige automatiseringen.");
    }
    if (selectedYear) {
      if (!fastlonnFilesForPeriod.length) {
        notes.push(`Mangler fastlønnfil for ${selectedPeriodKey === "05-01" ? "01.05" : "31.12"} i ${selectedYear}.`);
      }
      if (!overtidRows.length) {
        notes.push(`Fant ingen overtidsrader for ${selectedYear} basert på Utbetalingsdato.`);
      }
      if (!vakttilleggRows.length) {
        notes.push(`Fant ingen vakttilleggsrader for ${selectedYear} basert på Utbetalingsdato.`);
      }
      if (!foreldrepermisjonRows.length) {
        notes.push(`Fant ingen foreldrepermisjonsrader for ${selectedYear} basert på Start og Slutt.`);
      }
    }

    return {
      generatedAt: new Date().toISOString(),
      fastlonnSnapshotLabel: targetSnapshotDate ? `${selectedPeriodKey === "05-01" ? "01.05" : "31.12"}.${selectedYear}` : "",
      fastlonnSnapshotKey: targetSnapshotDate ? `${selectedYear}-${selectedPeriodKey}` : "",
      fastlonn,
      employment,
      overtid,
      overtidParticipants: summarizeVariablePayParticipants(overtidParticipants),
      vakttillegg,
      vakttilleggParticipants: summarizeVariablePayParticipants(vakttilleggParticipants),
      foreldrepermisjon,
      notes
    };
  }

  async function parseExcelFiles(files) {
    const parsed = [];

    for (const file of files) {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const headerRow = matrix.find((row) => row.some((cell) => String(cell).trim() !== "")) || [];
      const headerIndex = matrix.indexOf(headerRow);
      const rows = matrix
        .slice(headerIndex + 1)
        .filter((row) => row.some((cell) => String(cell).trim() !== ""));
      const rowObjects = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const meta = detectFileMeta(file.name, headerRow.map((cell) => String(cell)));
      const derivedFileYear = deriveFileYear(meta.role, rowObjects, meta.fileYear || "");

      parsed.push({
        fileName: file.name,
        label: meta.label,
        role: meta.role,
        source: meta.source,
        snapshotKey: meta.snapshotKey || "",
        snapshotLabel: meta.snapshotLabel || "",
        fileYear: derivedFileYear,
        rowCount: rows.length,
        columnCount: headerRow.length,
        header: headerRow.map((cell) => String(cell)),
        preview: rows.slice(0, 5).map((row) => row.map((cell) => String(cell))),
        rows: rowObjects
      });
    }

    return parsed;
  }

  function fileIdentity(file) {
    if (file.role !== "ukjent") {
      return `${file.role}::${file.source}::${file.snapshotKey || ""}::${file.fileYear || ""}`;
    }
    return `${file.fileName}::${file.source}`;
  }

  function mergeFiles(existingFiles, newFiles) {
    const merged = [...existingFiles];

    for (const newFile of newFiles) {
      const identity = fileIdentity(newFile);
      const existingIndex = merged.findIndex((file) => fileIdentity(file) === identity);

      if (existingIndex >= 0) {
        merged[existingIndex] = newFile;
      } else {
        merged.push(newFile);
      }
    }

    return merged;
  }

  function containsNorwegianCharacters(value) {
    return norwegianCharacterPattern.test(String(value ?? ""));
  }

  function buildReadErrorMessage(files, fallbackMessage, caughtError) {
    const baseMessage = caughtError?.message || fallbackMessage;
    const hasNorwegianCharacters = files.some((file) => containsNorwegianCharacters(file?.name));

    if (!hasNorwegianCharacters) return baseMessage;

    return `${baseMessage} Filnavnet inneholder æ, ø eller å, og det kan gjøre maskinell lesing vanskeligere.`;
  }

  async function applyFiles(files, sourceLabel) {
    const parsed = await parseExcelFiles(files);
    uploadedFiles = mergeFiles(uploadedFiles, parsed);
    dataSourceLabel = sourceLabel;
  }

  async function loadDefaultFiles() {
    if (!defaultDataFiles.length) return;

    loading = true;
    error = "";

    try {
      const results = await Promise.allSettled(
        defaultDataFiles.map(async (fileName) => {
          const url = `${base}/default_data/${encodeURIComponent(fileName)}`;
          const response = await fetch(url);
          if (!response.ok) {
            throw new Error(`Kunne ikke lese standardfilen ${fileName}.`);
          }

          const blob = await response.blob();
          return new File([blob], fileName, {
            type: blob.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          });
        })
      );

      const successfulFiles = results
        .filter((result) => result.status === "fulfilled")
        .map((result) => result.value);
      const failedMessages = results
        .filter((result) => result.status === "rejected")
        .map((result) => result.reason?.message || "Ukjent feil ved innlasting av standardfil.");

      if (successfulFiles.length) {
        await applyFiles(successfulFiles, "Standardfiler fra default_data");
      }

      if (failedMessages.length) {
        error = failedMessages.join(" ");
      }
    } catch (caughtError) {
      error = caughtError?.message || "Kunne ikke lese standardfilene.";
    } finally {
      loading = false;
    }
  }

  async function handleUpload(event) {
    const incoming = Array.from(event.currentTarget.files || []);
    if (!incoming.length) return;

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
      event.currentTarget.value = "";
    }
  }

  async function handleRoleUpload(role, event) {
    const incoming = Array.from(event.currentTarget.files || []);
    if (!incoming.length) return;

    loading = true;
    error = "";

    try {
      const parsed = await parseExcelFiles(incoming);
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
      event.currentTarget.value = "";
    }
  }

  function clearUploadedFiles() {
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

  const reportTableBorderColor = "000000";
  const reportTableHeaderFill = "EAF1FB";

  function reportTableBorders() {
    return {
      top: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor },
      left: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor },
      right: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor }
    };
  }

  function reportTableCell(text, options = {}) {
    const {
      bold = false,
      align = AlignmentType.LEFT,
      shaded = false,
      rowSpan,
      columnSpan
    } = options;

    return new TableCell({
      ...(rowSpan ? { rowSpan } : {}),
      ...(columnSpan ? { columnSpan } : {}),
      shading: shaded ? { type: ShadingType.CLEAR, color: "auto", fill: reportTableHeaderFill } : undefined,
      borders: reportTableBorders(),
      children: [
        new Paragraph({
          alignment: align,
          children: [
            new TextRun({
              text: String(text),
              bold,
              color: "000000"
            })
          ]
        })
      ]
    });
  }

  function tableFromRows(rows) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: rows.map((row, rowIndex) =>
        new TableRow({
          children: row.map((cell) =>
            reportTableCell(cell, {
              bold: rowIndex === 0,
              align: rowIndex === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
              shaded: rowIndex === 0
            })
          )
        })
      )
    });
  }

  function emptyLine() {
    return new Paragraph({
      text: "",
      spacing: {
        after: 260
      }
    });
  }

  function promptParagraph(text) {
    return new Paragraph({
      spacing: {
        before: 120,
        after: 180
      },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: "F4F7FB" },
      children: [
        new TextRun({
          text,
          italics: true,
          color: "000000"
        })
      ]
    });
  }

  function noteParagraph(text) {
    return new Paragraph({
      spacing: {
        before: 80,
        after: 120
      },
      children: [
        new TextRun({
          text,
          italics: true
        })
      ]
    });
  }

  function standardParagraph(text) {
    return new Paragraph({
      spacing: {
        before: 120,
        after: 180
      },
      children: [new TextRun({ text })]
    });
  }

  function createEmptyManualReportInputs() {
    return {
      nyansatteCount: "",
      samletSykefravaer: "",
      sykefravaerKvinner: "",
      sykefravaerMenn: ""
    };
  }

  function normalizeManualReportInputs(value) {
    const defaults = createEmptyManualReportInputs();
    const input = value && typeof value === "object" ? value : {};
    return { ...defaults, ...input };
  }

  function manualValue(value, fallbackLabel) {
    const trimmed = String(value ?? "").trim();
    return trimmed || `[MANUELL: ${fallbackLabel}]`;
  }

  function manualPercentValue(value, fallbackLabel) {
    const trimmed = String(value ?? "").trim();
    if (!trimmed) return `[MANUELL: ${fallbackLabel}]`;
    return trimmed.includes("%") ? trimmed : `${trimmed} %`;
  }

  function updateManualReportInput(field, value) {
    manualReportInputs = {
      ...manualReportInputs,
      [field]: value
    };
    persistManualReportInputs();
  }

  function persistManualReportInputs() {
    if (typeof window === "undefined") return;
    manualReportInputCache = {
      ...manualReportInputCache,
      [activeManualReportInputKey]: manualReportInputs
    };
    window.localStorage.setItem(manualReportStorageKey, JSON.stringify(manualReportInputCache));
  }

  function manualNoteBlock(prompt) {
    return [
      promptParagraph(`Manuell vurdering/notat: ${prompt}`),
      emptyLine()
    ];
  }

  function totalGenderRow(summary) {
    return summary?.fastlonn?.genderBalance?.find((row) => row.group === "Totalt antall ansatte") ?? null;
  }

  function leaderSummary(summary) {
    const rows = summary?.fastlonn?.genderBalance ?? [];
    const leaderGroups = ["Direktørgruppen", "Seksjonssjefgruppen"];
    const selected = rows.filter((row) => leaderGroups.includes(row.group));
    return {
      total: selected.reduce((sum, row) => sum + row.total, 0),
      women: selected.reduce((sum, row) => sum + row.women, 0)
    };
  }

  function standardReportNarratives(summary, selectedYear, manualInputs = createEmptyManualReportInputs()) {
    const year = selectedYear || new Date(summary.generatedAt).getFullYear();
    const totalRow = totalGenderRow(summary);
    const leaders = leaderSummary(summary);
    const overtimeParticipants = summary?.overtidParticipants?.total ?? summary?.overtid?.find((row) => row.group === "Snitt alle ansatte")?.n ?? 0;
    const vakttilleggParticipants = summary?.vakttilleggParticipants?.total ?? summary?.vakttillegg?.find((row) => row.group === "Snitt alle ansatte")?.n ?? 0;

    return {
      introduction:
        totalRow
          ? `Det ble ansatt ${manualValue(manualInputs.nyansatteCount, "antall nyansatte")} nye medarbeidere i ${year}, og ved slutten av året var det totalt ${formatInteger(totalRow.total)} ansatte i virksomheten. Virksomheten hadde et samlet sykefravær på ${manualPercentValue(manualInputs.samletSykefravaer, "samlet sykefravær i prosent")}.`
          : "",
      genderBalance:
        totalRow && leaders.total
          ? `Ved utgangen av ${year} var det ${formatPercent(totalRow.womenShare)} kvinner og ${formatPercent(totalRow.menShare)} menn i virksomheten. Antall ledere med personalansvar var ${formatInteger(leaders.total)}, hvorav ${formatInteger(leaders.women)} var kvinner. Andel kvinner i ledelsen er dermed ${formatPercent((leaders.women / leaders.total) * 100)}.`
          : "",
      fastlonn:
        summary?.fastlonnSnapshotLabel
          ? `Lønnskartleggingen er basert på fastlønn per ${summary.fastlonnSnapshotLabel}. Tabellen viser gjennomsnittlig fastlønn per stillingsgruppe fordelt på kvinner og menn.`
          : "",
      overtid:
        summary?.overtid
          ? `I løpet av ${year} opparbeidet ${formatInteger(overtimeParticipants)} ansatte overtid, noe som utgjør ${formatPercent(totalRow?.total ? (overtimeParticipants / totalRow.total) * 100 : 0)} av Nkoms ansatte. Tabellen viser gjennomsnittlig overtidskompensasjon per gruppe fordelt på kvinner og menn.`
          : "",
      vakttillegg:
        summary?.vakttillegg
          ? `Nkom har tre beredskapsvaktordninger: for hendelser i ekom-nettene, logisk sikkerhet (EkomCERT) og intern IT-drift. I ${year} deltok totalt ${formatInteger(vakttilleggParticipants)} ansatte i disse ordningene, fordelt på ${formatInteger(summary?.vakttilleggParticipants?.women ?? 0)} kvinner og ${formatInteger(summary?.vakttilleggParticipants?.men ?? 0)} menn. Tabellen viser gjennomsnittlig vakttillegg per gruppe fordelt på kvinner og menn.`
          : "",
      employment:
        summary?.employment
          ? `Tabellen viser fordelingen av midlertidige ansatte og faktisk deltid ved utgangen av ${year}, fordelt på kjønn.`
          : "",
      foreldrepermisjon:
        summary?.foreldrepermisjon
          ? `Det var i ${year} totalt ${formatInteger(summary.foreldrepermisjon.totalEmployees)} medarbeidere som tok ut foreldrepermisjon, ${formatInteger(summary.foreldrepermisjon.womenCount)} kvinner og ${formatInteger(summary.foreldrepermisjon.menCount)} menn. Uttak av foreldrepermisjon basert på gjennomsnittlig antall uker fordeler seg med henholdsvis ${formatNumber(summary.foreldrepermisjon.womenAvgWeeks)} uker for kvinner og ${formatNumber(summary.foreldrepermisjon.menAvgWeeks)} uker for menn. Uttak av foreldrepermisjon basert på andel dager av total er på henholdsvis ${formatPercent(summary.foreldrepermisjon.womenShareDays)} for kvinner og ${formatPercent(summary.foreldrepermisjon.menShareDays)} for menn.`
          : "",
      sykefravaer:
        `Det totale sykefraværet for ${year} var på ${manualPercentValue(manualInputs.samletSykefravaer, "samlet sykefravær %")}, hvorav det legemeldte sykefraværet ble henholdsvis ${manualPercentValue(manualInputs.sykefravaerKvinner, "sykefravær kvinner %")} for kvinner og ${manualPercentValue(manualInputs.sykefravaerMenn, "sykefravær menn %")} for menn.`
    };
  }

  function reportHistoryYears(selectedYear, summary, count = 5) {
    const endYear = Number(selectedYear || new Date(summary.generatedAt).getFullYear());
    return Array.from({ length: count }, (_, index) => String(endYear - (count - 1) + index));
  }

  function narrativeBlock(title, prompt) {
    return [
      new Paragraph({ text: title, heading: HeadingLevel.HEADING_2 }),
      promptParagraph(prompt),
      emptyLine()
    ];
  }

  function genderBalanceLabel(group) {
    if (group === "Totalt antall ansatte") return "Totalt antall ansatte *";
    if (group === "Direktørgruppen") return "Direktørgruppen**";
    return group;
  }

  function genderBalanceTable(rows) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            reportTableCell("Stillingsgrupper", { bold: true, shaded: true, rowSpan: 2 }),
            reportTableCell("Kvinner", { bold: true, shaded: true, align: AlignmentType.CENTER, columnSpan: 2 }),
            reportTableCell("Menn", { bold: true, shaded: true, align: AlignmentType.CENTER, columnSpan: 2 }),
            reportTableCell("Totalt", { bold: true, shaded: true, align: AlignmentType.CENTER })
          ]
        }),
        new TableRow({
          children: [
            reportTableCell("Antall", { bold: true, shaded: true, align: AlignmentType.CENTER }),
            reportTableCell("Andel", { bold: true, shaded: true, align: AlignmentType.CENTER }),
            reportTableCell("Antall", { bold: true, shaded: true, align: AlignmentType.CENTER }),
            reportTableCell("Andel", { bold: true, shaded: true, align: AlignmentType.CENTER }),
            reportTableCell("Antall", { bold: true, shaded: true, align: AlignmentType.CENTER })
          ]
        }),
        ...rows.map(
          (row) =>
            new TableRow({
              children: [
                reportTableCell(genderBalanceLabel(row.group)),
                ...[
                  row.women,
                  formatPercentValue(row.womenShare),
                  row.men,
                  formatPercentValue(row.menShare),
                  row.total
                ].map((cell) => reportTableCell(cell, { align: AlignmentType.CENTER }))
              ]
            })
        )
      ]
    });
  }

  async function downloadReport() {
    const summary = report;
    if (!summary) return;
    const narratives = standardReportNarratives(summary, selectedFastlonnYear, manualReportInputs);
    const historyYears = reportHistoryYears(selectedFastlonnYear, summary);
    const reportYear = selectedFastlonnYear || new Date(summary.generatedAt).getFullYear();

    const children = [
      new Paragraph({ text: `HR Årsrapport ${reportYear}`, heading: HeadingLevel.TITLE }),
      new Paragraph({
        children: [new TextRun(`Generert: ${new Date(summary.generatedAt).toLocaleString("nb-NO")}`)]
      }),
      ...(summary.fastlonnSnapshotLabel
        ? [new Paragraph({ children: [new TextRun(`Fastlønn brukt per: ${summary.fastlonnSnapshotLabel}`)] })]
        : []),
      promptParagraph("Denne Word-filen er ikke låst. Erstatt grå hjelpetekst med endelig formulering direkte i dokumentet etter nedlasting."),
      emptyLine(),
      new Paragraph({ text: "Opplastede filer", heading: HeadingLevel.HEADING_1 }),
      tableFromRows([
        ["Fil", "Type", "Kilde", "Lønnsdato", "Rader", "Kolonner"],
        ...uploadedFiles.map((file) => [
          file.fileName,
          file.label,
          file.source,
          file.snapshotLabel || "",
          file.rowCount,
          file.columnCount
        ])
      ]),
      emptyLine(),
      new Paragraph({ text: "Personalmessige forhold", heading: HeadingLevel.HEADING_2 }),
      ...(narratives.introduction ? [standardParagraph(narratives.introduction)] : []),
      ...manualNoteBlock("Skriv kort oppsummering av bemanningssituasjon, nyansettelser og hovedtrekk for året."),
      ...narrativeBlock(
        "Positiv utvikling i ansettelser av personer med funksjonsnedsettelse / inkludering i rekruttering",
        "Skriv vurderingstekst som følger malens struktur: tiltak, utvikling fra tidligere år, resultater og planlagte forbedringer."
      ),
      promptParagraph(
        "Manuell utfylling: legg inn historiske tall for ansettelser, målgruppe, prosentoppnåelse og praksisplasser her dersom disse ikke kommer fra opplastede rådatafiler."
      ),
      tableFromRows([
        ["Kategori", ...historyYears],
        ["Totalt antall ansettelser", "", "", "", "", ""],
        ["Ansettelser i målgruppen - totalt", "", "", "", "", ""],
        ["Faste ansettelser", "", "", "", "", ""],
        ["Midlertidige administrative ansettelser", "", "", "", "", ""],
        ["Traineeprogrammet i staten", "", "", "", "", ""],
        ["Midlertidig m/ lønnstilskudd", "", "", "", "", ""],
        ["Overgang fra midlertidig til fast (3-årsregelen)", "", "", "", "", ""],
        ["Prosentoppnåelse", "", "", "", "", ""],
        ["Totalt", "", "", "", "", ""],
        ["Praksisplasser / arbeidstrening (ikke med i tallene)", "", "", "", "", ""]
      ]),
      emptyLine(),
      ...narrativeBlock(
        "Samarbeid med utdanningsinstitusjoner for kompetansebygging",
        "Skriv tekst om praksisplasser, lærlinger, trainee-ordninger og samarbeid med utdanningsinstitusjoner."
      ),
      ...narrativeBlock(
        "Forsterket aktivitets- og redegjørelsesplikt (ARP)",
        "Skriv tekst om hvordan ARP er fulgt opp i virksomheten, hvilke tiltak som er gjennomført og hvordan arbeidet evalueres."
      ),
      new Paragraph({
        text: `Kartlegging av kjønnsbalanse, lønn og ufrivillig deltid${summary.fastlonnSnapshotLabel ? ` (fastlønn per ${summary.fastlonnSnapshotLabel})` : ""}`,
        heading: HeadingLevel.HEADING_1
      })
    ];

    if (summary.fastlonn) {
      children.push(new Paragraph({ text: "KJØNNSBALANSE", heading: HeadingLevel.HEADING_2 }));
      children.push(genderBalanceTable(summary.fastlonn.genderBalance));
      children.push(
        noteParagraph("*inkludert direktør"),
        noteParagraph("**inkluderer direktør samt to underdirektører med personalansvar som inngår i virksomhetens ledergruppe."),
        ...(narratives.genderBalance ? [standardParagraph(narratives.genderBalance)] : []),
        ...manualNoteBlock("Forklar kjønnsbalanse totalt, i ledelsen og eventuelle forhold ved fagområder eller rekruttering.")
      );

      children.push(new Paragraph({ text: "FASTLØNN", heading: HeadingLevel.HEADING_2 }));
      children.push(
        tableFromRows([
          ["Gruppe", "N", "Kvinner", "Menn", "Totalt", "Kvinner i % av menn"],
          ...summary.fastlonn.fastlonn.map((row) => [
            row.group,
            row.n,
            formatNumber(row.womenAvg),
            formatNumber(row.menAvg),
            formatNumber(row.totalAvg),
            formatPercent(row.womenPctOfMen)
          ])
        ])
      );
      children.push(
        ...(narratives.fastlonn ? [standardParagraph(narratives.fastlonn)] : []),
        ...manualNoteBlock("Forklar eventuelle avvik mellom gruppene og legg inn individuell faglig vurdering.")
      );
    }

    if (summary.overtid) {
      children.push(new Paragraph({ text: "VARIABLE TILLEGG - OVERTID", heading: HeadingLevel.HEADING_2 }));
      children.push(
        tableFromRows([
          ["Gruppe", "N", "Kvinner", "Menn", "Totalt", "Kvinner i % av menn"],
          ...summary.overtid.map((row) => [
            `${row.group} (N=${row.n})`,
            row.n,
            formatNumber(row.womenAvg),
            formatNumber(row.menAvg),
            formatNumber(row.totalAvg),
            formatPercent(row.womenPctOfMen)
          ])
        ])
      );
      children.push(
        ...(narratives.overtid ? [standardParagraph(narratives.overtid)] : []),
        ...manualNoteBlock("Beskriv omfang, mulige årsaker til forskjeller og eventuelle merknader til små utvalg.")
      );
    }

    if (summary.vakttillegg) {
      children.push(new Paragraph({ text: "VARIABLE TILLEGG - VAKTTILLEGG", heading: HeadingLevel.HEADING_2 }));
      children.push(
        tableFromRows([
          ["Gruppe", "N", "Kvinner", "Menn", "Totalt", "Kvinner i % av menn"],
          ...summary.vakttillegg.map((row) => [
            `${row.group} (N=${row.n})`,
            row.n,
            formatNumber(row.womenAvg),
            formatNumber(row.menAvg),
            formatNumber(row.totalAvg),
            formatPercent(row.womenPctOfMen)
          ])
        ])
      );
      children.push(
        ...(narratives.vakttillegg ? [standardParagraph(narratives.vakttillegg)] : []),
        ...manualNoteBlock("Beskriv beredskapsordninger, involverte fagmiljøer og forklar kjønnsforskjeller.")
      );
    }

    if (summary.employment || summary.foreldrepermisjon) {
      children.push(
        new Paragraph({
          text: "MIDLERTIDIG ANSATTE, DELTID, FORELDREPERMISJONER OG LEGEMELDT SYKEFRAVÆR",
          heading: HeadingLevel.HEADING_1
        })
      );
    }

    if (summary.employment) {
      children.push(new Paragraph({ text: "Midlertidige ansatte og faktisk deltid", heading: HeadingLevel.HEADING_2 }));
      children.push(
        tableFromRows([
          ["Kategori", "Kvinner antall", "Kvinner andel", "Menn antall", "Menn andel"],
          [
            "Midlertidig ansatte",
            summary.employment.temporary.women,
            formatPercent(summary.employment.totalWomen ? summary.employment.temporary.women / summary.employment.totalWomen * 100 : 0),
            summary.employment.temporary.men,
            formatPercent(summary.employment.totalMen ? summary.employment.temporary.men / summary.employment.totalMen * 100 : 0)
          ],
          [
            "Faktisk deltid",
            summary.employment.partTime.women,
            formatPercent(summary.employment.totalWomen ? summary.employment.partTime.women / summary.employment.totalWomen * 100 : 0),
            summary.employment.partTime.men,
            formatPercent(summary.employment.totalMen ? summary.employment.partTime.men / summary.employment.totalMen * 100 : 0)
          ]
        ])
      );
      children.push(
        ...(narratives.employment ? [standardParagraph(narratives.employment)] : []),
        ...manualNoteBlock("Beskriv midlertidighet og deltid, og vurder om deltid er frivillig eller ufrivillig.")
      );
    }

    if (summary.foreldrepermisjon) {
      children.push(new Paragraph({ text: "Foreldrepermisjon", heading: HeadingLevel.HEADING_2 }));
      children.push(
        tableFromRows([
          ["Måling", "Kvinner", "Menn"],
          ["Antall ansatte med foreldrepermisjon", summary.foreldrepermisjon.womenCount, summary.foreldrepermisjon.menCount],
          ["Gjennomsnittlig uttak i uker", formatNumber(summary.foreldrepermisjon.womenAvgWeeks), formatNumber(summary.foreldrepermisjon.menAvgWeeks)],
          ["Andel permisjonsdager av total", formatPercent(summary.foreldrepermisjon.womenShareDays), formatPercent(summary.foreldrepermisjon.menShareDays)]
        ])
      );
      children.push(
        ...(narratives.foreldrepermisjon ? [standardParagraph(narratives.foreldrepermisjon)] : []),
        ...manualNoteBlock("Beskriv omfang, kjønnsfordeling og eventuelle forbehold i datagrunnlaget.")
      );
    }

    children.push(
      new Paragraph({ text: "Legemeldt sykefravær", heading: HeadingLevel.HEADING_2 }),
      standardParagraph(narratives.sykefravaer),
      promptParagraph("Skriv inn sykefraværstall manuelt her dersom de ikke kommer fra rådatafilene."),
      emptyLine()
    );

    if (summary.notes.length) {
      children.push(new Paragraph({ text: "Merknader", heading: HeadingLevel.HEADING_1 }));
      for (const note of summary.notes) {
        children.push(new Paragraph({ text: note }));
      }
    }

    const doc = new Document({
      sections: [{ properties: {}, children }]
    });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    link.href = url;
    link.download = `hr-arsrapport-${reportYear}.docx`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function toggleSidebar() {
    sidebarCollapsed = !sidebarCollapsed;
    window.localStorage.setItem("hr-arsrapport-sidebar-collapsed", String(sidebarCollapsed));
  }

  onMount(() => {
    sidebarCollapsed = window.localStorage.getItem("hr-arsrapport-sidebar-collapsed") === "true";
    manualReportInputCache = JSON.parse(window.localStorage.getItem(manualReportStorageKey) || "{}");
    manualReportInputs = normalizeManualReportInputs(manualReportInputCache[activeManualReportInputKey]);
    loadedManualReportInputKey = activeManualReportInputKey;
    loadDefaultFiles();
  });
</script>

<svelte:head>
  <title>HR Årsrapport</title>
</svelte:head>

<section class="app-shell">
  <div class:sidebar-collapsed={sidebarCollapsed} class="app-frame">
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
      <div class="workspace-tablist" role="tablist" aria-label="Arbeidsflater">
        <button class={`workspace-tab ${activeTab === "opplasting" ? "active" : ""}`} role="tab" aria-selected={activeTab === "opplasting"} aria-label="Datagrunnlag" on:click={() => (activeTab = "opplasting")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8.5A2.5 2.5 0 0 1 6.5 6H10l2 2h5.5A2.5 2.5 0 0 1 20 10.5v7A2.5 2.5 0 0 1 17.5 20h-11A2.5 2.5 0 0 1 4 17.5Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M12 16v-5m0 0-2 2m2-2 2 2" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Datagrunnlag</span></button>
        <button class={`workspace-tab ${activeTab === "fastlonn" ? "active" : ""}`} role="tab" aria-selected={activeTab === "fastlonn"} aria-label="Fastlønn" on:click={() => (activeTab = "fastlonn")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8.5A2.5 2.5 0 0 1 6.5 6h11A2.5 2.5 0 0 1 20 8.5v7A2.5 2.5 0 0 1 17.5 18h-11A2.5 2.5 0 0 1 4 15.5Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M15.5 12h.01M8 12h3.5m-1.75-1.75v3.5" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Fastlønn</span></button>
        <button class={`workspace-tab ${activeTab === "overtid" ? "active" : ""}`} role="tab" aria-selected={activeTab === "overtid"} aria-label="Overtid" on:click={() => (activeTab = "overtid")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 7v5l3 2.5M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Overtid</span></button>
        <button class={`workspace-tab ${activeTab === "vakttillegg" ? "active" : ""}`} role="tab" aria-selected={activeTab === "vakttillegg"} aria-label="Vakttillegg" on:click={() => (activeTab = "vakttillegg")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M15.5 4.5A7.5 7.5 0 1 0 19 18a7 7 0 1 1-3.5-13.5Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M16.5 9.5h.01" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Vakttillegg</span></button>
        <button class={`workspace-tab ${activeTab === "foreldrepermisjon" ? "active" : ""}`} role="tab" aria-selected={activeTab === "foreldrepermisjon"} aria-label="Foreldrepermisjon" on:click={() => (activeTab = "foreldrepermisjon")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 10.5a2.5 2.5 0 1 1 5 0c0 1.3-.6 2.3-2.5 4-1.9-1.7-2.5-2.7-2.5-4Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/><path d="M5 19a4.5 4.5 0 0 1 4.5-4.5h4A4.5 4.5 0 0 1 18 19M8 8a2 2 0 1 1-4 0 2 2 0 0 1 4 0Zm12 0a2 2 0 1 1-4 0 2 2 0 0 1 4 0Z" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Foreldrepermisjon</span></button>
        <button class={`workspace-tab ${activeTab === "arsrapport" ? "active" : ""}`} role="tab" aria-selected={activeTab === "arsrapport"} aria-label="Årsrapport" on:click={() => (activeTab = "arsrapport")}><svg class="workspace-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3h7l5 5v13H7zM14 3v5h5M9 13h6M9 17h6" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8"/></svg><span>Årsrapport</span></button>
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

      {#if error}
        <div class="error-box mb-3 px-3 py-2 text-xs">{error}</div>
      {/if}

      {#if loading}
        <div class="loading-box mb-3 p-4 text-xs">Leser filer og bygger oversikt...</div>
      {/if}

      {#if activeTab === "opplasting"}
        <div class="space-y-3">
        <div class="content-card">
          <div class="panel-actions">
            <div>
              <p class="panel-eyebrow">Datagrunnlag</p>
              <h2 class="panel-title">Last opp lønnsgrunnlaget</h2>
            </div>
          </div>

          <div class="upload-grid gap-3">
            {#each uploadCardConfig as card}
              <label class={`upload-card cursor-pointer min-h-[196px] p-4 ${roleUploads[card.role] ? "upload-card-ready" : "upload-card-pending"}`}>
                <input class="hidden" type="file" accept=".xlsx,.xls" on:change={(event) => handleRoleUpload(card.role, event)} />
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
                  <span class={`upload-dropzone text-sm leading-6 ${roleUploads[card.role] ? "upload-dropzone-filled" : ""}`}>
                    {roleUploads[card.role] ? roleUploads[card.role].fileName : "Velg Excel-fil"}
                  </span>
                  <span class="upload-card-hint text-xs">XLSX eller XLS</span>
                </div>
              </label>
            {/each}
          </div>

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
                  <div class="snapshot-periods" role="group" aria-label="Velg lønnsdato">
                    {#each fastlonnPeriodChoices as period}
                      <button
                        type="button"
                        class={`snapshot-period ${selectedFastlonnPeriodKey === period.key ? "snapshot-period-active" : ""}`}
                        disabled={!period.available}
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
              <button class="app-button app-button-danger upload-meta-action" type="button" on:click={clearUploadedFiles} disabled={!uploadedFiles.length}>
                Tøm lagret data
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

          {#if report?.notes?.length}
            <div class="info-box mt-4 px-3 py-2 text-xs">
              <ul class="list-disc pl-5">
                {#each report.notes as note}
                  <li>{note}</li>
                {/each}
              </ul>
            </div>
          {/if}

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
        {#if report}
          <div class="fastlonn-shell space-y-3">
          <div class="content-card fastlonn-hero">
            <div class="fastlonn-hero-head">
              <div>
                <p class="panel-eyebrow">Fastlønn</p>
                <h2 class="panel-title">Fastlønn og representasjon</h2>
                <p class="section-note mt-2">
                  Fokus på signaler som bør forklares i årsrapporten{report.fastlonnSnapshotLabel ? `, basert på fastlønn per ${report.fastlonnSnapshotLabel}` : ""}.
                </p>
              </div>
            </div>
            <div class="fastlonn-kpi-grid">
              {#each priorityCards as card}
                <article class={`summary-card summary-card-quiet summary-card-kpi fastlonn-kpi-card ${card.tone}`}>
                  <p class="summary-card-label">{card.eyebrow}</p>
                  <p class="summary-card-metric">{card.metric}</p>
                  <h3 class="summary-card-title">{card.title}</h3>
                  <p class="summary-card-text">{card.takeaway}</p>
                </article>
              {/each}
            </div>
	          </div>
          <div class="fastlonn-main-grid">
            <article class="chart-card chart-card-feature chart-card-wide">
              <div class="chart-card-head">
                <div>
                  <p class="chart-card-eyebrow">Rapporttabell</p>
                  <h3>Representasjon per gruppe</h3>
                </div>
                <p class="chart-card-note">Bruk denne når du skal skrive presist om sammensetning og skjevfordeling.</p>
              </div>
              <div
                class="ledger-table-wrap mt-4 cursor-zoom-in"
                role="button"
                tabindex="0"
                aria-label="Åpne representasjon per gruppe i større visning"
                on:click={() => openExpandedView("fastlonn-representation-table", "Representasjon per gruppe")}
                on:keydown={(event) => handleExpandKeydown(event, "fastlonn-representation-table", "Representasjon per gruppe")}
              >
                <table class="ledger-table">
                  <thead>
                    <tr>
                      <th>Gruppe</th>
                      <th class="text-right">Kvinner</th>
                      <th class="text-right">Menn</th>
                      <th class="text-right">Kvinneandel</th>
                      <th class="text-right">Mannandel</th>
                      <th class="text-right">Totalt</th>
                    </tr>
                  </thead>
                  <tbody>
                    {#each fastlonnRepresentationTableRows as row}
                      <tr>
                        <td class="font-medium text-slate-900">{row.gruppe}</td>
                        <td class="text-right tabular-nums">{formatInteger(row.kvinner)}</td>
                        <td class="text-right tabular-nums">{formatInteger(row.menn)}</td>
                        <td class="text-right tabular-nums">{row.kvinneandel}</td>
                        <td class="text-right tabular-nums">{row.mannandel}</td>
                        <td class="text-right tabular-nums">{formatInteger(row.totalt)}</td>
                      </tr>
                    {/each}
                  </tbody>
                </table>
              </div>
            </article>
            <article class="chart-card chart-card-feature chart-card-wide">
              <div class="chart-card-head">
                <div>
                  <p class="chart-card-eyebrow">Rapporttabell</p>
                  <h3>Lønn per gruppe</h3>
                </div>
                <p class="chart-card-note">Hovedtabellen for lønnsnivå, sammenligning mellom kjønn og forklaring i rapporttekst.</p>
              </div>
              <div
                class="ledger-table-wrap mt-4 cursor-zoom-in"
                role="button"
                tabindex="0"
                aria-label="Åpne lønn per gruppe i større visning"
                on:click={() => openExpandedView("fastlonn-salary-table", "Lønn per gruppe")}
                on:keydown={(event) => handleExpandKeydown(event, "fastlonn-salary-table", "Lønn per gruppe")}
              >
                <table class="ledger-table">
                  <thead>
                    <tr>
                      <th>Gruppe</th>
                      <th class="text-right">Ansatte</th>
                      <th class="text-right">Kvinner</th>
                      <th class="text-right">Menn</th>
                      <th class="text-right">Totalt</th>
                      <th class="text-right">Kvinner av menn</th>
                    </tr>
                  </thead>
                  <tbody>
                    {#each fastlonnSalaryTableRows as row}
                      <tr>
                        <td class="font-medium text-slate-900">{row.gruppe}</td>
                        <td class="text-right tabular-nums">{formatInteger(row.ansatte)}</td>
                        <td class="text-right tabular-nums">{row.kvinner}</td>
                        <td class="text-right tabular-nums">{row.menn}</td>
                        <td class="text-right tabular-nums">{row.totalt}</td>
                        <td class="text-right tabular-nums">{row.lønnsforhold}</td>
                      </tr>
                    {/each}
                  </tbody>
                </table>
              </div>
            </article>
            {#if genderBalanceRows.length}
              <article class="chart-card chart-card-visual">
                <div class="chart-card-head">
                  <div>
                    <p class="chart-card-eyebrow">Graf</p>
                    <h3>Kjønnsbalanse per gruppe</h3>
                  </div>
                  <p class="chart-card-note">Rask oversikt over størrelse og skjevhet før du går til tabellen.</p>
                </div>
                <div
                  class="chart-card-figure cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne kjønnsbalanse per gruppe i større visning"
                  on:click={() => openExpandedView("fastlonn-gender-chart", "Kjønnsbalanse per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "fastlonn-gender-chart", "Kjønnsbalanse per gruppe")}
                >
                  <BarChart data={genderBalanceChartData} x="gruppe" y="antall" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={genderSeriesColors} yFmt="num0" xAxisTitle="Gruppe" yAxisTitle="Antall ansatte" />
                </div>
              </article>
            {/if}
            {#if salarySpreadRows.length}
              <article class="chart-card chart-card-visual">
                <div class="chart-card-head">
                  <div>
                    <p class="chart-card-eyebrow">Graf</p>
                    <h3>Lønnsspredning per gruppe</h3>
                  </div>
                  <p class="chart-card-note">Viser spenn og median, slik at du ser variasjon og ikke bare gjennomsnitt.</p>
                </div>
                <div
                  class="chart-card-figure cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne lønnsspredning per gruppe i større visning"
                  on:click={() => openExpandedView("fastlonn-salary-spread-chart", "Lønnsspredning per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "fastlonn-salary-spread-chart", "Lønnsspredning per gruppe")}
                >
                  <BoxPlot data={salarySpreadRows} name="gruppe" min="min" intervalBottom="q1" midpoint="median" intervalTop="q3" max="max" swapXY={true} yFmt="#,##0" yAxisTitle="Kroner" seriesColors={salarySpreadSeriesColors} />
                </div>
              </article>
            {/if}
            <article class="chart-card chart-card-support chart-card-wide">
              <div class="chart-card-head">
                <div>
                  <p class="chart-card-eyebrow">Støtte</p>
                  <h3>Støttetabeller og datagrunnlag</h3>
                </div>
                <p class="chart-card-note">Sekundær informasjon for analyse og kontroll, uten å fylle hovedbildet.</p>
              </div>
              <div class="mt-4 grid gap-3 lg:grid-cols-2">
                {#if fastlonnSpreadTableRows.length}
                  <article class="metric-card fastlonn-support-table">
                    <p class="summary-card-label">Lønnsspredning</p>
                    <div
                      class="ledger-table-wrap mt-3 cursor-zoom-in"
                      role="button"
                      tabindex="0"
                      aria-label="Åpne lønnsspredningstabell i større visning"
                      on:click={() => openExpandedView("fastlonn-spread-table", "Lønnsspredning")}
                      on:keydown={(event) => handleExpandKeydown(event, "fastlonn-spread-table", "Lønnsspredning")}
                    >
                      <table class="ledger-table ledger-table-dense">
                        <thead>
                          <tr>
                            <th>Gruppe</th>
                            <th class="text-right">Min</th>
                            <th class="text-right">Q1</th>
                            <th class="text-right">Median</th>
                            <th class="text-right">Q3</th>
                            <th class="text-right">Maks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {#each fastlonnSpreadTableRows as row}
                            <tr>
                              <td class="font-medium text-slate-900">{row.gruppe}</td>
                              <td class="text-right tabular-nums">{row.min}</td>
                              <td class="text-right tabular-nums">{row.q1}</td>
                              <td class="text-right tabular-nums">{row.median}</td>
                              <td class="text-right tabular-nums">{row.q3}</td>
                              <td class="text-right tabular-nums">{row.maks}</td>
                            </tr>
                          {/each}
                        </tbody>
                      </table>
                    </div>
                  </article>
                {/if}
                {#if fastlonnEmploymentTableRows.length}
                  <article class="metric-card fastlonn-support-table">
                    <p class="summary-card-label">Arbeidsforhold som støtteforklaring</p>
                    <div
                      class="ledger-table-wrap mt-3 cursor-zoom-in"
                      role="button"
                      tabindex="0"
                      aria-label="Åpne arbeidsforholdstabell i større visning"
                      on:click={() => openExpandedView("fastlonn-employment-table", "Arbeidsforhold som støtteforklaring")}
                      on:keydown={(event) => handleExpandKeydown(event, "fastlonn-employment-table", "Arbeidsforhold som støtteforklaring")}
                    >
                      <table class="ledger-table ledger-table-dense">
                        <thead>
                          <tr>
                            <th>Kategori</th>
                            <th>Kvinner</th>
                            <th>Menn</th>
                          </tr>
                        </thead>
                        <tbody>
                          {#each fastlonnEmploymentTableRows as row}
                            <tr>
                              <td class="font-medium text-slate-900">{row.kategori}</td>
                              <td>{row.kvinner}</td>
                              <td>{row.menn}</td>
                            </tr>
                          {/each}
                        </tbody>
                      </table>
                    </div>
                  </article>
                {/if}
              </div>
              <details class="fastlonn-details mt-4">
                <summary class="cursor-pointer list-none px-4 py-3 text-sm font-medium text-[#0f5368]">Vis datagrunnlag</summary>
                <div class="space-y-3 border-t border-slate-200 p-4">
                  {#if fastlonnEmployeePreview.length}
                    <div
                      class="ledger-table-wrap cursor-zoom-in"
                      role="button"
                      tabindex="0"
                      aria-label="Åpne ansatte i utvalg i større visning"
                      on:click={() => openExpandedView("fastlonn-employee-preview-table", "Ansatte i utvalg")}
                      on:keydown={(event) => handleExpandKeydown(event, "fastlonn-employee-preview-table", "Ansatte i utvalg")}
                    >
                      <table class="ledger-table ledger-table-dense">
                        <thead>
                          <tr>
                            <th>Ansatt</th>
                            <th>Gruppe</th>
                            <th>Kjønn</th>
                            <th class="text-right">Lønn</th>
                            <th>Arbeidsforhold</th>
                            <th class="text-right">Deltid %</th>
                          </tr>
                        </thead>
                        <tbody>
                          {#each fastlonnEmployeePreview as employee}
                            <tr>
                              <td class="font-medium text-slate-900">{employee.name}</td>
                              <td>{employee.group}</td>
                              <td>{employee.gender}</td>
                              <td class="text-right tabular-nums">{formatCurrency(employee.salary)}</td>
                              <td>{employee.employmentGroup}</td>
                              <td class="text-right tabular-nums">{formatNumber(employee.partTimePercent)}</td>
                            </tr>
                          {/each}
                        </tbody>
                      </table>
                    </div>
                  {/if}
                </div>
              </details>
            </article>
          </div>
          </div>
        {:else}
          <div class="content-card">
            <p class="panel-eyebrow">Fastlønn</p>
            <h2 class="panel-title">Ingen fastlønnsanalyse ennå</h2>
            <p class="section-note mt-2">Last opp grunnlagsfiler i fanen Datagrunnlag for å bygge oversikten.</p>
          </div>
        {/if}
      {/if}

      {#if activeTab === "overtid"}
        {#if report}
          <section class="space-y-3">
            <div class="content-card">
              <div class="fastlonn-hero-head">
                <div>
                  <p class="panel-eyebrow">Overtid</p>
                  <h2 class="panel-title">Overtid</h2>
                  <p class="section-note mt-2">Brukes som støtte for å forklare driftsbelastning og forskjeller i samlet kompensasjon.</p>
                </div>
                <div class="fastlonn-hero-meta">
                  <div class="fastlonn-meta-pill">
                    <span>Total overtid</span>
                    <strong>{formatCurrency(overtidParticipants.reduce((sum, employee) => sum + employee.amount, 0))}</strong>
                  </div>
                  <div class="fastlonn-meta-pill">
                    <span>Ansatte med overtid</span>
                    <strong>{formatInteger(overtidParticipants.length)}</strong>
                  </div>
                </div>
              </div>
            </div>
            <div class="grid gap-3">
              {#if overtidAnalysisRows.length}
                <div class="grid gap-3 xl:grid-cols-2">
                  <article class="chart-card">
                    <h3>Total overtid per gruppe</h3>
                    <p class="mt-1 text-xs text-slate-600">Dette er hovedbildet: hvor de samlede overtidskostnadene faktisk havner, vist i 1000 NOK for lettere sammenligning.</p>
                    <div
                      class="mt-4 cursor-zoom-in"
                      role="button"
                      tabindex="0"
                      aria-label="Åpne total overtid per gruppe i større visning"
                      on:click={() => openExpandedView("overtid-total-chart", "Total overtid per gruppe")}
                      on:keydown={(event) => handleExpandKeydown(event, "overtid-total-chart", "Total overtid per gruppe")}
                    >
                      <BarChart data={overtidTotalChartData} x="gruppe" y="verdi_tusen" yFmt="#,##0.0" xAxisTitle="Gruppe" yAxisTitle="1000 NOK" swapXY={true} colorPalette={overtidTotalChartPalette} seriesOptions={{ colorBy: "data" }} />
                    </div>
                  </article>

                  <article class="chart-card">
                    <h3>Deltakelsesgrad per gruppe</h3>
                    <p class="mt-1 text-xs text-slate-600">Viser om overtiden er bredt fordelt eller bæres av få ansatte i hver gruppe.</p>
                    <div
                      class="mt-4 cursor-zoom-in"
                      role="button"
                      tabindex="0"
                  aria-label="Åpne deltakelsesgrad for overtid i større visning"
                  on:click={() => openExpandedView("overtid-participation-chart", "Deltakelsesgrad per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "overtid-participation-chart", "Deltakelsesgrad per gruppe")}
                >
                      <BarChart data={overtidParticipationChartData} x="gruppe" y="andel" yFmt="pct1" xAxisTitle="Gruppe" yAxisTitle="Andel ansatte" swapXY={true} colorPalette={overtidParticipationChartPalette} seriesOptions={{ colorBy: "data" }} />
                    </div>
                  </article>

                  <article class="chart-card">
                    <h3>Overtid per gruppe</h3>
                    <p class="mt-1 text-xs text-slate-600">Gjennomsnittlig overtidskompensasjon per deltaker, fordelt på kvinner og menn.</p>
                    <div
                      class="mt-4 cursor-zoom-in"
                      role="button"
                      tabindex="0"
                      aria-label="Åpne overtid per gruppe i større visning"
                      on:click={() => openExpandedView("overtid-chart", "Overtid per gruppe")}
                      on:keydown={(event) => handleExpandKeydown(event, "overtid-chart", "Overtid per gruppe")}
                    >
                      <BarChart data={overtidChartData} x="gruppe" y="verdi" series="serie" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={paySeriesColors} yFmt="#,##0" xAxisTitle="Gruppe" yAxisTitle="Kroner" swapXY={true} />
                    </div>
                    {#if salaryGapRow}
                      <p class="mt-3 text-xs leading-5 text-slate-600">Les denne sammen med lønnstabellene. Overtid kan forklare deler av forskjellen i samlet kompensasjon, men ikke fastlønn.</p>
                    {/if}
                  </article>
                  {#if overtidMonthlyTotalChartData.length}
                    <article class="chart-card">
                      <h3>Utvikling gjennom året</h3>
                      <p class="mt-1 text-xs text-slate-600">Månedsbildet gjør det lett å se sesongtopper og perioder med ekstra belastning.</p>
                      <div
                        class="mt-4 cursor-zoom-in"
                        role="button"
                        tabindex="0"
                        aria-label="Åpne total overtid per måned i større visning"
                        on:click={() => openExpandedView("overtid-monthly-total-chart", "Utvikling gjennom året")}
                        on:keydown={(event) => handleExpandKeydown(event, "overtid-monthly-total-chart", "Utvikling gjennom året")}
                      >
                        <LineChart data={overtidMonthlyTotalChartData} x="måned" y="verdi" yFmt="#,##0" xAxisTitle="Måned" yAxisTitle="Kroner" />
                      </div>
                    </article>
                  {/if}
                </div>
              {/if}

              <details class="fastlonn-details">
                <summary class="cursor-pointer list-none px-4 py-3 text-sm font-medium text-[#0f5368]">Vis datagrunnlag</summary>
                <div class="space-y-3 border-t border-slate-200 p-4">
                  <p class="text-xs text-slate-600">
                    Viser hvilke overtidsfiler og hvilke ansatte som inngår etter filtrering på valgt år og valgt lønnsdato i fastlønn. Popupen viser alle rader og alle kolonner fra rådataene.
                  </p>
                  <div
                    class="overflow-x-auto cursor-zoom-in"
                    role="button"
                    tabindex="0"
                    aria-label="Åpne datagrunnlag for overtid i større visning"
                    on:click={() => openExpandedView("overtid-table", "Datagrunnlag for overtid")}
                    on:keydown={(event) => handleExpandKeydown(event, "overtid-table", "Datagrunnlag for overtid")}
                  >
                    <table class="min-w-full border-collapse text-left text-xs">
                      <thead>
                        <tr class="bg-[#eaf1f8] text-[#17365f]">
                          <th class="border px-2 py-1.5 font-semibold">Ansatt</th>
                          <th class="border px-2 py-1.5 font-semibold">Gruppe</th>
                          <th class="border px-2 py-1.5 font-semibold">Kjønn</th>
                          <th class="border px-2 py-1.5 font-semibold">Sum overtid</th>
                        </tr>
                      </thead>
                      <tbody>
                        {#each overtidParticipants.slice(0, 20) as employee}
                          <tr class="odd:bg-white even:bg-[#f7fafe]">
                            <td class="border px-2 py-1.5 align-top">{employee.name}</td>
                            <td class="border px-2 py-1.5 align-top">{employee.group}</td>
                            <td class="border px-2 py-1.5 align-top">{employee.gender}</td>
                            <td class="border px-2 py-1.5 align-top">{formatCurrency(employee.amount)}</td>
                          </tr>
                        {/each}
                      </tbody>
                    </table>
                  </div>
                </div>
              </details>
            </div>
          </section>
        {:else}
          <div class="content-card">
            <p class="panel-eyebrow">Overtid</p>
            <h2 class="panel-title">Ingen overtidsdata ennå</h2>
            <p class="section-note mt-2">Last opp grunnlagsfiler i fanen Datagrunnlag for å vise analysen.</p>
          </div>
        {/if}
      {/if}

      {#if activeTab === "vakttillegg"}
        <section class="space-y-3">
          <div class="content-card">
            <div class="fastlonn-hero-head">
              <div>
                <p class="panel-eyebrow">Vakttillegg</p>
                <h2 class="panel-title">Vakttillegg</h2>
                <p class="section-note mt-2">Vakttillegg sier ofte mer om rollefordeling og beredskap enn om lønnspolitikk alene.</p>
              </div>
              <div class="fastlonn-hero-meta">
                <div class="fastlonn-meta-pill">
                  <span>Total vakttillegg</span>
                  <strong>{formatCurrency(vakttilleggParticipants.reduce((sum, employee) => sum + employee.amount, 0))}</strong>
                </div>
                <div class="fastlonn-meta-pill">
                  <span>Ansatte med vakttillegg</span>
                  <strong>{formatInteger(vakttilleggParticipants.length)}</strong>
                </div>
              </div>
            </div>
          </div>
          {#if vakttilleggAnalysisRows.length}
            <div class="grid gap-3 xl:grid-cols-2">
              <article class="chart-card">
                <h3>Total vakttillegg per gruppe</h3>
                <p class="mt-1 text-xs text-slate-600">Viser hvor de samlede kostnadene faktisk ligger, ikke bare gjennomsnitt per deltaker.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne total vakttillegg per gruppe i større visning"
                  on:click={() => openExpandedView("vakttillegg-total-chart", "Total vakttillegg per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "vakttillegg-total-chart", "Total vakttillegg per gruppe")}
                >
                  <BarChart data={vakttilleggTotalChartData} x="gruppe" y="verdi" yFmt="#,##0" xAxisTitle="Gruppe" yAxisTitle="Kroner" swapXY={true} colorPalette={vakttilleggTotalChartPalette} seriesOptions={{ colorBy: "data" }} />
                </div>
              </article>

              <article class="chart-card">
                <h3>Deltakelsesgrad per gruppe</h3>
                <p class="mt-1 text-xs text-slate-600">Viser om vakttillegg er bredt fordelt eller bæres av få ansatte i hver gruppe.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne deltakelsesgrad for vakttillegg i større visning"
                  on:click={() => openExpandedView("vakttillegg-participation-chart", "Deltakelsesgrad per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "vakttillegg-participation-chart", "Deltakelsesgrad per gruppe")}
                >
                  <BarChart data={vakttilleggParticipationChartData} x="gruppe" y="andel" yFmt="pct1" xAxisTitle="Gruppe" yAxisTitle="Andel ansatte" swapXY={true} colorPalette={vakttilleggParticipationChartPalette} seriesOptions={{ colorBy: "data" }} />
                </div>
              </article>

              <article class="chart-card">
                <h3>Vakttillegg per gruppe</h3>
                <p class="mt-1 text-xs text-slate-600">Grafen gjør rollefordeling og forskjeller mellom kjønn synlige med mindre lesing.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne vakttillegg per gruppe i større visning"
                  on:click={() => openExpandedView("vakttillegg-chart", "Vakttillegg per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "vakttillegg-chart", "Vakttillegg per gruppe")}
                >
                  <BarChart data={vakttilleggChartData} x="gruppe" y="verdi" series="serie" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={paySeriesColors} yFmt="#,##0" xAxisTitle="Gruppe" yAxisTitle="Kroner" swapXY={true} />
                </div>
              </article>

              {#if vakttilleggSpreadRows.length}
                <article class="chart-card">
                  <h3>Spredning i vakttillegg</h3>
                  <p class="mt-1 text-xs text-slate-600">Viser variasjonen i vakttillegg innen hver gruppe, ikke bare totalsum eller gjennomsnitt.</p>
                  <div
                    class="mt-4 cursor-zoom-in"
                    role="button"
                    tabindex="0"
                    aria-label="Åpne spredning i vakttillegg i større visning"
                    on:click={() => openExpandedView("vakttillegg-spread-chart", "Spredning i vakttillegg")}
                    on:keydown={(event) => handleExpandKeydown(event, "vakttillegg-spread-chart", "Spredning i vakttillegg")}
                  >
                    <BoxPlot data={vakttilleggSpreadRows} name="gruppe" min="min" intervalBottom="q1" midpoint="median" intervalTop="q3" max="max" swapXY={true} yFmt="#,##0" yAxisTitle="Kroner" seriesColors={vakttilleggSpreadSeriesColors} />
                  </div>
                </article>
              {/if}
            </div>

            <details class="fastlonn-details">
              <summary class="cursor-pointer list-none px-4 py-3 text-sm font-medium text-[#0f5368]">Vis datagrunnlag</summary>
              <div class="space-y-3 border-t border-slate-200 p-4">
                <p class="text-xs text-slate-600">Viser hvilke vakttilleggsrader som inngår i valgt år og hvilke ansatte som er matchet mot fastlønn.</p>
                <div class="grid gap-2 md:grid-cols-3">
                  <article class="metric-card">
                    <p class="summary-card-label">Kilde-filer</p>
                    <p class="metric-card-value">{formatInteger(vakttilleggSourceFiles.length)}</p>
                  </article>
                  <article class="metric-card">
                    <p class="summary-card-label">Rader i år</p>
                    <p class="metric-card-value">{formatInteger(vakttilleggSourceRows.length)}</p>
                  </article>
                  <article class="metric-card">
                    <p class="summary-card-label">Ansatte matchet</p>
                    <p class="metric-card-value">{formatInteger(vakttilleggParticipants.length)}</p>
                  </article>
                </div>
                {#if vakttilleggSpreadRows.length}
                  <div
                    class="ledger-table-wrap cursor-zoom-in"
                    role="button"
                    tabindex="0"
                    aria-label="Åpne spredning i vakttillegg i større visning"
                    on:click={() => openExpandedView("vakttillegg-spread-table", "Spredning i vakttillegg")}
                    on:keydown={(event) => handleExpandKeydown(event, "vakttillegg-spread-table", "Spredning i vakttillegg")}
                  >
                    <table class="ledger-table ledger-table-dense">
                      <thead>
                        <tr>
                          <th>Gruppe</th>
                          <th class="text-right">Min</th>
                          <th class="text-right">Q1</th>
                          <th class="text-right">Median</th>
                          <th class="text-right">Q3</th>
                          <th class="text-right">Maks</th>
                        </tr>
                      </thead>
                      <tbody>
                        {#each vakttilleggSpreadRows as row}
                          <tr>
                            <td class="font-medium text-slate-900">{row.gruppe}</td>
                            <td class="text-right tabular-nums">{formatCurrency(row.min)}</td>
                            <td class="text-right tabular-nums">{formatCurrency(row.q1)}</td>
                            <td class="text-right tabular-nums">{formatCurrency(row.median)}</td>
                            <td class="text-right tabular-nums">{formatCurrency(row.q3)}</td>
                            <td class="text-right tabular-nums">{formatCurrency(row.max)}</td>
                          </tr>
                        {/each}
                      </tbody>
                    </table>
                  </div>
                {/if}
                <div
                  class="overflow-x-auto cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne datagrunnlag for vakttillegg i større visning"
                  on:click={() => openExpandedView("vakttillegg-table", "Datagrunnlag for vakttillegg")}
                  on:keydown={(event) => handleExpandKeydown(event, "vakttillegg-table", "Datagrunnlag for vakttillegg")}
                >
                  <table class="min-w-full border-collapse text-left text-xs">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        <th class="border px-2 py-1.5 font-semibold">Ansatt</th>
                        <th class="border px-2 py-1.5 font-semibold">Gruppe</th>
                        <th class="border px-2 py-1.5 font-semibold">Kjønn</th>
                        <th class="border px-2 py-1.5 font-semibold">Sum vakttillegg</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each vakttilleggParticipants.slice(0, 20) as employee}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          <td class="border px-2 py-1.5 align-top">{employee.name}</td>
                          <td class="border px-2 py-1.5 align-top">{employee.group}</td>
                          <td class="border px-2 py-1.5 align-top">{employee.gender}</td>
                          <td class="border px-2 py-1.5 align-top">{formatCurrency(employee.amount)}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                </div>
              </div>
            </details>
          {:else}
            <div class="content-card">
              <p class="section-note">Last opp rådatafilene i fanen Datagrunnlag for å vise vakttillegg.</p>
            </div>
          {/if}
        </section>
      {/if}

      {#if activeTab === "foreldrepermisjon"}
        <section class="space-y-3">
          <div class="content-card">
            <div class="fastlonn-hero-head">
              <div>
                <p class="panel-eyebrow">Foreldrepermisjon</p>
                <h2 class="panel-title">Foreldrepermisjon</h2>
                <p class="section-note mt-2">Her vises permisjonsuttak og arbeidsmønster som bør forklares i årsrapporten.</p>
              </div>
              <div class="fastlonn-hero-meta">
                <div class="fastlonn-meta-pill">
                  <span>Ansatte med uttak</span>
                  <strong>{formatInteger(report?.foreldrepermisjon?.totalEmployees ?? foreldrepermisjonEmployees.length)}</strong>
                </div>
                <div class="fastlonn-meta-pill">
                  <span>Snitt uker</span>
                  <strong>{formatNumber(
                    report?.foreldrepermisjon?.totalEmployees
                      ? ((report.foreldrepermisjon.womenAvgWeeks * report.foreldrepermisjon.womenCount) +
                          (report.foreldrepermisjon.menAvgWeeks * report.foreldrepermisjon.menCount)) /
                        report.foreldrepermisjon.totalEmployees
                      : 0
                  )}</strong>
                </div>
              </div>
            </div>
          </div>
          <div class="grid gap-3 xl:grid-cols-2">
            {#if foreldrepermisjonCountChartData.length}
              <article class="chart-card">
                <h3>Ansatte med uttak</h3>
                <p class="mt-1 text-xs text-slate-600">Viser hvor mange kvinner og menn som faktisk har registrert uttak i valgt år.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne ansatte med uttak i større visning"
                  on:click={() => openExpandedView("foreldrepermisjon-count-chart", "Ansatte med uttak")}
                  on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-count-chart", "Ansatte med uttak")}
                >
                  <BarChart data={foreldrepermisjonCountChartData} x="måling" y="verdi" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="num0" xAxisTitle="Måling" yAxisTitle="Antall ansatte" swapXY={true} />
                </div>
              </article>
            {/if}
            {#if foreldrepermisjonGroupCountChartData.length}
              <article class="chart-card">
                <h3>Uttak per gruppe</h3>
                <p class="mt-1 text-xs text-slate-600">Viser hvilke grupper som faktisk har permisjonsuttak, fordelt på kvinner og menn.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne uttak per gruppe i større visning"
                  on:click={() => openExpandedView("foreldrepermisjon-group-count-chart", "Uttak per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-group-count-chart", "Uttak per gruppe")}
                >
                  <BarChart data={foreldrepermisjonGroupCountChartData} x="gruppe" y="verdi" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="num0" xAxisTitle="Gruppe" yAxisTitle="Antall ansatte" swapXY={true} />
                </div>
              </article>
            {/if}
            {#if foreldrepermisjonGroupParticipationChartData.length}
              <article class="chart-card">
                <h3>Andel ansatte med uttak per gruppe</h3>
                <p class="mt-1 text-xs text-slate-600">Testvariant som viser andeler som intensitet, slik at relative forskjeller mellom gruppene blir lettere å skanne.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne andel ansatte med uttak per gruppe i større visning"
                  on:click={() => openExpandedView("foreldrepermisjon-group-participation-chart", "Andel ansatte med uttak per gruppe")}
                  on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-group-participation-chart", "Andel ansatte med uttak per gruppe")}
                >
                  <Heatmap data={foreldrepermisjonGroupParticipationHeatmapData} x="gruppe" y="måling" value="andel" colorScale={[graphColors.neutral, graphColors.lightPurple]} valueFmt="pct1" />
                </div>
              </article>
            {/if}
            {#if foreldrepermisjonWeeksChartData.length}
              <article class="chart-card">
                <h3>Gjennomsnittlig permisjonsuttak</h3>
                <p class="mt-1 text-xs text-slate-600">Her er graf best for å se forskjellen i uttak mellom kvinner og menn med én gang.</p>
                <div
                  class="mt-4 cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne gjennomsnittlig permisjonsuttak i større visning"
                  on:click={() => openExpandedView("foreldrepermisjon-weeks-chart", "Gjennomsnittlig permisjonsuttak")}
                  on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-weeks-chart", "Gjennomsnittlig permisjonsuttak")}
                >
                  <BarChart data={foreldrepermisjonWeeksChartData} x="måling" y="verdi" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="#,##0.0" xAxisTitle="Måling" yAxisTitle="Uker" swapXY={true} />
                </div>
              </article>
            {/if}
            <details class="fastlonn-details xl:col-span-2">
              <summary class="cursor-pointer list-none px-4 py-3 text-sm font-medium text-[#0f5368]">Vis datagrunnlag</summary>
              <div class="space-y-3 border-t border-slate-200 p-4">
                <article class="chart-card">
                  <h3>Nøkkeltall for foreldrepermisjon</h3>
                  <p class="mt-1 text-xs text-slate-600">Denne tabellen samler de viktigste tallene for ansatte, uttakstid og fordeling mellom kjønn.</p>
                  <div
                    class="ledger-table-wrap mt-4 cursor-zoom-in"
                    role="button"
                    tabindex="0"
                    aria-label="Åpne nøkkeltall for foreldrepermisjon i større visning"
                    on:click={() => openExpandedView("foreldrepermisjon-summary-table", "Nøkkeltall for foreldrepermisjon")}
                    on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-summary-table", "Nøkkeltall for foreldrepermisjon")}
                  >
                    <table class="ledger-table ledger-table-dense">
                      <thead>
                        <tr>
                          <th>Måling</th>
                          <th class="text-right">Kvinner</th>
                          <th class="text-right">Menn</th>
                          <th class="text-right">Totalt</th>
                        </tr>
                      </thead>
                      <tbody>
                        {#each foreldrepermisjonSummaryRows as row}
                          <tr>
                            <td class="font-medium text-slate-900">{row.måling}</td>
                            <td class="text-right tabular-nums">{row.måling === "Andel permisjonsdager" ? formatPercent(row.kvinner) : formatNumber(row.kvinner)}</td>
                            <td class="text-right tabular-nums">{row.måling === "Andel permisjonsdager" ? formatPercent(row.menn) : formatNumber(row.menn)}</td>
                            <td class="text-right tabular-nums">{row.måling === "Andel permisjonsdager" ? formatPercent(row.totalt) : formatNumber(row.totalt)}</td>
                          </tr>
                        {/each}
                      </tbody>
                    </table>
                  </div>
                </article>
                <p class="text-xs text-slate-600">Viser hvilke permisjonsrader som inngår i valgt år og et utvalg ansatte etter filtreringen.</p>
                <div class="grid gap-2 md:grid-cols-3">
                  <article class="metric-card">
                    <p class="summary-card-label">Kilde-filer</p>
                    <p class="metric-card-value">{formatInteger(foreldrepermisjonSourceFiles.length)}</p>
                  </article>
                  <article class="metric-card">
                    <p class="summary-card-label">Rader i år</p>
                    <p class="metric-card-value">{formatInteger(foreldrepermisjonSourceRows.length)}</p>
                  </article>
                  <article class="metric-card">
                    <p class="summary-card-label">Ansatte matchet</p>
                    <p class="metric-card-value">{formatInteger(foreldrepermisjonEmployees.length)}</p>
                  </article>
                </div>
                <div
                  class="overflow-x-auto cursor-zoom-in"
                  role="button"
                  tabindex="0"
                  aria-label="Åpne datagrunnlag for foreldrepermisjon i større visning"
                  on:click={() => openExpandedView("foreldrepermisjon-table", "Datagrunnlag for foreldrepermisjon")}
                  on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-table", "Datagrunnlag for foreldrepermisjon")}
                >
                  <table class="min-w-full border-collapse text-left text-xs">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        <th class="border px-2 py-1.5 font-semibold">Ansatt</th>
                        <th class="border px-2 py-1.5 font-semibold">Kjønn</th>
                        <th class="border px-2 py-1.5 font-semibold">Fraværsdager</th>
                        <th class="border px-2 py-1.5 font-semibold">Vektede dager</th>
                        <th class="border px-2 py-1.5 font-semibold">Uker</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each foreldrepermisjonEmployees.slice(0, 20) as employee}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          <td class="border px-2 py-1.5 align-top">{employee.name}</td>
                          <td class="border px-2 py-1.5 align-top">{employee.gender}</td>
                          <td class="border px-2 py-1.5 align-top">{formatNumber(employee.days)}</td>
                          <td class="border px-2 py-1.5 align-top">{formatNumber(employee.weightedDays)}</td>
                          <td class="border px-2 py-1.5 align-top">{formatNumber(employee.weeks)}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                </div>
                {#if foreldrepermisjonSpreadRows.length}
                  <div
                    class="ledger-table-wrap cursor-zoom-in"
                    role="button"
                    tabindex="0"
                    aria-label="Åpne spredning i permisjonsuttak i større visning"
                    on:click={() => openExpandedView("foreldrepermisjon-spread-table", "Spredning i permisjonsuttak")}
                    on:keydown={(event) => handleExpandKeydown(event, "foreldrepermisjon-spread-table", "Spredning i permisjonsuttak")}
                  >
                    <table class="ledger-table ledger-table-dense">
                      <thead>
                        <tr>
                          <th>Gruppe</th>
                          <th class="text-right">Min</th>
                          <th class="text-right">Q1</th>
                          <th class="text-right">Median</th>
                          <th class="text-right">Q3</th>
                          <th class="text-right">Maks</th>
                        </tr>
                      </thead>
                      <tbody>
                        {#each foreldrepermisjonSpreadRows as row}
                          <tr>
                            <td class="font-medium text-slate-900">{row.gruppe}</td>
                            <td class="text-right tabular-nums">{formatNumber(row.min)}</td>
                            <td class="text-right tabular-nums">{formatNumber(row.q1)}</td>
                            <td class="text-right tabular-nums">{formatNumber(row.median)}</td>
                            <td class="text-right tabular-nums">{formatNumber(row.q3)}</td>
                            <td class="text-right tabular-nums">{formatNumber(row.max)}</td>
                          </tr>
                        {/each}
                      </tbody>
                    </table>
                  </div>
                {/if}
              </div>
            </details>
          </div>
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
              <button class="app-button app-button-download" on:click={downloadReport} disabled={!report}>Last ned årsrapport (DOCX)</button>
            </div>
          </div>
          <div class="content-card">
            <p class="panel-eyebrow">Forhåndsutfylling</p>
            <h2 class="panel-title">Manuelle felt før eksport</h2>
            <p class="section-note mt-2">
              Verdiene under lagres lokalt i nettleseren for {selectedFastlonnYear || "gjeldende utvalg"} og settes rett inn i DOCX-filen ved eksport.
            </p>

            <div class="report-manual-grid mt-4">
              <label class="report-manual-field">
                <span class="report-manual-label">Antall nyansatte</span>
                <input class="report-manual-input" type="text" value={manualReportInputs.nyansatteCount} on:input={(event) => updateManualReportInput("nyansatteCount", event.currentTarget.value)} />
              </label>
              <label class="report-manual-field">
                <span class="report-manual-label">Samlet sykefravær %</span>
                <input class="report-manual-input" type="text" value={manualReportInputs.samletSykefravaer} on:input={(event) => updateManualReportInput("samletSykefravaer", event.currentTarget.value)} />
              </label>
              <label class="report-manual-field">
                <span class="report-manual-label">Sykefravær kvinner %</span>
                <input class="report-manual-input" type="text" value={manualReportInputs.sykefravaerKvinner} on:input={(event) => updateManualReportInput("sykefravaerKvinner", event.currentTarget.value)} />
              </label>
              <label class="report-manual-field">
                <span class="report-manual-label">Sykefravær menn %</span>
                <input class="report-manual-input" type="text" value={manualReportInputs.sykefravaerMenn} on:input={(event) => updateManualReportInput("sykefravaerMenn", event.currentTarget.value)} />
              </label>
            </div>
          </div>
          {#if reportHighlights.length}
            <div class="content-card">
              <p class="panel-eyebrow">Årsoversikt</p>
              <h2 class="panel-title">Hovedfunn for valgt grunnlag</h2>
              <p class="section-note mt-2">Dette er de viktigste signalene å løfte inn i årsrapporten for {selectedFastlonnYear || "valgt år"} og uttrekk {selectedFastlonnPeriodChoice?.label || "—"}.</p>
              <div class="report-highlights-grid mt-4">
                {#each reportHighlights as item}
                  <article class="summary-card summary-card-quiet">
                    <p class="summary-card-label">{item.title}</p>
                    <p class="summary-card-metric">{item.metric}</p>
                    <p class="summary-card-text">{item.detail}</p>
                  </article>
                {/each}
              </div>
            </div>
          {/if}
        </section>
      {/if}
    </main>
  </div>

  {#if expandedView}
    <div class="fixed inset-0 z-[80] flex items-center justify-center p-4" role="dialog" aria-modal="true" aria-label={expandedView.title}>
      <button class="absolute inset-0 bg-slate-950/60" type="button" aria-label="Lukk popup" on:click={closeExpandedView}></button>
      <div class="relative z-10 max-h-[92vh] w-[min(96vw,1400px)] overflow-y-auto rounded-2xl border border-slate-200 bg-white shadow-2xl">
        <div class="sticky top-0 z-10 flex items-start justify-between gap-4 border-b border-slate-200 bg-white px-5 py-4">
          <div>
            <p class="text-[11px] font-medium uppercase tracking-[0.18em] text-slate-500">Større visning</p>
            <h3 class="mt-1 text-lg font-semibold text-[#0f2747]">{expandedView.title}</h3>
            {#if expandedView.note}
              <p class="mt-1 text-sm text-slate-600">{expandedView.note}</p>
            {/if}
          </div>
          <button class="rounded-lg border border-slate-200 px-3 py-1.5 text-sm font-medium text-slate-700 hover:bg-slate-50" type="button" on:click={closeExpandedView}>
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
          {:else if expandedView.id === "fastlonn-representation-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table">
                <thead>
                  <tr>
                    <th>Gruppe</th>
                    <th class="text-right">Kvinner</th>
                    <th class="text-right">Menn</th>
                    <th class="text-right">Kvinneandel</th>
                    <th class="text-right">Mannandel</th>
                    <th class="text-right">Totalt</th>
                  </tr>
                </thead>
                <tbody>
                  {#each fastlonnRepresentationTableRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.gruppe}</td>
                      <td class="text-right tabular-nums">{formatInteger(row.kvinner)}</td>
                      <td class="text-right tabular-nums">{formatInteger(row.menn)}</td>
                      <td class="text-right tabular-nums">{row.kvinneandel}</td>
                      <td class="text-right tabular-nums">{row.mannandel}</td>
                      <td class="text-right tabular-nums">{formatInteger(row.totalt)}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "fastlonn-salary-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table">
                <thead>
                  <tr>
                    <th>Gruppe</th>
                    <th class="text-right">Ansatte</th>
                    <th class="text-right">Kvinner</th>
                    <th class="text-right">Menn</th>
                    <th class="text-right">Totalt</th>
                    <th class="text-right">Kvinner av menn</th>
                  </tr>
                </thead>
                <tbody>
                  {#each fastlonnSalaryTableRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.gruppe}</td>
                      <td class="text-right tabular-nums">{formatInteger(row.ansatte)}</td>
                      <td class="text-right tabular-nums">{row.kvinner}</td>
                      <td class="text-right tabular-nums">{row.menn}</td>
                      <td class="text-right tabular-nums">{row.totalt}</td>
                      <td class="text-right tabular-nums">{row.lønnsforhold}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "fastlonn-gender-chart"}
            <BarChart data={genderBalanceChartData} x="gruppe" y="antall" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={genderSeriesColors} yFmt="num0" xAxisTitle="Gruppe" yAxisTitle="Antall ansatte" />
          {:else if expandedView.id === "fastlonn-salary-spread-chart"}
            <BoxPlot data={salarySpreadRows} name="gruppe" min="min" intervalBottom="q1" midpoint="median" intervalTop="q3" max="max" swapXY={true} yFmt="#,##0" yAxisTitle="Kroner" seriesColors={salarySpreadSeriesColors} />
          {:else if expandedView.id === "fastlonn-spread-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table ledger-table-dense">
                <thead>
                  <tr>
                    <th>Gruppe</th>
                    <th class="text-right">Min</th>
                    <th class="text-right">Q1</th>
                    <th class="text-right">Median</th>
                    <th class="text-right">Q3</th>
                    <th class="text-right">Maks</th>
                  </tr>
                </thead>
                <tbody>
                  {#each fastlonnSpreadTableRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.gruppe}</td>
                      <td class="text-right tabular-nums">{row.min}</td>
                      <td class="text-right tabular-nums">{row.q1}</td>
                      <td class="text-right tabular-nums">{row.median}</td>
                      <td class="text-right tabular-nums">{row.q3}</td>
                      <td class="text-right tabular-nums">{row.maks}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "fastlonn-employment-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table ledger-table-dense">
                <thead>
                  <tr>
                    <th>Kategori</th>
                    <th>Kvinner</th>
                    <th>Menn</th>
                  </tr>
                </thead>
                <tbody>
                  {#each fastlonnEmploymentTableRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.kategori}</td>
                      <td>{row.kvinner}</td>
                      <td>{row.menn}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "fastlonn-source-summary-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table ledger-table-dense">
                <thead>
                  <tr>
                    <th>Fil</th>
                    <th>Snapshot</th>
                    <th>År</th>
                    <th class="text-right">Rader</th>
                  </tr>
                </thead>
                <tbody>
                  {#each fastlonnSourceSummaryRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.fil}</td>
                      <td>{row.snapshot}</td>
                      <td>{row.år}</td>
                      <td class="text-right tabular-nums">{row.rader}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "fastlonn-employment-heatmap"}
            <Heatmap data={employmentHeatmapData} x="måling" y="kjønn" value="andel" colorScale={[graphColors.neutral, graphColors.green]} valueFmt="pct1" />
          {:else if expandedView.id === "fastlonn-salary-scatter"}
            <ScatterPlot data={salaryScatterData} x="menn" y="kvinner" size="antall" series="gruppe" xFmt="#,##0" yFmt="#,##0" xAxisTitle="Menns gjennomsnittslønn" yAxisTitle="Kvinners gjennomsnittslønn" seriesColors={salaryScatterSeriesColors} />
          {:else if expandedView.id === "fastlonn-employee-preview-table"}
            <div class="space-y-4">
              <div class="snapshot-controls">
                <div class="snapshot-periods">
                  <button
                    type="button"
                    class={`snapshot-period ${fastlonnPopupColumnMode === "default" ? "snapshot-period-active" : ""}`}
                    on:click={() => (fastlonnPopupColumnMode = "default")}
                  >
                    Standardvisning
                  </button>
                  <button
                    type="button"
                    class={`snapshot-period ${fastlonnPopupColumnMode === "all" ? "snapshot-period-active" : ""}`}
                    on:click={() => (fastlonnPopupColumnMode = "all")}
                  >
                    Alle kolonner
                  </button>
                </div>
              </div>
              {#if fastlonnPopupColumnMode === "all"}
                <div class="popup-top-scroll" bind:this={fastlonnTopScroll} on:scroll={() => handleTopScrollbarScroll("fastlonn")}>
                  <div style={`width: ${fastlonnTopScrollWidth}px; height: 1px;`}></div>
                </div>
              {/if}
              <div class="overflow-x-auto max-w-full" bind:this={fastlonnBottomScroll} on:scroll={() => handleBottomScrollbarScroll("fastlonn")}>
                {#if fastlonnPopupColumnMode === "all"}
                  <table class="min-w-max border-collapse text-left text-sm whitespace-nowrap">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        {#each fastlonnSourceHeaders as header}
                          <th class="border px-3 py-2 font-semibold">{header}</th>
                        {/each}
                      </tr>
                    </thead>
                    <tbody>
                      {#each fastlonnPreviewRows as row}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          {#each fastlonnSourceHeaders as header}
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
                        <th class="border px-3 py-2 font-semibold">Lønn</th>
                        <th class="border px-3 py-2 font-semibold">Arbeidsforhold</th>
                        <th class="border px-3 py-2 font-semibold">Deltid %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each report.fastlonn.employees as employee}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          <td class="border px-3 py-2 align-top">{employee.name}</td>
                          <td class="border px-3 py-2 align-top">{employee.group}</td>
                          <td class="border px-3 py-2 align-top">{employee.gender}</td>
                          <td class="border px-3 py-2 align-top">{formatCurrency(employee.salary)}</td>
                          <td class="border px-3 py-2 align-top">{employee.employmentGroup}</td>
                          <td class="border px-3 py-2 align-top">{formatNumber(employee.partTimePercent)}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {/if}
              </div>
            </div>
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
          {:else if expandedView.id === "overtid-participation-chart"}
            <BarChart data={overtidParticipationChartData} x="gruppe" y="andel" yFmt="pct1" xAxisTitle="Gruppe" yAxisTitle="Andel ansatte" swapXY={true} colorPalette={overtidParticipationChartPalette} seriesOptions={{ colorBy: "data" }} />
          {:else if expandedView.id === "overtid-total-chart"}
            <BarChart data={overtidTotalChartData} x="gruppe" y="verdi_tusen" yFmt="#,##0.0" xAxisTitle="Gruppe" yAxisTitle="1000 NOK" swapXY={true} colorPalette={overtidTotalChartPalette} seriesOptions={{ colorBy: "data" }} />
          {:else if expandedView.id === "overtid-monthly-total-chart"}
            <LineChart data={overtidMonthlyTotalChartData} x="måned" y="verdi" yFmt="#,##0" xAxisTitle="Måned" yAxisTitle="Kroner" />
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
          {:else if expandedView.id === "vakttillegg-total-chart"}
            <BarChart data={vakttilleggTotalChartData} x="gruppe" y="verdi" yFmt="#,##0" xAxisTitle="Gruppe" yAxisTitle="Kroner" swapXY={true} colorPalette={vakttilleggTotalChartPalette} seriesOptions={{ colorBy: "data" }} />
          {:else if expandedView.id === "vakttillegg-participation-chart"}
            <BarChart data={vakttilleggParticipationChartData} x="gruppe" y="andel" yFmt="pct1" xAxisTitle="Gruppe" yAxisTitle="Andel ansatte" swapXY={true} colorPalette={vakttilleggParticipationChartPalette} seriesOptions={{ colorBy: "data" }} />
          {:else if expandedView.id === "vakttillegg-spread-chart"}
            <BoxPlot data={vakttilleggSpreadRows} name="gruppe" min="min" intervalBottom="q1" midpoint="median" intervalTop="q3" max="max" swapXY={true} yFmt="#,##0" yAxisTitle="Kroner" seriesColors={vakttilleggSpreadSeriesColors} />
          {:else if expandedView.id === "vakttillegg-spread-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table ledger-table-dense">
                <thead>
                  <tr>
                    <th>Gruppe</th>
                    <th class="text-right">Min</th>
                    <th class="text-right">Q1</th>
                    <th class="text-right">Median</th>
                    <th class="text-right">Q3</th>
                    <th class="text-right">Maks</th>
                  </tr>
                </thead>
                <tbody>
                  {#each vakttilleggSpreadRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.gruppe}</td>
                      <td class="text-right tabular-nums">{formatCurrency(row.min)}</td>
                      <td class="text-right tabular-nums">{formatCurrency(row.q1)}</td>
                      <td class="text-right tabular-nums">{formatCurrency(row.median)}</td>
                      <td class="text-right tabular-nums">{formatCurrency(row.q3)}</td>
                      <td class="text-right tabular-nums">{formatCurrency(row.max)}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "foreldrepermisjon-table"}
            <div class="space-y-4">
              <div class="snapshot-controls">
                <div class="snapshot-periods">
                  <button
                    type="button"
                    class={`snapshot-period ${foreldrepermisjonPopupColumnMode === "default" ? "snapshot-period-active" : ""}`}
                    on:click={() => (foreldrepermisjonPopupColumnMode = "default")}
                  >
                    Standardvisning
                  </button>
                  <button
                    type="button"
                    class={`snapshot-period ${foreldrepermisjonPopupColumnMode === "all" ? "snapshot-period-active" : ""}`}
                    on:click={() => (foreldrepermisjonPopupColumnMode = "all")}
                  >
                    Alle kolonner
                  </button>
                </div>
              </div>
              {#if foreldrepermisjonPopupColumnMode === "all"}
                <div class="popup-top-scroll" bind:this={foreldrepermisjonTopScroll} on:scroll={() => handleTopScrollbarScroll("foreldrepermisjon")}>
                  <div style={`width: ${foreldrepermisjonTopScrollWidth}px; height: 1px;`}></div>
                </div>
              {/if}
              <div class="overflow-x-auto max-w-full" bind:this={foreldrepermisjonBottomScroll} on:scroll={() => handleBottomScrollbarScroll("foreldrepermisjon")}>
                {#if foreldrepermisjonPopupColumnMode === "all"}
                  <table class="min-w-max border-collapse text-left text-sm whitespace-nowrap">
                    <thead>
                      <tr class="bg-[#eaf1f8] text-[#17365f]">
                        {#each foreldrepermisjonSourceHeaders as header}
                          <th class="border px-3 py-2 font-semibold">{header}</th>
                        {/each}
                      </tr>
                    </thead>
                    <tbody>
                      {#each foreldrepermisjonSourceRows as row}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          {#each foreldrepermisjonSourceHeaders as header}
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
                        <th class="border px-3 py-2 font-semibold">Kjønn</th>
                        <th class="border px-3 py-2 font-semibold">Fraværsdager</th>
                        <th class="border px-3 py-2 font-semibold">Vektede dager</th>
                        <th class="border px-3 py-2 font-semibold">Uker</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each foreldrepermisjonEmployees as employee}
                        <tr class="odd:bg-white even:bg-[#f7fafe]">
                          <td class="border px-3 py-2 align-top">{employee.name}</td>
                          <td class="border px-3 py-2 align-top">{employee.gender}</td>
                          <td class="border px-3 py-2 align-top">{formatNumber(employee.days)}</td>
                          <td class="border px-3 py-2 align-top">{formatNumber(employee.weightedDays)}</td>
                          <td class="border px-3 py-2 align-top">{formatNumber(employee.weeks)}</td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {/if}
              </div>
            </div>
          {:else if expandedView.id === "foreldrepermisjon-summary-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table ledger-table-dense">
                <thead>
                  <tr>
                    <th>Måling</th>
                    <th class="text-right">Kvinner</th>
                    <th class="text-right">Menn</th>
                    <th class="text-right">Totalt</th>
                  </tr>
                </thead>
                <tbody>
                  {#each foreldrepermisjonSummaryRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.måling}</td>
                      <td class="text-right tabular-nums">{row.måling === "Andel permisjonsdager" ? formatPercent(row.kvinner) : formatNumber(row.kvinner)}</td>
                      <td class="text-right tabular-nums">{row.måling === "Andel permisjonsdager" ? formatPercent(row.menn) : formatNumber(row.menn)}</td>
                      <td class="text-right tabular-nums">{row.måling === "Andel permisjonsdager" ? formatPercent(row.totalt) : formatNumber(row.totalt)}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "foreldrepermisjon-count-chart"}
            <BarChart data={foreldrepermisjonCountChartData} x="måling" y="verdi" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="num0" xAxisTitle="Måling" yAxisTitle="Antall ansatte" swapXY={true} />
          {:else if expandedView.id === "foreldrepermisjon-group-count-chart"}
            <BarChart data={foreldrepermisjonGroupCountChartData} x="gruppe" y="verdi" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="num0" xAxisTitle="Gruppe" yAxisTitle="Antall ansatte" swapXY={true} />
          {:else if expandedView.id === "foreldrepermisjon-group-participation-chart"}
            <Heatmap data={foreldrepermisjonGroupParticipationHeatmapData} x="gruppe" y="måling" value="andel" colorScale={[graphColors.neutral, graphColors.lightPurple]} valueFmt="pct1" />
          {:else if expandedView.id === "foreldrepermisjon-spread-table"}
            <div class="ledger-table-wrap">
              <table class="ledger-table ledger-table-dense">
                <thead>
                  <tr>
                    <th>Gruppe</th>
                    <th class="text-right">Min</th>
                    <th class="text-right">Q1</th>
                    <th class="text-right">Median</th>
                    <th class="text-right">Q3</th>
                    <th class="text-right">Maks</th>
                  </tr>
                </thead>
                <tbody>
                  {#each foreldrepermisjonSpreadRows as row}
                    <tr>
                      <td class="font-medium text-slate-900">{row.gruppe}</td>
                      <td class="text-right tabular-nums">{formatNumber(row.min)}</td>
                      <td class="text-right tabular-nums">{formatNumber(row.q1)}</td>
                      <td class="text-right tabular-nums">{formatNumber(row.median)}</td>
                      <td class="text-right tabular-nums">{formatNumber(row.q3)}</td>
                      <td class="text-right tabular-nums">{formatNumber(row.max)}</td>
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else if expandedView.id === "foreldrepermisjon-weeks-chart"}
            <BarChart data={foreldrepermisjonWeeksChartData} x="måling" y="verdi" series="kjønn" type="grouped" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="#,##0.0" xAxisTitle="Måling" yAxisTitle="Uker" swapXY={true} />
          {:else if expandedView.id === "foreldrepermisjon-share-chart"}
            <BarChart data={foreldrepermisjonShareChartData} x="måling" y="verdi" series="kjønn" type="stacked100" seriesOrder={["Kvinner", "Menn"]} seriesColors={leaveSeriesColors} yFmt="pct1" xAxisTitle="Måling" yAxisTitle="Andel" />
          {:else if expandedView.id === "foreldrepermisjon-employment-heatmap"}
            <Heatmap data={employmentHeatmapData} x="måling" y="kjønn" value="andel" colorScale={[graphColors.neutral, graphColors.green]} valueFmt="pct1" />
          {/if}
        </div>
      </div>
    </div>
  {/if}
</section>
