import { AlignmentType, BorderStyle, Document, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType, HeadingLevel, TableLayoutType, VerticalAlign } from 'docx';
import { formatNumber, formatPercent } from './formatting.js';
import { standardReportNarratives, reportHistoryYears, normalizeManualReportInputs, validateManualField } from './report-text.js';

export const reportTableBorderColor = "D9D9D9";
const contentWidth = 9638;

export const reportTableHeaderFill = "EAF1FB";

export function reportTableBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor },
    left: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor },
    right: { style: BorderStyle.SINGLE, size: 4, color: reportTableBorderColor }
  };
}

export function reportTableCell(text, options = {}) {
  const {
    bold = false,
    align = AlignmentType.LEFT,
    shaded = false,
    rowSpan,
    columnSpan,
    width
  } = options;

  return new TableCell({
    ...(width ? { width: { size: width, type: WidthType.DXA } } : {}),
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 110, right: 110 },
    ...(rowSpan ? { rowSpan } : {}),
    ...(columnSpan ? { columnSpan } : {}),
    shading: shaded ? { type: ShadingType.CLEAR, color: "auto", fill: reportTableHeaderFill } : undefined,
    borders: reportTableBorders(),
    children: [
      new Paragraph({
        alignment: align,
        keepNext: bold,
        spacing: { before: 0, after: 0, line: 240 },
        children: [
          new TextRun({
            text: String(text),
            bold,
            color: "000000",
            font: 'Arial',
            size: 19
          })
        ]
      })
    ]
  });
}

export function tableFromRows(rows) {
  const count = rows[0].length;
  const first = rows[0][0];
  const weights = first === 'Fil' ? [34, 21, 10, 15, 9, 11]
    : first === 'Stillingsgruppe' ? [32, 11, 14, 11, 14, 18]
    : count === 6 && first === 'Gruppe' ? [28, 8, 16, 16, 16, 16]
    : [count === 3 ? 50 : 40, ...Array(count - 1).fill((count === 3 ? 50 : 60) / (count - 1))];
  const widths = weights.map((weight) => Math.floor(contentWidth * weight / 100));
  widths[widths.length - 1] += contentWidth - widths.reduce((sum, width) => sum + width, 0);
  return new Table({
    width: { size: contentWidth, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    rows: rows.map((row, rowIndex) =>
      new TableRow({
        tableHeader: rowIndex === 0,
        cantSplit: true,
        children: row.map((cell, column) =>
          reportTableCell(cell, {
            width: widths[column],
            bold: rowIndex === 0,
            align: rowIndex === 0 ? AlignmentType.CENTER : column === 0 || first === 'Fil' ? AlignmentType.LEFT : AlignmentType.RIGHT,
            shaded: rowIndex === 0
          })
        )
      })
    )
  });
}

export function emptyLine() {
  return new Paragraph({
    text: "",
    spacing: {
      after: 260
    }
  });
}

export function promptParagraph(text) {
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

export function noteParagraph(text) {
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

export function standardParagraph(text) {
  return new Paragraph({
    spacing: {
      before: 120,
      after: 180
    },
    children: [new TextRun({ text })]
  });
}

export function manualNoteBlock(prompt) {
  return [
    promptParagraph(`Manuell vurdering/notat: ${prompt}`),
    emptyLine()
  ];
}

export function narrativeBlock(title, prompt) {
  return [
    new Paragraph({ text: title, heading: HeadingLevel.HEADING_2 }),
    promptParagraph(prompt),
    emptyLine()
  ];
}

export function genderBalanceLabel(group) {
  if (group === "Totalt antall ansatte") return "Totalt antall ansatte *";
  if (group === "Direktørgruppen") return "Direktørgruppen**";
  return group;
}

export function genderBalanceTable(rows) {
  return tableFromRows([
    ['Stillingsgruppe', 'Kvinner antall', 'Andel kvinner', 'Menn antall', 'Andel menn', 'Totalt'],
    ...rows.map((row) => [row.group, row.women, formatPercent(row.womenShare), row.men, formatPercent(row.menShare), row.total])
  ]);
}

export function createReportDocument(summary, { year: selectedFastlonnYear, files: uploadedFiles = [], manualInputs = {} } = {}) {
  if (!summary?.fastlonn) throw new Error('Last opp et gyldig fastlønnsuttrekk før eksport.');
  const manualReportInputs = normalizeManualReportInputs(manualInputs);
  for (const [field, value] of Object.entries(manualReportInputs)) {
    if (value.trim() && validateManualField(field, value)) throw new Error('Rett de manuelle tallfeltene før eksport.');
  }
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
    promptParagraph("Rapportutkast. Kontroller merknadene og erstatt grå hjelpetekst med faglige vurderinger før rapporten ferdigstilles."),
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
      ["Totalt antall ansettelser", ...historyYears.map(() => "")],
      ["Ansettelser i målgruppen - totalt", ...historyYears.map(() => "")],
      ["Faste ansettelser", ...historyYears.map(() => "")],
      ["Midlertidige administrative ansettelser", ...historyYears.map(() => "")],
      ["Traineeprogrammet i staten", ...historyYears.map(() => "")],
      ["Midlertidig m/ lønnstilskudd", ...historyYears.map(() => "")],
      ["Overgang fra midlertidig til fast (3-årsregelen)", ...historyYears.map(() => "")],
      ["Prosentoppnåelse", ...historyYears.map(() => "")],
      ["Totalt", ...historyYears.map(() => "")],
      ["Praksisplasser / arbeidstrening (ikke med i tallene)", ...historyYears.map(() => "")]
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
      text: `Kartlegging av kjønnsbalanse, lønn og deltid${summary.fastlonnSnapshotLabel ? ` (fastlønn per ${summary.fastlonnSnapshotLabel})` : ""}`,
      heading: HeadingLevel.HEADING_1
    })
  ];

  if (summary.fastlonn) {
    children.push(new Paragraph({ text: "KJØNNSBALANSE", heading: HeadingLevel.HEADING_2 }));
    children.push(genderBalanceTable(summary.fastlonn.genderBalance));
    children.push(
      noteParagraph("Direktørgruppen omfatter direktør, avdelingsdirektører og underdirektører med oppgitt personalansvar. Uavklarte grupper og eksplisitte avgrensninger fremgår av merknadene."),
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
        ["Kategori", "Kvinner antall", "Andel kvinner", "Menn antall", "Andel menn"],
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

  if (summary.notes?.length) {
    children.push(new Paragraph({ text: "Merknader", heading: HeadingLevel.HEADING_1 }));
    for (const note of summary.notes) {
      children.push(new Paragraph({ text: note }));
    }
  }

  const doc = new Document({
    creator: 'HR Årsrapport',
    title: 'HR Årsrapport ' + reportYear,
    description: 'Redigerbart rapportutkast med beregnede tabeller og manuelle vurderingsfelt.',
    styles: { default: {
      document: { run: { font: 'Arial', size: 21, color: '000000', language: { value: 'nb-NO' } }, paragraph: { spacing: { after: 120, line: 276 } } },
      title: { run: { font: 'Arial', size: 42, color: '000000', bold: true }, paragraph: { spacing: { before: 0, after: 200 }, keepNext: true } },
      heading1: { run: { font: 'Arial', size: 30, color: '000000', bold: true }, paragraph: { spacing: { before: 280, after: 140 }, keepNext: true } },
      heading2: { run: { font: 'Arial', size: 24, color: '000000', bold: true }, paragraph: { spacing: { before: 220, after: 100 }, keepNext: true } }
    } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } }, children }]
  });
  return doc;
}
