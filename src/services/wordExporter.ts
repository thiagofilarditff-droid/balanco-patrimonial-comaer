import {
  AlignmentType,
  BorderStyle,
  Document,
  HeightRule,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
  convertMillimetersToTwip,
} from 'docx';
import { BalancoData, StatementItem, WordExportConfig } from '../types/balanco';
import { formatThousands } from '../utils/numberFormat';

export const DEFAULT_WORD_CONFIG: WordExportConfig = {
  fontName: 'Times New Roman',
  fontSizePt: 10.5,
  marginsCm: 1.5,
  primaryColorHex: '285F98', // Azul oficial quadro principal
  secondaryColorHex: '31549A', // Azul oficial quadros complementares
  headerTextColorHex: 'FFFFFF',
  separatePages: true,
  showFonteSiafi: true,
};

/**
 * Creates a clean table cell without vertical borders and fine padding
 */
function createStyledCell(options: {
  text?: string;
  runs?: TextRun[];
  widthPct: number;
  align?: (typeof AlignmentType)[keyof typeof AlignmentType];
  bold?: boolean;
  underline?: boolean;
  italic?: boolean;
  color?: string;
  bgColor?: string;
  indentTwips?: number;
  fontSizePt?: number;
  fontName?: string;
  isBottomTotal?: boolean;
  isHeader?: boolean;
}): TableCell {
  const font = options.fontName || 'Times New Roman';
  const sizeHalfPt = Math.round((options.fontSizePt || 10.5) * 2);
  const align = options.align || AlignmentType.LEFT;

  let childrenRuns: TextRun[] = [];

  if (options.runs && options.runs.length > 0) {
    childrenRuns = options.runs;
  } else if (options.text !== undefined) {
    childrenRuns = [
      new TextRun({
        text: options.text,
        font,
        size: sizeHalfPt,
        bold: options.bold || options.isHeader || false,
        italics: options.italic || false,
        underline: options.underline ? { type: UnderlineType.SINGLE } : undefined,
        color: options.color || (options.isHeader || options.isBottomTotal ? 'FFFFFF' : '000000'),
      }),
    ];
  }

  const paragraph = new Paragraph({
    alignment: align,
    indent: options.indentTwips ? { left: options.indentTwips } : undefined,
    spacing: {
      before: 15,
      after: 15,
      line: 220,
    },
    children: childrenRuns,
  });

  const noBorder = {
    style: BorderStyle.NONE,
    size: 0,
    color: 'FFFFFF',
  };

  return new TableCell({
    width: {
      size: options.widthPct,
      type: WidthType.PERCENTAGE,
    },
    shading: options.bgColor
      ? {
          type: ShadingType.CLEAR,
          fill: options.bgColor,
        }
      : undefined,
    margins: {
      top: 40,
      bottom: 40,
      left: 60,
      right: 60,
    },
    borders: {
      top: noBorder,
      bottom: noBorder,
      left: noBorder,
      right: noBorder,
    },
    children: [paragraph],
  });
}

/**
 * Builds table rows for 4-column statements
 */
function buildFourColumnTable(
  titleLabel: string,
  items: StatementItem[],
  currYear: string,
  prevYear: string,
  config: WordExportConfig,
  themeColorHex: string,
  includeBottomBar = false
): Table {
  const tableRows: TableRow[] = [];
  const primaryBg = themeColorHex;

  // Header Row
  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: [
      createStyledCell({
        text: titleLabel,
        widthPct: 64,
        align: AlignmentType.LEFT,
        bgColor: primaryBg,
        isHeader: true,
        fontSizePt: config.fontSizePt,
        fontName: config.fontName,
      }),
      createStyledCell({
        text: 'NE',
        widthPct: 8,
        align: AlignmentType.CENTER,
        bgColor: primaryBg,
        isHeader: true,
        fontSizePt: config.fontSizePt,
        fontName: config.fontName,
      }),
      createStyledCell({
        text: currYear,
        widthPct: 14,
        align: AlignmentType.RIGHT,
        bgColor: primaryBg,
        isHeader: true,
        fontSizePt: config.fontSizePt,
        fontName: config.fontName,
      }),
      createStyledCell({
        text: prevYear,
        widthPct: 14,
        align: AlignmentType.RIGHT,
        bgColor: primaryBg,
        isHeader: true,
        fontSizePt: config.fontSizePt,
        fontName: config.fontName,
      }),
    ],
  });
  tableRows.push(headerRow);

  // Content Rows (Uniform white background, no alternating gray)
  items.forEach(item => {
    if (
      item.omitIfZeroBalance &&
      (item.valCurrent === null || item.valCurrent === 0) &&
      (item.valPrevious === null || item.valPrevious === 0)
    ) {
      return;
    }

    const isTotal = !!item.isTotal;
    const isBold = !!item.isBold || isTotal;
    const isUnderline = !!item.isUnderline && !isTotal;
    const isUnderlineVal = !!item.isUnderlineValue && !isTotal;
    const indent = isTotal ? 0 : item.level * 160;
    const rowBg = isTotal ? primaryBg : undefined;
    const textColor = isTotal ? 'FFFFFF' : '000000';

    // NE Cell Run
    let neRuns: TextRun[] = [];
    if (!isTotal && item.neValue !== null && item.neValue !== undefined && String(item.neValue).trim() !== '') {
      neRuns = [
        new TextRun({
          text: String(item.neValue),
          font: config.fontName,
          size: Math.round(config.fontSizePt * 2),
          color: '0563C1', // Blue link color per images
          underline: { type: UnderlineType.SINGLE },
        }),
      ];
    }

    const valCurrFormatted = item.isUnmapped
      ? 'Pendente de mapeamento'
      : formatThousands(item.valCurrent);

    const valPrevFormatted = item.isUnmapped
      ? 'Pendente de mapeamento'
      : formatThousands(item.valPrevious);

    const row = new TableRow({
      cantSplit: true,
      children: [
        createStyledCell({
          text: item.label,
          widthPct: 64,
          align: AlignmentType.LEFT,
          bold: isBold,
          underline: isUnderline,
          indentTwips: indent,
          bgColor: rowBg,
          color: textColor,
          isBottomTotal: isTotal,
          fontSizePt: config.fontSizePt,
          fontName: config.fontName,
        }),
        createStyledCell({
          runs: neRuns,
          widthPct: 8,
          align: AlignmentType.CENTER,
          bgColor: rowBg,
          isBottomTotal: isTotal,
          fontSizePt: config.fontSizePt,
          fontName: config.fontName,
        }),
        createStyledCell({
          text: valCurrFormatted,
          widthPct: 14,
          align: AlignmentType.RIGHT,
          bold: isBold,
          underline: isUnderlineVal,
          bgColor: rowBg,
          color: textColor,
          isBottomTotal: isTotal,
          fontSizePt: config.fontSizePt,
          fontName: config.fontName,
        }),
        createStyledCell({
          text: valPrevFormatted,
          widthPct: 14,
          align: AlignmentType.RIGHT,
          bold: isBold,
          underline: isUnderlineVal,
          bgColor: rowBg,
          color: textColor,
          isBottomTotal: isTotal,
          fontSizePt: config.fontSizePt,
          fontName: config.fontName,
        }),
      ],
    });

    tableRows.push(row);
  });

  // Faixa azul inferior adicional se solicitada (ex: Quadro de Compensações)
  if (includeBottomBar) {
    const bottomBarRow = new TableRow({
      cantSplit: true,
      children: [
        createStyledCell({
          text: '',
          widthPct: 64,
          bgColor: primaryBg,
          fontSizePt: 4,
          fontName: config.fontName,
        }),
        createStyledCell({
          text: '',
          widthPct: 8,
          bgColor: primaryBg,
          fontSizePt: 4,
          fontName: config.fontName,
        }),
        createStyledCell({
          text: '',
          widthPct: 14,
          bgColor: primaryBg,
          fontSizePt: 4,
          fontName: config.fontName,
        }),
        createStyledCell({
          text: '',
          widthPct: 14,
          bgColor: primaryBg,
          fontSizePt: 4,
          fontName: config.fontName,
        }),
      ],
    });
    tableRows.push(bottomBarRow);
  }

  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: tableRows,
  });
}

/**
 * Builds table rows for Quadro 5: Demonstrativo do Superávit/Déficit Financeiro (1 exercício)
 */
function buildSuperavitDeficitTable(
  items: StatementItem[],
  currYear: string,
  config: WordExportConfig,
  themeColorHex: string
): Table {
  const tableRows: TableRow[] = [];
  const primaryBg = themeColorHex;

  // Header Row: Célula de descrição vazia, sem "ESPECIFICAÇÃO"
  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: [
      createStyledCell({
        text: '',
        widthPct: 75,
        align: AlignmentType.LEFT,
        bgColor: primaryBg,
        isHeader: true,
        fontSizePt: config.fontSizePt,
        fontName: config.fontName,
      }),
      createStyledCell({
        text: currYear,
        widthPct: 25,
        align: AlignmentType.RIGHT,
        bgColor: primaryBg,
        isHeader: true,
        fontSizePt: config.fontSizePt,
        fontName: config.fontName,
      }),
    ],
  });
  tableRows.push(headerRow);

  // Content Rows
  items.forEach(item => {
    const isTotal = !!item.isTotal;
    const isBold = !!item.isBold || isTotal;
    const indent = isTotal ? 0 : item.level * 180;
    const rowBg = isTotal ? primaryBg : undefined;
    const textColor = isTotal ? 'FFFFFF' : '000000';

    const valCurrFormatted = item.isUnmapped
      ? 'Pendente de mapeamento'
      : formatThousands(item.valCurrent);

    const row = new TableRow({
      cantSplit: true,
      children: [
        createStyledCell({
          text: item.label,
          widthPct: 75,
          align: AlignmentType.LEFT,
          bold: isBold,
          indentTwips: indent,
          bgColor: rowBg,
          color: textColor,
          isBottomTotal: isTotal,
          fontSizePt: config.fontSizePt,
          fontName: config.fontName,
        }),
        createStyledCell({
          text: valCurrFormatted,
          widthPct: 25,
          align: AlignmentType.RIGHT,
          bold: isBold,
          bgColor: rowBg,
          color: textColor,
          isBottomTotal: isTotal,
          fontSizePt: config.fontSizePt,
          fontName: config.fontName,
        }),
      ],
    });

    tableRows.push(row);
  });

  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: tableRows,
  });
}

function createUnitParagraph(config: WordExportConfig): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 60, after: 40 },
    children: [
      new TextRun({
        text: 'R$ milhares',
        font: config.fontName,
        size: Math.round((config.fontSizePt - 1.5) * 2),
        italics: false,
      }),
    ],
  });
}

function createFonteParagraph(config: WordExportConfig): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 40, after: 140 },
    children: [
      new TextRun({
        text: 'Fonte: SIAFI.',
        font: config.fontName,
        size: Math.round((config.fontSizePt - 1.5) * 2),
        italics: false,
      }),
    ],
  });
}

function createTitleParagraph(title: string, config: WordExportConfig): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 40 },
    children: [
      new TextRun({
        text: title,
        font: config.fontName,
        size: Math.round((config.fontSizePt + 2) * 2),
        bold: true,
      }),
    ],
  });
}

/**
 * Main function that compiles the entire Balanço Patrimonial into an authentic, editable .docx binary
 */
export function buildBalancoSections(
  data: BalancoData,
  config?: Partial<WordExportConfig>
) {
  const mergedConfig: WordExportConfig = {
    ...DEFAULT_WORD_CONFIG,
    ...(config || {}),
  };
  const currYear = data.metadata.colCurrent || '2026';
  const prevYear = data.metadata.colPrevious || '2025';
  const marginsCm = mergedConfig.marginsCm ?? 1.5;
  const marginTwips = convertMillimetersToTwip(marginsCm * 10);
  const primaryBlue = mergedConfig.primaryColorHex || '285F98';
  const secondaryBlue = mergedConfig.secondaryColorHex || '31549A';

  // SECTION 1: ATIVO (Page 1)
  const childrenP1 = [
    createTitleParagraph('Balanço Patrimonial', mergedConfig),
    createUnitParagraph(mergedConfig),
    buildFourColumnTable('ATIVO', data.ativo, currYear, prevYear, mergedConfig, primaryBlue, false),
    createFonteParagraph(mergedConfig),
  ];

  // SECTION 2: PASSIVO E PATRIMÔNIO LÍQUIDO (Page 2)
  const childrenP2 = [
    createUnitParagraph(mergedConfig),
    buildFourColumnTable('PASSIVO E PATRIMÔNIO LÍQUIDO', data.passivoPL, currYear, prevYear, mergedConfig, primaryBlue, false),
    createFonteParagraph(mergedConfig),
  ];

  // SECTION 3: QUADROS COMPLEMENTARES (Page 3)
  const childrenP3 = [
    // Quadro 3: Ativos e Passivos Financeiros e Permanente
    createTitleParagraph('Quadro de Ativos e Passivos Financeiros e Permanente', mergedConfig),
    createUnitParagraph(mergedConfig),
    buildFourColumnTable('', data.financeiroPermanente, currYear, prevYear, mergedConfig, secondaryBlue, false),
    createFonteParagraph(mergedConfig),

    // Quadro 4: Quadro de Compensações (com faixa azul inferior antes de Fonte SIAFI)
    createTitleParagraph('Quadro de Compensações', mergedConfig),
    createUnitParagraph(mergedConfig),
    buildFourColumnTable('', data.compensacoes, currYear, prevYear, mergedConfig, secondaryBlue, true),
    createFonteParagraph(mergedConfig),

    // Quadro 5: Superávit / Déficit Financeiro
    createTitleParagraph('Demonstrativo do Superávit / Déficit Financeiro Apurado no Balanço Patrimonial', mergedConfig),
    createUnitParagraph(mergedConfig),
    buildSuperavitDeficitTable(data.superavitDeficit, currYear, mergedConfig, secondaryBlue),
    createFonteParagraph(mergedConfig),
  ];

  return [
      {
        properties: {
          page: {
            margin: {
              top: marginTwips,
              bottom: marginTwips,
              left: marginTwips,
              right: marginTwips,
            },
          },
        },
        children: childrenP1,
      },
      {
        properties: {
          page: {
            margin: {
              top: marginTwips,
              bottom: marginTwips,
              left: marginTwips,
              right: marginTwips,
            },
          },
        },
        children: childrenP2,
      },
      {
        properties: {
          page: {
            margin: {
              top: marginTwips,
              bottom: marginTwips,
              left: marginTwips,
              right: marginTwips,
            },
          },
        },
        children: childrenP3,
      },
    ];
}

export async function generateBalancoWordDocument(data: BalancoData, config?: Partial<WordExportConfig>): Promise<Blob> {
 return Packer.toBlob(new Document({sections:buildBalancoSections(data,config)}));
}

