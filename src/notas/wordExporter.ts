import { TableData } from './types';
import { formatDateDDMMYYYY } from './utils/parser';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow as DocxTableRow, 
  TableCell, 
  WidthType, 
  BorderStyle, 
  AlignmentType, 
  HeadingLevel,
  VerticalAlign,
  ImageRun
} from 'docx';

  const formatBRL = (val: number) => {
    const rounded = Math.round((val || 0) / 1000);
    const formatted = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(Math.abs(rounded));
    return rounded < 0 ? `(${formatted})` : formatted;
  };

  const formatPercentWord = (val: number, isTotal100 = false) => {
    if (isTotal100 && Math.abs((val || 0) - 100) < 0.01) return '100';
    const absVal = Math.abs(val || 0).toFixed(2).replace('.', ',');
    return (val || 0) < 0 ? `(${absVal})` : absVal;
  };


export function buildNotesChildren(reports: TableData[]) {
      const allElements: (Paragraph | Table)[] = [];

      for (let idx = 0; idx < reports.length; idx++) {
        const report = reports[idx];
        
        // Note Title: Times New Roman, 12, Bold, Left Aligned
        if (report.noteTitle) {
          allElements.push(new Paragraph({
            children: [new TextRun({ text: report.noteTitle, bold: true, size: 24, font: "Times New Roman" })],
            alignment: AlignmentType.LEFT,
            spacing: { before: 400, after: 200 }
          }));
        }

        // Pre-Commentary text (if any)
        if (report.preCommentary) {
          allElements.push(new Paragraph({
            children: [new TextRun({ text: report.preCommentary, font: "Times New Roman", size: 24 })],
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 200 }
          }));
        }

        // Table Title: Above table, Times New Roman, 12, Bold, Left Aligned
        allElements.push(new Paragraph({
          children: [new TextRun({ text: report.title, bold: true, size: 24, font: "Times New Roman" })],
          alignment: AlignmentType.LEFT,
          spacing: { before: 200, after: 0 }
        }));

        const currentDateStr = formatDateDDMMYYYY(report.currentDate || report.currentPeriod);
        const previousDateStr = formatDateDDMMYYYY(report.previousDate || report.previousPeriod);

        const isOnlyAV = !!report.onlyAV;
        const isHideAV = !!report.hideAV;
        const isHideAH = !!report.hideAH;

        const colWidths = isOnlyAV
          ? [60, 20, 20]
          : (isHideAV && isHideAH)
            ? [50, 25, 25]
            : isHideAV
              ? [52, 16, 16, 16]
              : [44, 14, 14, 14, 14];

        const headerBorder = {
          top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
          bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
        };

        const headerShading = { fill: "D9E1F2" }; // Azul, Ênfase 1, Mais claro 80%

        const buildRowCells = (
          rubrica: string,
          atual: number,
          anterior: number,
          variacao: number | undefined,
          av: number | undefined,
          isBold = false,
          isItalic = false,
          bgColor?: string,
          textColor?: string
        ) => {
          const cells: TableCell[] = [];
          const textProps = {
            font: "Times New Roman",
            size: 20,
            bold: isBold,
            italics: isItalic,
            ...(textColor ? { color: textColor } : {})
          };
          const cellShading = bgColor ? { fill: bgColor } : undefined;

          // Col 0: Rubrica
          cells.push(new TableCell({
            width: { size: colWidths[0], type: WidthType.PERCENTAGE },
            shading: cellShading,
            children: [new Paragraph({ children: [new TextRun({ ...textProps, text: rubrica || "" })], alignment: AlignmentType.LEFT })]
          }));

          // Col 1: Atual
          cells.push(new TableCell({
            width: { size: colWidths[1], type: WidthType.PERCENTAGE },
            shading: cellShading,
            children: [new Paragraph({ children: [new TextRun({ ...textProps, text: formatBRL(atual || 0) })], alignment: AlignmentType.RIGHT })]
          }));

          // Col 2: Anterior (if not onlyAV)
          if (!isOnlyAV) {
            cells.push(new TableCell({
              width: { size: colWidths[2], type: WidthType.PERCENTAGE },
              shading: cellShading,
              children: [new Paragraph({ children: [new TextRun({ ...textProps, text: formatBRL(anterior || 0) })], alignment: AlignmentType.RIGHT })]
            }));
          }

          // Col 3: AH (%) (if not onlyAV and not hideAH)
          if (!isOnlyAV && !isHideAH) {
            cells.push(new TableCell({
              width: { size: colWidths[isHideAV ? 3 : 3], type: WidthType.PERCENTAGE },
              shading: cellShading,
              children: [new Paragraph({ children: [new TextRun({ ...textProps, text: variacao !== undefined && !isNaN(variacao) ? formatPercentWord(variacao) : "" })], alignment: AlignmentType.RIGHT })]
            }));
          }

          // Col 4: AV (%) (if not hideAV OR if onlyAV)
          if (!isHideAV || isOnlyAV) {
            const avColIdx = isOnlyAV ? 2 : 4;
            cells.push(new TableCell({
              width: { size: colWidths[avColIdx], type: WidthType.PERCENTAGE },
              shading: cellShading,
              children: [new Paragraph({ children: [new TextRun({ ...textProps, text: av !== undefined && !isNaN(av) ? formatPercentWord(av, isBold && Math.abs((av || 0) - 100) < 0.01) : "" })], alignment: AlignmentType.RIGHT })]
            }));
          }

          return cells;
        };

        // Header cells
        const headerCells: TableCell[] = [
          new TableCell({ 
            width: { size: colWidths[0], type: WidthType.PERCENTAGE },
            shading: headerShading,
            borders: headerBorder,
            children: [new Paragraph({ children: [new TextRun({ text: "Rubrica", bold: true, font: "Times New Roman", size: 20, color: "000000" })], alignment: AlignmentType.LEFT })] 
          }),
          new TableCell({ 
            width: { size: colWidths[1], type: WidthType.PERCENTAGE },
            shading: headerShading,
            borders: headerBorder,
            children: [new Paragraph({ children: [new TextRun({ text: currentDateStr, bold: true, font: "Times New Roman", size: 20, color: "000000" })], alignment: AlignmentType.RIGHT })] 
          }),
          ...(!isOnlyAV ? [
            new TableCell({ 
              width: { size: colWidths[2], type: WidthType.PERCENTAGE },
              shading: headerShading,
              borders: headerBorder,
              children: [new Paragraph({ children: [new TextRun({ text: previousDateStr, bold: true, font: "Times New Roman", size: 20, color: "000000" })], alignment: AlignmentType.RIGHT })] 
            })
          ] : []),
          ...(!isOnlyAV && !isHideAH ? [
            new TableCell({ 
              width: { size: colWidths[isHideAV ? 3 : 3], type: WidthType.PERCENTAGE },
              shading: headerShading,
              borders: headerBorder,
              children: [new Paragraph({ children: [new TextRun({ text: "AH (%)", bold: true, font: "Times New Roman", size: 20, color: "000000" })], alignment: AlignmentType.RIGHT })] 
            })
          ] : []),
          ...((!isHideAV || isOnlyAV) ? [
            new TableCell({ 
              width: { size: colWidths[isOnlyAV ? 2 : 4], type: WidthType.PERCENTAGE },
              shading: headerShading,
              borders: headerBorder,
              children: [new Paragraph({ children: [new TextRun({ text: "AV (%)", bold: true, font: "Times New Roman", size: 20, color: "000000" })], alignment: AlignmentType.RIGHT })] 
            })
          ] : [])
        ];

        // Create Table
        const docxTable = new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE },
            insideHorizontal: { style: BorderStyle.NONE },
            insideVertical: { style: BorderStyle.NONE },
          },
          rows: [
            // First Row (Header)
            new DocxTableRow({
              children: headerCells
            }),
            // Data Rows
            ...report.rows.map(row => new DocxTableRow({
              children: buildRowCells(
                row.rubrica,
                row.atual,
                row.anterior,
                row.variacao,
                row.av,
                !!row.isSubtotal,
                !!row.isSubtotal
              )
            })),
            // Subtotals
            ...(report.subtotals || []).map(sub => new DocxTableRow({
              children: buildRowCells(
                sub.rubrica,
                sub.atual,
                sub.anterior,
                sub.variacao,
                sub.av,
                true,
                true
              )
            })),
            // Total Row (Fundo azul escuro, texto branco em negrito)
            new DocxTableRow({
              children: buildRowCells(
                report.total.rubrica || "Total Geral",
                report.total.atual,
                report.total.anterior,
                report.total.variacao,
                report.total.av,
                true,
                false,
                "1F4E79",
                "FFFFFF"
              )
            }),
          ],
        });

        allElements.push(docxTable);

        // Source Text: Times New Roman, 10, Alignment Left
        allElements.push(new Paragraph({
          children: [new TextRun({ text: "Fonte: SIAFI.", font: "Times New Roman", size: 20, italics: true })],
          alignment: AlignmentType.LEFT,
          spacing: { before: 100, after: 200 }
        }));

        // Commentary: Justified, Times New Roman, 12
        if (report.commentary) {
          allElements.push(new Paragraph({
            children: [new TextRun({ text: report.commentary, font: "Times New Roman", size: 24 })],
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 400 }
          }));
        }
      }


return allElements;
}
export async function generateNotesWord(reports: TableData[]) { return Packer.toBlob(new Document({sections:[{children:buildNotesChildren(reports)}]})); }
