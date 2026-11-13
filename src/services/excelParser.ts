import { extractReportPeriod } from '../utils/reportPeriod';
import * as XLSX from 'xlsx';
import { BalancoData, DocumentMetadata, StatementItem } from '../types/balanco';
import { parseExcelNumber } from '../utils/numberFormat';
import {
  INITIAL_NE_MAPPINGS,
  getBlankAtivoStructure,
  getBlankPassivoStructure,
  getBlankFPStructure,
  getBlankCompensacoesStructure,
  getBlankSuperavitStructure,
} from './statementStructure';

export interface ParseResult {
  data: BalancoData;
  warnings: string[];
  errors: string[];
}

/**
 * Normalizes string for text comparison while retaining character identity
 */
export function normalizeLabel(str: string): string {
  return str
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[()\-–—.,_]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

type CellParseStatus = 'number' | 'dash' | 'empty' | 'error' | 'unmapped';

interface CellReadResult {
  val: number | null;
  ref: string;
  status: CellParseStatus;
  rawText: string;
}

/**
 * Universal Excel Parser for COMAER SIAFI Financial Statements
 */
export async function parseBalancoExcel(
  fileBuffer: ArrayBuffer,
  forcedSheetName?: string
): Promise<ParseResult> {
  const warnings: string[] = [];
  const errors: string[] = [];

  const workbook = XLSX.read(fileBuffer, {
    type: 'array',
    cellFormula: true,
    cellStyles: true,
    cellDates: true,
  });

  const sheetNames = workbook.SheetNames;
  if (!sheetNames || sheetNames.length === 0) {
    throw new Error('O arquivo Excel não contém nenhuma planilha.');
  }

  let targetSheetName = forcedSheetName || '';
  if (!targetSheetName) {
    for (const name of sheetNames) {
      const lower = name.toLowerCase();
      if (lower.includes('balanco') || lower.includes('balanço') || lower.includes('896') || lower.includes('bp')) {
        targetSheetName = name;
        break;
      }
    }

    if (!targetSheetName) {
      for (const name of sheetNames) {
        const ws = workbook.Sheets[name];
        const text = JSON.stringify(ws);
        if (text.includes('BALANÇO PATRIMONIAL') || text.includes('Ativo Circulante')) {
          targetSheetName = name;
          break;
        }
      }
    }

    if (!targetSheetName) {
      targetSheetName = sheetNames[0];
      warnings.push(`Nenhuma aba com nome explícito identificada. Utilizando aba: "${targetSheetName}".`);
    }
  }

  const worksheet = workbook.Sheets[targetSheetName];
  if (!worksheet) {
    throw new Error(`Planilha "${targetSheetName}" não encontrada no arquivo.`);
  }

  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z250');
  const maxR = Math.max(range.e.r, 200);
  const maxC = Math.max(range.e.c, 25);

  const grid: { val: unknown; raw: string; address: string }[][] = [];
  for (let r = 0; r <= maxR; r++) {
    const row: { val: unknown; raw: string; address: string }[] = [];
    for (let c = 0; c <= maxC; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c });
      const cell = worksheet[cellAddress];
      let val: unknown = null;
      let raw = '';
      if (cell) {
        val = cell.v !== undefined ? cell.v : null;
        raw = cell.w || (cell.v !== undefined ? String(cell.v) : '');
      }
      row.push({ val, raw, address: cellAddress });
    }
    grid.push(row);
  }

  // 1. EXTRAÇÃO DE METADADOS DO CABEÇALHO
  const metadata: DocumentMetadata = {
    title: 'Balanço Patrimonial',
    orgaoCode: '52111',
    orgaoName: 'Comando da Aeronáutica',
    exercicio: 2026,
    periodo: '',
    dataEmissao: '',
    unidadeOriginal: 'milhares de reais',
    unidadeNormalizada: 'R$ milhares',
    unitScaleFactor: 1, // O arquivo original já está em milhares de reais. NÃO dividir por 1.000!
    colCurrent: '2026',
    colPrevious: '2025',
    sourceSheetName: targetSheetName,
    availableSheets: sheetNames,
    unitConfirmed: true,
  };

  for (let r = 0; r <= Math.min(15, maxR); r++) {
    const rowText = grid[r].map(c => String(c.raw || c.val || '')).join(' ').trim();
    if (!rowText) continue;

    const orgaoMatch = rowText.match(/(?:Órgão|Orgao|Unidade Gestora|UG):\s*([0-9]{5,6})\s*[-—–]\s*([^\n;]+)/i);
    if (orgaoMatch) {
      metadata.orgaoCode = orgaoMatch[1].trim();
      metadata.orgaoName = orgaoMatch[2].trim();
    }

    const exMatch = rowText.match(/(?:Exercício|Exercicio|Ano)\s*:?\s*(20\d\d)/i);
    if (exMatch) {
      metadata.exercicio = parseInt(exMatch[1], 10);
      metadata.colCurrent = String(metadata.exercicio);
    }

    const dtMatch = rowText.match(/(?:Data\s+de\s+Emissão|Emissão|Data):\s*(\d{2}\/\d{2}\/\d{4})/i);
    if (dtMatch) {
      metadata.dataEmissao = dtMatch[1].trim();
    }
  }

  metadata.periodo = extractReportPeriod(grid.map(row => row.map(cell => cell.raw || cell.val)));
  if (!metadata.periodo) warnings.push('Período não identificado no cabeçalho do BP. Nenhum mês padrão foi atribuído.');

  // Identificação de colunas comparativas (anos 2026 e 2025)
  for (let r = 10; r <= Math.min(16, maxR); r++) {
    for (let c = 0; c <= maxC; c++) {
      const text = String(grid[r][c].val || '').trim();
      const yrMatch = text.match(/^(20\d\d)$/);
      if (yrMatch) {
        const yr = parseInt(yrMatch[1], 10);
        if (!metadata.colCurrent || yr > parseInt(metadata.colCurrent, 10)) {
          metadata.colPrevious = metadata.colCurrent || String(yr - 1);
          metadata.colCurrent = String(yr);
        } else if (yr < parseInt(metadata.colCurrent, 10)) {
          metadata.colPrevious = String(yr);
        }
      }
    }
  }

  // Função para ler célula com rigor absoluto
  const readCell = (r: number, c: number): CellReadResult => {
    if (r < 0 || r >= grid.length || c < 0 || c >= grid[r].length) {
      return { val: null, ref: '-', status: 'unmapped', rawText: '' };
    }
    const cell = grid[r][c];
    const raw = String(cell.raw || cell.val || '').trim();

    if (raw === '' || cell.val === null || cell.val === undefined) {
      return { val: null, ref: cell.address, status: 'empty', rawText: raw };
    }

    if (raw === '-' || raw === '–' || raw === '—') {
      return { val: null, ref: cell.address, status: 'dash', rawText: raw };
    }

    if (raw.startsWith('#')) {
      return { val: null, ref: cell.address, status: 'error', rawText: raw };
    }

    const parsed = parseExcelNumber(cell.val !== null ? cell.val : raw);
    if (parsed === null) {
      return { val: null, ref: cell.address, status: 'empty', rawText: raw };
    }

    return { val: parsed, ref: cell.address, status: 'number', rawText: raw };
  };

  // Verificador se um texto representa conta redutora
  const isRedutoraText = (text: string): boolean => {
    if (!text) return false;
    const norm = normalizeLabel(text);
    return (
      text.includes('(-)') ||
      norm.includes('perda') ||
      norm.includes('deprec') ||
      norm.includes('amortiz') ||
      norm.includes('exaust')
    );
  };

  // Leitor resiliente de descrição da linha (considera Col B=1, Col A=0 e Col C=2 para Ativo, e Col L=11 para Passivo)
  const getRowDescription = (r: number, preferredCol: number): string => {
    if (r < 0 || r >= grid.length) return '';
    const primary = String(grid[r][preferredCol]?.val || grid[r][preferredCol]?.raw || '').trim();
    if (primary) return primary;
    if (preferredCol === 1) {
      const colA = String(grid[r][0]?.val || grid[r][0]?.raw || '').trim();
      if (colA) return colA;
      const colC = String(grid[r][2]?.val || grid[r][2]?.raw || '').trim();
      if (colC) return colC;
    }
    if (preferredCol === 11) {
      const colK = String(grid[r][10]?.val || grid[r][10]?.raw || '').trim();
      if (colK) return colK;
      const colM = String(grid[r][12]?.val || grid[r][12]?.raw || '').trim();
      if (colM) return colM;
    }
    return '';
  };

  // Localizador por chave composta e restrição de seção hierárquica
  const findRowByCompositeKey = (
    descCol: number,
    label: string,
    expectedRow1Indexed: number | undefined,
    minRow1Indexed: number,
    maxRow1Indexed: number,
    forceRedutora?: boolean
  ): number => {
    const normTarget = normalizeLabel(label);
    const targetKeywords = normTarget.split(' ').filter(w => w.length > 2);
    const targetIsRedutora = forceRedutora !== undefined ? forceRedutora : isRedutoraText(label);

    const startR = Math.max(0, minRow1Indexed - 1);
    const endR = Math.min(grid.length - 1, maxRow1Indexed - 1);

    // 1. Testa a linha esperada e vizinhança imediata por ordem de proximidade
    if (expectedRow1Indexed !== undefined) {
      const offsets = [0, 1, -1, 2, -2, 3, -3, 4, -4, 5, -5];
      for (const offset of offsets) {
        const targetR = expectedRow1Indexed - 1 + offset;
        if (targetR >= startR && targetR <= endR) {
          const cellDesc = getRowDescription(targetR, descCol);
          const normCell = normalizeLabel(cellDesc);
          if (normCell) {
            const cellIsRedutora = isRedutoraText(cellDesc);
            // Redutoras só podem casar com redutoras; rubricas normais nunca com redutoras
            if (targetIsRedutora !== cellIsRedutora) {
              continue;
            }

            if (
              normCell === normTarget ||
              normCell.includes(normTarget) ||
              (targetKeywords.length > 0 && targetKeywords.every(kw => normCell.includes(kw))) ||
              (normCell.slice(0, 15) === normTarget.slice(0, 15) && normCell.length > 5)
            ) {
              return targetR;
            }
          }
        }
      }
    }

    // 2a. Match exato normalizado
    for (let r = startR; r <= endR; r++) {
      const cellDesc = getRowDescription(r, descCol);
      if (!cellDesc) continue;
      const normCell = normalizeLabel(cellDesc);
      const cellIsRedutora = isRedutoraText(cellDesc);
      if (targetIsRedutora !== cellIsRedutora) continue;

      if (normCell === normTarget) {
        return r;
      }
    }

    // 2b. Match por inclusão / palavras-chave principais dentro do bloco
    for (let r = startR; r <= endR; r++) {
      const cellDesc = getRowDescription(r, descCol);
      if (!cellDesc) continue;
      const normCell = normalizeLabel(cellDesc);
      const cellIsRedutora = isRedutoraText(cellDesc);
      if (targetIsRedutora !== cellIsRedutora) continue;

      if (
        (normCell.length > 5 && normCell.includes(normTarget)) ||
        (targetKeywords.length > 1 && targetKeywords.every(kw => normCell.includes(kw)))
      ) {
        return r;
      }
    }

    // 2c. Match por prefixo significativo (ex: "Demais Obrigações", "Provisões", etc.)
    const firstTwoWords = targetKeywords.slice(0, 2).join(' ');
    if (firstTwoWords.length > 6) {
      for (let r = startR; r <= endR; r++) {
        const cellDesc = getRowDescription(r, descCol);
        if (!cellDesc) continue;
        const normCell = normalizeLabel(cellDesc);
        const cellIsRedutora = isRedutoraText(cellDesc);
        if (targetIsRedutora !== cellIsRedutora) continue;

        if (normCell.startsWith(firstTwoWords) || normCell.includes(firstTwoWords)) {
          return r;
        }
      }
    }

    // 3. Fallback determinístico para a linha esperada se definida dentro dos limites da grade
    if (expectedRow1Indexed !== undefined && expectedRow1Indexed - 1 < grid.length) {
      return expectedRow1Indexed - 1;
    }

    return -1;
  };

  // 2. EXTRAÇÃO DAS RUBRICAS DO ATIVO (Col B desc, Col I curr, Col K prev)
  const blankAtivo = getBlankAtivoStructure();
  const ativoItems: StatementItem[] = blankAtivo.map(item => {
    // -------------------------------------------------------------
    // CRÉDITOS A CURTO PRAZO E SUBORDINADOS (Linhas 18 a 26)
    // -------------------------------------------------------------

    // Linha 18: Créditos a Curto Prazo (Subtotal Líquido informado)
    if (item.id === 'at_cred_cp') {
      const r18 = 17; // 0-indexed para linha 18
      const cCurr = readCell(r18, 8); // Col I = 8
      const cPrev = readCell(r18, 10); // Col K = 10
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I18',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K18',
        rowSource: 18,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 21: Empréstimos e Financiamentos Concedidos (Líquido)
    if (item.id === 'at_cred_cp_emp_liq') {
      const r21 = 20; // 0-indexed para linha 21
      const cCurr = readCell(r21, 8);
      const cPrev = readCell(r21, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I21',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K21',
        rowSource: 21,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 22: Empréstimos e Financiamentos Concedidos (Bruto)
    if (item.id === 'at_cred_cp_emp_bruto') {
      const r22 = 21; // 0-indexed para linha 22
      const cCurr = readCell(r22, 8);
      const cPrev = readCell(r22, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I22',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K22',
        rowSource: 22,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 23: (-) Perdas Estimadas em Empréstimos e Financ. Concedidos (Redutora)
    if (item.id === 'at_cred_cp_emp_perdas') {
      const r23 = 22; // 0-indexed para linha 23
      const cCurr = readCell(r23, 8);
      const cPrev = readCell(r23, 10);
      const finalCurr = cCurr.val !== null ? -Math.abs(cCurr.val) : null;
      const finalPrev = cPrev.val !== null ? -Math.abs(cPrev.val) : null;
      return {
        ...item,
        valCurrent: finalCurr,
        valPrevious: finalPrev,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I23',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K23',
        rowSource: 23,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 24: Demais Créditos e Valores a Curto Prazo (Líquido)
    if (item.id === 'at_cred_cp_demais_liq') {
      const r24 = 23; // 0-indexed para linha 24
      const cCurr = readCell(r24, 8);
      const cPrev = readCell(r24, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I24',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K24',
        rowSource: 24,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 25: Demais Créditos e Valores a Curto Prazo (Bruto)
    if (item.id === 'at_cred_cp_demais_bruto') {
      const r25 = 24; // 0-indexed para linha 25
      const cCurr = readCell(r25, 8);
      const cPrev = readCell(r25, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I25',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K25',
        rowSource: 25,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 26: (-) Perdas Estimadas em Demais Créditos a Curto Prazo (Redutora)
    if (item.id === 'at_cred_cp_demais_perdas') {
      const r26 = 25; // 0-indexed para linha 26
      let targetR = r26;
      let cCurr = readCell(targetR, 8);
      let cPrev = readCell(targetR, 10);

      // Se a linha 26 não contiver o saldo, busca nas linhas 24..27 por "perdas"
      const desc26 = getRowDescription(targetR, 1);
      if (cCurr.val === null || !isRedutoraText(desc26)) {
        for (let r = 24; r <= 27; r++) {
          const d = getRowDescription(r, 1);
          if (isRedutoraText(d)) {
            const testCurr = readCell(r, 8);
            if (testCurr.val !== null) {
              targetR = r;
              cCurr = testCurr;
              cPrev = readCell(r, 10);
              break;
            }
          }
        }
      }

      // Preservar sempre o sinal negativo da conta redutora
      const finalCurr = cCurr.val !== null ? -Math.abs(cCurr.val) : null;
      const finalPrev = cPrev.val !== null ? -Math.abs(cPrev.val) : null;

      return {
        ...item,
        valCurrent: finalCurr,
        valPrevious: finalPrev,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : `I${targetR + 1}`,
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : `K${targetR + 1}`,
        rowSource: targetR + 1,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // -------------------------------------------------------------
    // INVESTIMENTOS E APLICAÇÕES TEMPORÁRIAS A CURTO PRAZO (Linha 27)
    // -------------------------------------------------------------
    if (item.id === 'at_inv_cp') {
      const r27 = 26; // 0-indexed para linha 27
      let matchedR = findRowByCompositeKey(1, 'Investimentos e Aplicações Temporárias a Curto Prazo', 27, 16, 30, false);
      if (matchedR === -1) {
        // Busca flexível por "aplicações temporárias" ou "investimentos" no bloco circulante (linhas 16 a 30)
        for (let r = 15; r <= 29; r++) {
          const desc = getRowDescription(r, 1);
          const norm = normalizeLabel(desc);
          if (
            (norm.includes('invest') && (norm.includes('temp') || norm.includes('aplic'))) ||
            norm.includes('aplicacoes temporarias')
          ) {
            matchedR = r;
            break;
          }
        }
      }

      if (matchedR !== -1) {
        const cCurr = readCell(matchedR, 8);
        const cPrev = readCell(matchedR, 10);
        const rawCurr = cCurr.rawText ? String(cCurr.rawText).trim() : '';
        const rawPrev = cPrev.rawText ? String(cPrev.rawText).trim() : '';
        const isCurrUnreadable = rawCurr !== '' && rawCurr !== '-' && cCurr.val === null;
        const isPrevUnreadable = rawPrev !== '' && rawPrev !== '-' && cPrev.val === null;

        return {
          ...item,
          valCurrent: cCurr.val,
          valPrevious: cPrev.val,
          cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : `I${matchedR + 1}`,
          cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : `K${matchedR + 1}`,
          rowSource: matchedR + 1,
          isUnmapped: isCurrUnreadable || isPrevUnreadable,
          neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
        };
      }

      // Se a linha 27 tiver descrição de investimento ou estiver no layout padrão
      const desc27 = getRowDescription(r27, 1);
      const c27Curr = readCell(r27, 8);
      const c27Prev = readCell(r27, 10);
      if (desc27 && normalizeLabel(desc27).includes('invest')) {
        return {
          ...item,
          valCurrent: c27Curr.val,
          valPrevious: c27Prev.val,
          cellRefCurrent: c27Curr.ref !== '-' ? c27Curr.ref : 'I27',
          cellRefPrevious: c27Prev.ref !== '-' ? c27Prev.ref : 'K27',
          rowSource: 27,
          isUnmapped: false,
          neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
        };
      }

      // Caso a rubrica não esteja presente no arquivo, mantém saldo nulo sem erro
      return {
        ...item,
        valCurrent: null,
        valPrevious: null,
        cellRefCurrent: 'I27',
        cellRefPrevious: 'K27',
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // -------------------------------------------------------------
    // ESTOQUES (Linha 28)
    // -------------------------------------------------------------
    if (item.id === 'at_estoques') {
      const matchedR = findRowByCompositeKey(1, 'Estoques', 28, 16, 30, false);
      const targetR = matchedR !== -1 ? matchedR : 27; // linha 28 = index 27
      const cCurr = readCell(targetR, 8);
      const cPrev = readCell(targetR, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : `I${targetR + 1}`,
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : `K${targetR + 1}`,
        rowSource: targetR + 1,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // -------------------------------------------------------------
    // CRÉDITOS A LONGO PRAZO E SUBORDINADOS (Linhas 33 a 39)
    // -------------------------------------------------------------
    if (item.id === 'at_rlp_cred_lp') {
      const r33 = 32;
      const cCurr = readCell(r33, 8);
      const cPrev = readCell(r33, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I33',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K33',
        rowSource: 33,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    if (item.id === 'at_rlp_emp_liq') {
      const r34 = 33;
      const cCurr = readCell(r34, 8);
      const cPrev = readCell(r34, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I34',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K34',
        rowSource: 34,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    if (item.id === 'at_rlp_emp_bruto') {
      const r35 = 34;
      const cCurr = readCell(r35, 8);
      const cPrev = readCell(r35, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I35',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K35',
        rowSource: 35,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    if (item.id === 'at_rlp_emp_perdas') {
      const r36 = 35;
      const cCurr = readCell(r36, 8);
      const cPrev = readCell(r36, 10);
      const finalCurr = cCurr.val !== null ? -Math.abs(cCurr.val) : null;
      const finalPrev = cPrev.val !== null ? -Math.abs(cPrev.val) : null;
      return {
        ...item,
        valCurrent: finalCurr,
        valPrevious: finalPrev,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I36',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K36',
        rowSource: 36,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    if (item.id === 'at_rlp_demais_liq') {
      const r37 = 36;
      const cCurr = readCell(r37, 8);
      const cPrev = readCell(r37, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I37',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K37',
        rowSource: 37,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    if (item.id === 'at_rlp_demais_bruto') {
      const r38 = 37;
      const cCurr = readCell(r38, 8);
      const cPrev = readCell(r38, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I38',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K38',
        rowSource: 38,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    if (item.id === 'at_rlp_demais_perdas') {
      const r39 = 38;
      const cCurr = readCell(r39, 8);
      const cPrev = readCell(r39, 10);
      const finalCurr = cCurr.val !== null ? -Math.abs(cCurr.val) : null;
      const finalPrev = cPrev.val !== null ? -Math.abs(cPrev.val) : null;
      return {
        ...item,
        valCurrent: finalCurr,
        valPrevious: finalPrev,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I39',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K39',
        rowSource: 39,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // -------------------------------------------------------------
    // IMOBILIZADO: BENS MÓVEIS E BENS IMÓVEIS (Linhas 66 a 79)
    // -------------------------------------------------------------
    // Linha 66: Bens Móveis (Líquido)
    if (item.id === 'at_imob_moveis_liq') {
      const r66 = 65;
      const cCurr = readCell(r66, 8);
      const cPrev = readCell(r66, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I66',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K66',
        rowSource: 66,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 68: Bens Móveis (Bruto)
    if (item.id === 'at_imob_moveis_bruto') {
      const r68 = 67;
      const cCurr = readCell(r68, 8);
      const cPrev = readCell(r68, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I68',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K68',
        rowSource: 68,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Tratamento especializado para Depreciação de Bens Móveis (linhas 70 e 71)
    if (item.id === 'at_imob_moveis_deprec') {
      const r70 = 69; // 0-indexed para linha 70
      const r71 = 70; // 0-indexed para linha 71
      const c70_curr = readCell(r70, 8); // Col I = 8
      const c71_curr = readCell(r71, 8);
      const c70_prev = readCell(r70, 10); // Col K = 10
      const c71_prev = readCell(r71, 10);

      // Exercício corrente (2026 / Col I): verifica linhas 70 e 71
      let valCurr: number | null = null;
      let refCurr = 'I70';
      let rowSrc = 70;

      if (c70_curr.val !== null && c71_curr.val !== null) {
        valCurr = c70_curr.val;
        refCurr = 'I70';
        rowSrc = 70;
      } else if (c70_curr.val !== null) {
        valCurr = c70_curr.val;
        refCurr = 'I70';
        rowSrc = 70;
      } else if (c71_curr.val !== null) {
        valCurr = c71_curr.val;
        refCurr = 'I71';
        rowSrc = 71;
      } else {
        const fallbackR = findRowByCompositeKey(1, 'depreciação', 70, 68, 73, true);
        if (fallbackR !== -1) {
          const fc = readCell(fallbackR, 8);
          valCurr = fc.val;
          refCurr = fc.ref;
          rowSrc = fallbackR + 1;
        }
      }

      // Exercício anterior (2025 / Col K): verifica linhas 70 e 71
      let valPrev: number | null = null;
      let refPrev = 'K70';

      if (c70_prev.val !== null && c71_prev.val !== null) {
        valPrev = c70_prev.val;
        refPrev = 'K70';
      } else if (c70_prev.val !== null) {
        valPrev = c70_prev.val;
        refPrev = 'K70';
      } else if (c71_prev.val !== null) {
        valPrev = c71_prev.val;
        refPrev = 'K71';
      } else if (rowSrc > 0) {
        const fp = readCell(rowSrc - 1, 10);
        valPrev = fp.val;
        refPrev = fp.ref;
      }

      // Preservar sempre o sinal negativo da conta redutora
      const finalCurr = valCurr !== null ? -Math.abs(valCurr) : null;
      const finalPrev = valPrev !== null ? -Math.abs(valPrev) : null;

      return {
        ...item,
        valCurrent: finalCurr,
        valPrevious: finalPrev,
        cellRefCurrent: refCurr,
        cellRefPrevious: refPrev,
        rowSource: rowSrc,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 74: Bens Imóveis (Líquido)
    if (item.id === 'at_imob_imoveis_liq') {
      const r74 = 73;
      const cCurr = readCell(r74, 8);
      const cPrev = readCell(r74, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I74',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K74',
        rowSource: 74,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Linha 76: Bens Imóveis (Bruto)
    if (item.id === 'at_imob_imoveis_bruto') {
      const r76 = 75;
      const cCurr = readCell(r76, 8);
      const cPrev = readCell(r76, 10);
      return {
        ...item,
        valCurrent: cCurr.val,
        valPrevious: cPrev.val,
        cellRefCurrent: cCurr.ref !== '-' ? cCurr.ref : 'I76',
        cellRefPrevious: cPrev.ref !== '-' ? cPrev.ref : 'K76',
        rowSource: 76,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    // Tratamento especializado para Depreciação de Bens Imóveis (linhas 78 e 79)
    if (item.id === 'at_imob_imoveis_deprec') {
      const r78 = 77; // 0-indexed para linha 78
      const r79 = 78; // 0-indexed para linha 79
      const c78_curr = readCell(r78, 8);
      const c79_curr = readCell(r79, 8);
      const c78_prev = readCell(r78, 10);
      const c79_prev = readCell(r79, 10);

      let valCurr: number | null = null;
      let refCurr = 'I78';
      let rowSrc = 78;

      if (c78_curr.val !== null && c79_curr.val !== null) {
        valCurr = c78_curr.val;
        refCurr = 'I78';
        rowSrc = 78;
      } else if (c78_curr.val !== null) {
        valCurr = c78_curr.val;
        refCurr = 'I78';
        rowSrc = 78;
      } else if (c79_curr.val !== null) {
        valCurr = c79_curr.val;
        refCurr = 'I79';
        rowSrc = 79;
      } else {
        const fallbackR = findRowByCompositeKey(1, 'depreciação', 78, 76, 81, true);
        if (fallbackR !== -1) {
          const fc = readCell(fallbackR, 8);
          valCurr = fc.val;
          refCurr = fc.ref;
          rowSrc = fallbackR + 1;
        }
      }

      let valPrev: number | null = null;
      let refPrev = 'K78';

      if (c78_prev.val !== null && c79_prev.val !== null) {
        valPrev = c78_prev.val;
        refPrev = 'K78';
      } else if (c78_prev.val !== null) {
        valPrev = c78_prev.val;
        refPrev = 'K78';
      } else if (c79_prev.val !== null) {
        valPrev = c79_prev.val;
        refPrev = 'K79';
      } else if (rowSrc > 0) {
        const fp = readCell(rowSrc - 1, 10);
        valPrev = fp.val;
        refPrev = fp.ref;
      }

      // Preservar sempre o sinal negativo da conta redutora
      const finalCurr = valCurr !== null ? -Math.abs(valCurr) : null;
      const finalPrev = valPrev !== null ? -Math.abs(valPrev) : null;

      return {
        ...item,
        valCurrent: finalCurr,
        valPrevious: finalPrev,
        cellRefCurrent: refCurr,
        cellRefPrevious: refPrev,
        rowSource: rowSrc,
        isUnmapped: false,
        neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
      };
    }

    let minR = 15;
    let maxR = 125;
    if (
      item.id.startsWith('at_circ') ||
      item.id.startsWith('at_cx') ||
      item.id.startsWith('at_cred_cp') ||
      item.id.startsWith('at_inv_cp') ||
      item.id === 'at_estoques'
    ) {
      minR = 15;
      maxR = 30;
    } else if (item.id.startsWith('at_rlp') || item.id === 'at_invest') {
      minR = 31;
      maxR = 43;
    } else if (item.id.startsWith('at_imob') || item.id === 'at_intangivel') {
      minR = 63;
      maxR = 85;
    } else if (item.id === 'at_total') {
      minR = 117;
      maxR = 122;
    }

    const matchedR = findRowByCompositeKey(1, item.label, item.expectedRow, minR, maxR, item.isRedutora);
    if (matchedR === -1) {
      return {
        ...item,
        isUnmapped: true,
        valCurrent: null,
        valPrevious: null,
        cellRefCurrent: 'Pendente',
        cellRefPrevious: 'Pendente',
      };
    }

    const curr = readCell(matchedR, 8); // Col I = 8
    const prev = readCell(matchedR, 10); // Col K = 10

    let valCurr = curr.val;
    let valPrev = prev.val;

    // Preservar sinal negativo de contas redutoras (ex: perdas estimadas)
    if (item.isRedutora || item.label.startsWith('(-)')) {
      if (valCurr !== null) valCurr = -Math.abs(valCurr);
      if (valPrev !== null) valPrev = -Math.abs(valPrev);
    }

    return {
      ...item,
      valCurrent: valCurr,
      valPrevious: valPrev,
      cellRefCurrent: curr.ref,
      cellRefPrevious: prev.ref,
      rowSource: matchedR + 1,
      isUnmapped: false,
      neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
    };
  });

  // 3. EXTRAÇÃO DO PASSIVO E PATRIMÔNIO LÍQUIDO (Col L desc, Col O curr, Col P prev)
  const blankPassivo = getBlankPassivoStructure();
  const passivoItems: StatementItem[] = blankPassivo.map(item => {
    let minR = 16;
    let maxR = 125;
    if (
      item.compositeKey?.includes('Passivo Circulante') ||
      item.id.startsWith('pass_circ') ||
      item.id.endsWith('_cp') ||
      item.id === 'pass_trab' ||
      item.id === 'pass_fornec' ||
      item.id === 'pass_transf_fisc'
    ) {
      minR = 16;
      maxR = 25;
    } else if (
      item.compositeKey?.includes('Passivo Não Circulante') ||
      item.id.startsWith('pass_nao_circ') ||
      item.id.endsWith('_lp') ||
      item.id === 'pass_res_dif'
    ) {
      minR = 31;
      maxR = 40;
    } else if (item.id === 'pass_exig_total') {
      minR = 39;
      maxR = 41;
    } else if (item.compositeKey?.includes('Patrimônio Líquido') || item.id.startsWith('pl_')) {
      minR = 41;
      maxR = 75;
    } else if (item.id === 'pass_total') {
      minR = 115;
      maxR = 125;
    }

    const matchedR = findRowByCompositeKey(11, item.label, item.expectedRow, minR, maxR);
    if (matchedR === -1) {
      return {
        ...item,
        isUnmapped: true,
        valCurrent: null,
        valPrevious: null,
        cellRefCurrent: 'Pendente',
        cellRefPrevious: 'Pendente',
      };
    }

    const curr = readCell(matchedR, 14); // Col O = 14
    const prev = readCell(matchedR, 15); // Col P = 15

    return {
      ...item,
      valCurrent: curr.val,
      valPrevious: prev.val,
      cellRefCurrent: curr.ref,
      cellRefPrevious: prev.ref,
      rowSource: matchedR + 1,
      isUnmapped: false,
      neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
    };
  });

  // 4. EXTRAÇÃO DO QUADRO DE ATIVOS E PASSIVOS FINANCEIROS E PERMANENTE
  // Ativo Financeiro (I128/K128), Ativo Permanente (I129/K129)
  // Passivo Financeiro (O128/P128), Passivo Permanente (O129/P129)
  // Saldo Patrimonial (O130/P130)
  const blankFP = getBlankFPStructure();
  const atFinCurr = readCell(127, 8); // I128 (0-indexed = 127)
  const atFinPrev = readCell(127, 10); // K128
  const atPermCurr = readCell(128, 8); // I129
  const atPermPrev = readCell(128, 10); // K129

  const passFinCurr = readCell(127, 14); // O128
  const passFinPrev = readCell(127, 15); // P128
  const passPermCurr = readCell(128, 14); // O129
  const passPermPrev = readCell(128, 15); // P129

  const saldoPatrCurr = readCell(129, 14); // O130
  const saldoPatrPrev = readCell(129, 15); // P130

  // Cálculo agregado de Ativo (I) e Passivo (II)
  const calcAtivoICurr = (atFinCurr.val || 0) + (atPermCurr.val || 0);
  const calcAtivoIPrev = (atFinPrev.val || 0) + (atPermPrev.val || 0);
  const calcPassivoIICurr = (passFinCurr.val || 0) + (passPermCurr.val || 0);
  const calcPassivoIIPrev = (passFinPrev.val || 0) + (passPermPrev.val || 0);

  const fpItems: StatementItem[] = blankFP.map(item => {
    if (item.id === 'fp_at_i') {
      return {
        ...item,
        valCurrent: calcAtivoICurr > 0 ? calcAtivoICurr : null,
        valPrevious: calcAtivoIPrev > 0 ? calcAtivoIPrev : null,
        cellRefCurrent: 'I128+I129',
        cellRefPrevious: 'K128+K129',
      };
    }
    if (item.id === 'fp_at_fin') {
      return {
        ...item,
        valCurrent: atFinCurr.val,
        valPrevious: atFinPrev.val,
        cellRefCurrent: atFinCurr.ref,
        cellRefPrevious: atFinPrev.ref,
        rowSource: 128,
      };
    }
    if (item.id === 'fp_at_perm') {
      return {
        ...item,
        valCurrent: atPermCurr.val,
        valPrevious: atPermPrev.val,
        cellRefCurrent: atPermCurr.ref,
        cellRefPrevious: atPermPrev.ref,
        rowSource: 129,
      };
    }
    if (item.id === 'fp_pass_ii') {
      return {
        ...item,
        valCurrent: calcPassivoIICurr > 0 ? calcPassivoIICurr : null,
        valPrevious: calcPassivoIIPrev > 0 ? calcPassivoIIPrev : null,
        cellRefCurrent: 'O128+O129',
        cellRefPrevious: 'P128+P129',
      };
    }
    if (item.id === 'fp_pass_fin') {
      return {
        ...item,
        valCurrent: passFinCurr.val,
        valPrevious: passFinPrev.val,
        cellRefCurrent: passFinCurr.ref,
        cellRefPrevious: passFinPrev.ref,
        rowSource: 128,
      };
    }
    if (item.id === 'fp_pass_perm') {
      return {
        ...item,
        valCurrent: passPermCurr.val,
        valPrevious: passPermPrev.val,
        cellRefCurrent: passPermCurr.ref,
        cellRefPrevious: passPermPrev.ref,
        rowSource: 129,
      };
    }
    if (item.id === 'fp_saldo_patr') {
      const calcSaldoCurr = calcAtivoICurr - calcPassivoIICurr;
      const calcSaldoPrev = calcAtivoIPrev - calcPassivoIIPrev;
      return {
        ...item,
        valCurrent: saldoPatrCurr.val !== null ? saldoPatrCurr.val : calcSaldoCurr,
        valPrevious: saldoPatrPrev.val !== null ? saldoPatrPrev.val : calcSaldoPrev,
        cellRefCurrent: saldoPatrCurr.ref !== '-' ? saldoPatrCurr.ref : 'O130',
        cellRefPrevious: saldoPatrPrev.ref !== '-' ? saldoPatrPrev.ref : 'P130',
        rowSource: 130,
      };
    }
    return { ...item };
  });

  // 5. EXTRAÇÃO DO QUADRO DE COMPENSAÇÕES
  // Totais: G137/H137 e M137/N137
  // Componentes Ativos: G139:H142
  // Componentes Passivos: M139:N142
  const blankComp = getBlankCompensacoesStructure();
  const compItems: StatementItem[] = blankComp.map(item => {
    let rIdx = 136; // 0-indexed for 137
    if (item.expectedRow) {
      rIdx = item.expectedRow - 1;
    }
    const isAtivoSide = item.id.startsWith('comp_at_');
    const colCurr = isAtivoSide ? 6 : 12; // G=6, M=12
    const colPrev = isAtivoSide ? 7 : 13; // H=7, N=13

    const curr = readCell(rIdx, colCurr);
    const prev = readCell(rIdx, colPrev);

    return {
      ...item,
      valCurrent: curr.val,
      valPrevious: prev.val,
      cellRefCurrent: curr.ref,
      cellRefPrevious: prev.ref,
      rowSource: rIdx + 1,
      isUnmapped: false,
      neValue: item.neKey ? (INITIAL_NE_MAPPINGS[item.neKey] ?? item.neValue) : item.neValue,
    };
  });

  // 6. EXTRAÇÃO DO DEMONSTRATIVO DO SUPERÁVIT / DÉFICIT FINANCEIRO
  // Identificação semântica por rubrica e nível hierárquico
  let sdSectionStart = -1;
  for (let r = 0; r < grid.length; r++) {
    const desc = getRowDescription(r, 1);
    const norm = normalizeLabel(desc);
    if (
      norm.includes('superavit') ||
      norm.includes('deficit financeiro') ||
      norm.includes('demonstrativo do superavit') ||
      (norm.includes('recursos nao vinculados') && r > 40) ||
      (norm.includes('recursos vinculados') && r > 40)
    ) {
      sdSectionStart = r;
      break;
    }
  }

  if (sdSectionStart === -1) {
    sdSectionStart = Math.max(0, Math.min(140, grid.length - 15));
  }

  const sdSectionEnd = Math.min(grid.length - 1, sdSectionStart + 25);

  const getSdRowData = (r: number) => {
    const desc = getRowDescription(r, 1);
    let curr = readCell(r, 11); // Col L
    if (curr.val === null && curr.rawText === '') {
      for (const colIdx of [10, 12, 8, 9, 6, 7, 13, 2, 3, 4, 5]) {
        const probe = readCell(r, colIdx);
        if (probe.val !== null || (probe.rawText && probe.rawText !== '-')) {
          curr = probe;
          break;
        }
      }
    }
    return { desc, norm: normalizeLabel(desc), curr, rIdx: r };
  };

  const sdSectionRows: ReturnType<typeof getSdRowData>[] = [];
  for (let r = sdSectionStart; r <= sdSectionEnd; r++) {
    const rowData = getSdRowData(r);
    if (rowData.desc) {
      sdSectionRows.push(rowData);
    }
  }

  const matchedRows: { [key: string]: ReturnType<typeof getSdRowData> } = {};
  let extraorcRow: ReturnType<typeof getSdRowData> | null = null;

  for (const row of sdSectionRows) {
    const { norm } = row;

    if (
      !matchedRows['sd_rec_nao_vinc'] &&
      (norm.includes('nao vinculado') ||
        norm.includes('nao vinculada') ||
        norm.includes('recursos ordinarios') ||
        (norm.includes('nao') && norm.includes('vinculad')))
    ) {
      matchedRows['sd_rec_nao_vinc'] = row;
    } else if (
      !matchedRows['sd_rec_vinc'] &&
      (norm.includes('recursos vinculados') || norm === 'vinculados' || norm === 'recursos vinculados') &&
      !norm.includes('nao')
    ) {
      matchedRows['sd_rec_vinc'] = row;
    } else if (!matchedRows['sd_educ'] && norm.includes('educac')) {
      matchedRows['sd_educ'] = row;
    } else if (
      !matchedRows['sd_seg_soc'] &&
      (norm.includes('seguridade') || (norm.includes('saude') && !norm.includes('educac')))
    ) {
      matchedRows['sd_seg_soc'] = row;
    } else if (
      !matchedRows['sd_prev_soc'] &&
      (norm.includes('previdencia') || norm.includes('rpps'))
    ) {
      matchedRows['sd_prev_soc'] = row;
    } else if (!matchedRows['sd_div_pub'] && norm.includes('divida')) {
      matchedRows['sd_div_pub'] = row;
    } else if (
      !matchedRows['sd_fundos'] &&
      (norm.includes('fundos') ||
        norm.includes('orgaos e programas') ||
        norm.includes('outros recursos') ||
        norm.includes('demais recursos'))
    ) {
      matchedRows['sd_fundos'] = row;
    } else if (
      !extraorcRow &&
      (norm.includes('extraorcamentari') ||
        norm.includes('extra orcamentari') ||
        norm.includes('extra-orcamentari'))
    ) {
      extraorcRow = row;
    } else if (
      !matchedRows['sd_total'] &&
      (norm === 'total' ||
        norm === 'total das fontes de recursos' ||
        norm === 'total geral' ||
        norm === 'total do superavit' ||
        norm.startsWith('total das fontes') ||
        (norm.startsWith('total') &&
          !norm.includes('vinculad') &&
          !norm.includes('ativo') &&
          !norm.includes('passivo')))
    ) {
      matchedRows['sd_total'] = row;
    }
  }

  const sdItems: StatementItem[] = [];

  const rNaoVinc = matchedRows['sd_rec_nao_vinc'];
  sdItems.push({
    id: 'sd_rec_nao_vinc',
    compositeKey: 'Superávit/Déficit > Recursos Não Vinculados',
    label: rNaoVinc?.desc || 'Recursos Não Vinculados',
    level: 0,
    isBold: true,
    valCurrent: rNaoVinc ? rNaoVinc.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rNaoVinc ? rNaoVinc.curr.ref : 'Pendente',
    rowSource: rNaoVinc ? rNaoVinc.rIdx + 1 : undefined,
    isUnmapped: !rNaoVinc,
  });

  const rVinc = matchedRows['sd_rec_vinc'];
  sdItems.push({
    id: 'sd_rec_vinc',
    compositeKey: 'Superávit/Déficit > Recursos Vinculados',
    label: rVinc?.desc || 'Recursos Vinculados',
    level: 0,
    isBold: true,
    isSubtotal: true,
    valCurrent: rVinc ? rVinc.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rVinc ? rVinc.curr.ref : 'Pendente',
    rowSource: rVinc ? rVinc.rIdx + 1 : undefined,
    isUnmapped: !rVinc,
  });

  const rEduc = matchedRows['sd_educ'];
  sdItems.push({
    id: 'sd_educ',
    compositeKey: 'Superávit/Déficit > Recursos Vinculados > Educação',
    label: rEduc?.desc || 'Educação',
    level: 1,
    valCurrent: rEduc ? rEduc.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rEduc ? rEduc.curr.ref : 'Pendente',
    rowSource: rEduc ? rEduc.rIdx + 1 : undefined,
    isUnmapped: !rEduc,
  });

  const rSegSoc = matchedRows['sd_seg_soc'];
  sdItems.push({
    id: 'sd_seg_soc',
    compositeKey: 'Superávit/Déficit > Recursos Vinculados > Seguridade Social (Exceto Previdência)',
    label: rSegSoc?.desc || 'Seguridade Social (Exceto Previdência)',
    level: 1,
    valCurrent: rSegSoc ? rSegSoc.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rSegSoc ? rSegSoc.curr.ref : 'Pendente',
    rowSource: rSegSoc ? rSegSoc.rIdx + 1 : undefined,
    isUnmapped: !rSegSoc,
  });

  const rPrevSoc = matchedRows['sd_prev_soc'];
  sdItems.push({
    id: 'sd_prev_soc',
    compositeKey: 'Superávit/Déficit > Recursos Vinculados > Previdência Social (RPPS)',
    label: rPrevSoc?.desc || 'Previdência Social (RPPS)',
    level: 1,
    valCurrent: rPrevSoc ? rPrevSoc.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rPrevSoc ? rPrevSoc.curr.ref : 'Pendente',
    rowSource: rPrevSoc ? rPrevSoc.rIdx + 1 : undefined,
    isUnmapped: !rPrevSoc,
  });

  const rDivPub = matchedRows['sd_div_pub'];
  sdItems.push({
    id: 'sd_div_pub',
    compositeKey: 'Superávit/Déficit > Recursos Vinculados > Dívida Pública',
    label: rDivPub?.desc || 'Dívida Pública',
    level: 1,
    valCurrent: rDivPub ? rDivPub.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rDivPub ? rDivPub.curr.ref : 'Pendente',
    rowSource: rDivPub ? rDivPub.rIdx + 1 : undefined,
    isUnmapped: !rDivPub,
  });

  const rFundos = matchedRows['sd_fundos'];
  sdItems.push({
    id: 'sd_fundos',
    compositeKey: 'Superávit/Déficit > Recursos Vinculados > Fundos, Órgãos e Programas',
    label: rFundos?.desc || 'Fundos, Órgãos e Programas',
    level: 1,
    valCurrent: rFundos ? rFundos.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rFundos ? rFundos.curr.ref : 'Pendente',
    rowSource: rFundos ? rFundos.rIdx + 1 : undefined,
    isUnmapped: !rFundos,
  });

  if (extraorcRow) {
    sdItems.push({
      id: 'sd_extraorc',
      compositeKey: 'Superávit/Déficit > Recursos Vinculados > Recursos Extraorçamentários',
      label: extraorcRow.desc || 'Recursos Extraorçamentários',
      level: 1,
      valCurrent: extraorcRow.curr.val,
      valPrevious: null,
      cellRefCurrent: extraorcRow.curr.ref,
      rowSource: extraorcRow.rIdx + 1,
      isUnmapped: false,
    });
  }

  const rTotal = matchedRows['sd_total'];
  sdItems.push({
    id: 'sd_total',
    compositeKey: 'Superávit/Déficit > TOTAL',
    label: rTotal?.desc || 'TOTAL',
    level: 0,
    isBold: true,
    isTotal: true,
    valCurrent: rTotal ? rTotal.curr.val : null,
    valPrevious: null,
    cellRefCurrent: rTotal ? rTotal.curr.ref : 'Pendente',
    rowSource: rTotal ? rTotal.rIdx + 1 : undefined,
    isUnmapped: !rTotal,
  });

  // Verifica se há rubricas não mapeadas
  const allParsed = [...ativoItems, ...passivoItems, ...fpItems, ...compItems, ...sdItems];
  const unmappedCount = allParsed.filter(x => x.isUnmapped).length;
  if (unmappedCount > 0) {
    errors.push(`Existem ${unmappedCount} rubricas obrigatórias com status "Pendente de mapeamento".`);
  }

  const parsedData: BalancoData = {
    metadata,
    ativo: ativoItems,
    passivoPL: passivoItems,
    financeiroPermanente: fpItems,
    compensacoes: compItems,
    superavitDeficit: sdItems,
    rawParsedRowsCount: grid.length,
  };

  return {
    data: parsedData,
    warnings,
    errors,
  };
}
