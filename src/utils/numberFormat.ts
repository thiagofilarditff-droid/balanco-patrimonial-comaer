/**
 * Commercial rounding: Half away from zero (metade para cima em magnitude, inclusive para negativos)
 * Ex:
 * 32.417.479,66 -> 32.417.480
 * 239.797.037,65 -> 239.797.038
 * -22.908.399,73 -> -22.908.400 -> exibido como (22.908.400)
 */
export function roundCommercial(val: number | null | undefined): number | null {
  if (val === null || val === undefined || isNaN(val)) {
    return null;
  }
  const sign = val < 0 ? -1 : 1;
  const absVal = Math.abs(val);
  return sign * Math.round(absVal);
}

/**
 * Format integer or commercially rounded value for the Word / display tables:
 * - Null or NaN or 0: '-'
 * - Positive: '18.380.595'
 * - Negative: '(22.908.400)'
 */
export function formatThousands(val: number | null | undefined): string {
  if (val === null || val === undefined || isNaN(val)) {
    return '-';
  }
  const rounded = roundCommercial(val);
  if (rounded === null || rounded === 0) {
    return '-';
  }

  const isNeg = rounded < 0;
  const absVal = Math.abs(rounded);
  // Formata com separador de milhar brasileiro (ponto)
  const formattedAbs = absVal.toLocaleString('pt-BR', {
    maximumFractionDigits: 0,
    minimumFractionDigits: 0,
  });

  return isNeg ? `(${formattedAbs})` : formattedAbs;
}

/**
 * Format raw number with 2 decimal places (for validation panel and audit):
 * Ex: 239.797.037,65 or -22.908.399,73
 */
export function formatDecimal(val: number | null | undefined, decimals = 2): string {
  if (val === null || val === undefined || isNaN(val)) {
    return '-';
  }
  const isNeg = val < 0;
  const absVal = Math.abs(val);
  const formatted = absVal.toLocaleString('pt-BR', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  return isNeg ? `-${formatted}` : formatted;
}

/**
 * Format currency with full precision and symbol
 */
export function formatCurrencyBRL(val: number | null | undefined): string {
  if (val === null || val === undefined || isNaN(val)) {
    return '-';
  }
  return val.toLocaleString('pt-BR', {
    style: 'currency',
    currency: 'BRL',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/**
 * Robust string to number parser that supports:
 * - Brazilian format: "239.797.037,65"
 * - Negative in parentheses: "(22.908.399,73)"
 * - Minus prefix: "-22.908.399,73"
 * - Standard US float: 239797037.65
 * - Dash "-" or empty string -> null
 */
export function parseExcelNumber(val: unknown): number | null {
  if (val === null || val === undefined) {
    return null;
  }

  if (typeof val === 'number') {
    if (isNaN(val)) return null;
    return val;
  }

  const str = String(val).trim();
  if (str === '' || str === '-' || str === '–' || str === '—' || str.toLowerCase() === 'null') {
    return null;
  }

  // Check if enclosed in parentheses: (123,45) -> negative
  let isNegative = false;
  let clean = str;
  if (clean.startsWith('(') && clean.endsWith(')')) {
    isNegative = true;
    clean = clean.slice(1, -1).trim();
  } else if (clean.startsWith('-')) {
    isNegative = true;
    clean = clean.slice(1).trim();
  }

  // Remove potential non-numeric characters except comma and dot
  clean = clean.replace(/[^0-9,.-]/g, '');

  // Handle Brazilian formatting: 1.234.567,89
  if (clean.includes(',') && clean.includes('.')) {
    // Usually dot is thousand separator, comma is decimal
    clean = clean.replace(/\./g, '').replace(',', '.');
  } else if (clean.includes(',')) {
    // Only comma -> decimal separator
    clean = clean.replace(',', '.');
  }

  const parsed = parseFloat(clean);
  if (isNaN(parsed)) return null;

  return isNegative ? -Math.abs(parsed) : parsed;
}
