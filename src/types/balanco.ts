export interface StatementItem {
  id: string;
  compositeKey?: string; // Ex: "Ativo > Ativo Circulante > Créditos a Curto Prazo > Empréstimos (Líquido)"
  code?: string;
  label: string;
  displayLabel?: string;
  level: number; // 0 = Grupo principal, 1 = Subgrupo, 2 = Conta, 3 = Detalhe / Redutora
  isBold?: boolean;
  isUnderline?: boolean;
  isUnderlineValue?: boolean; // Se deve sublinhar também as células numéricas
  neKey?: string;
  neValue?: string | number | null;
  valCurrent: number | null;
  valPrevious: number | null;
  cellRefCurrent?: string;
  cellRefPrevious?: string;
  isRedutora?: boolean;
  isSubtotal?: boolean;
  isTotal?: boolean;
  isGroupHeader?: boolean;
  isUnmapped?: boolean; // Se a rubrica obrigatória não foi localizada no arquivo
  expectedRow?: number; // Linha exata de referência no SIAFI (1-indexed)
  rowSource?: number;
  omitIfZeroBalance?: boolean; // Ocultar da diagramação visual se saldo for nulo/zero nos dois exercícios
}

export interface StatementSection {
  id: string;
  title: string;
  showSingleYear?: boolean;
  items: StatementItem[];
  footnote?: string;
}

export interface DocumentMetadata {
  title: string;
  orgaoCode: string;
  orgaoName: string;
  exercicio: number;
  periodo: string;
  dataEmissao: string;
  unidadeOriginal: string;
  unidadeNormalizada: 'R$ milhares' | 'R$';
  unitScaleFactor: number; // 1 se já estiver em milhares, 0.001 se estiver em R$
  colCurrent: string; // Ex: '2026'
  colPrevious: string; // Ex: '2025'
  sourceSheetName: string;
  availableSheets?: string[];
  unitConfirmed: boolean;
}

export interface BalancoData {
  metadata: DocumentMetadata;
  ativo: StatementItem[];
  passivoPL: StatementItem[];
  financeiroPermanente: StatementItem[];
  compensacoes: StatementItem[];
  superavitDeficit: StatementItem[];
  rawParsedRowsCount?: number;
}

export type ValidationStatus = 'conforme' | 'aviso' | 'erro' | 'pendente';

export interface ComponentBreakdownItem {
  id: string;
  label: string;
  cellRef: string;
  valorOriginal: number;
  natureza?: 'bruto' | 'redutora' | 'subtotal' | 'parcela';
}

export interface ValidationRuleResult {
  id: string;
  quadro: string;
  exercicio: string;
  regra: string;
  celulasOrigem: string;
  valorInformado: number;
  valorCalculado: number;
  diferenca: number;
  unidade: string;
  tolerancia: number;
  situacao: ValidationStatus;
  explicacao: string;
  detalhes?: string;
  componentBreakdown?: ComponentBreakdownItem[];
}

export interface WordExportConfig {
  fontName: string; // 'Times New Roman'
  fontSizePt: number; // 10.5
  marginsCm: number; // 1.5
  primaryColorHex: string; // '#285F98' (Azul do quadro principal)
  secondaryColorHex: string; // '#31549A' (Azul dos quadros complementares)
  headerTextColorHex: string; // 'FFFFFF'
  separatePages: boolean; // Page break between sections
  showFonteSiafi: boolean;
}
