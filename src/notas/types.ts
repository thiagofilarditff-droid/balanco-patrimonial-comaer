export interface AccountingBalance {
  conta: string;
  nomeConta: string;
  entidadeCCor?: string;
  saldoAtual: number;
  saldoAnterior: number;
}

export interface TableRow {
  rubrica: string;
  atual: number;
  anterior: number;
  variacao: number;
  av: number;
  nota?: string;
  isSubtotal?: boolean;
}

export interface TableData {
  title: string;
  rows: TableRow[];
  total: TableRow;
  subtotals?: TableRow[];
  currentPeriod?: string;
  previousPeriod?: string;
  currentDate?: string;
  previousDate?: string;
  commentary?: string;
  noteTitle?: string;
  preCommentary?: string;
  hideAV?: boolean;
  hideAH?: boolean;
  onlyAV?: boolean;
}

