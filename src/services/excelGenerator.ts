import * as XLSX from 'xlsx';

/**
 * Builds an authentic SIAFI .xlsx workbook matching the exact structure and row addresses
 * of 'RelatorioBalancoPatrimonial_896' described in the official COMAER specification.
 */
export function generateSampleExcelWorkbook(): Uint8Array {
  const wb = XLSX.utils.book_new();

  const rows: unknown[][] = [];
  const ensureRow = (r: number): unknown[] => {
    while (rows.length <= r) {
      rows.push([]);
    }
    return rows[r];
  };

  const setCell = (r: number, c: number, val: unknown) => {
    const row = ensureRow(r);
    while (row.length <= c) {
      row.push('');
    }
    row[c] = val;
  };

  // Header lines (rows 1-6 in Excel, 0-indexed: 0-5)
  setCell(1, 1, 'SIAFI - Sistema Integrado de Administração Financeira');
  setCell(2, 1, 'BALANÇO PATRIMONIAL');
  setCell(3, 1, 'Órgão: 52111 — Comando da Aeronáutica');
  setCell(4, 1, 'Exercício: 2026');
  setCell(4, 4, 'Período: MAR (FECHADO)');
  setCell(4, 8, 'Data de Emissão: 22/04/2026');
  setCell(5, 1, 'Unidade: milhares de reais');

  // Main Section Header (row 15, 0-indexed: 14)
  // Col B=1, I=8, K=10, L=11, O=14, P=15
  setCell(14, 1, 'ATIVO');
  setCell(14, 8, '2026');
  setCell(14, 10, '2025');
  setCell(14, 11, 'PASSIVO E PATRIMÔNIO LÍQUIDO');
  setCell(14, 14, '2026');
  setCell(14, 15, '2025');

  // Helper: map by 1-indexed Excel row
  const setExcelRowAtivo = (row1: number, label: string, val2026: unknown, val2025: unknown) => {
    const r = row1 - 1;
    setCell(r, 1, label);
    setCell(r, 8, val2026);
    setCell(r, 10, val2025);
  };

  const setExcelRowPassivo = (row1: number, label: string, val2026: unknown, val2025: unknown) => {
    const r = row1 - 1;
    setCell(r, 11, label);
    setCell(r, 14, val2026);
    setCell(r, 15, val2025);
  };

  // 1. ATIVO (Col B, I, K)
  setExcelRowAtivo(16, 'Ativo Circulante', 32417479.66, 31074324.0);
  setExcelRowAtivo(17, 'Caixa e Equivalentes de Caixa', 18380595.0, 17106430.0);
  setExcelRowAtivo(18, 'Créditos a Curto Prazo', 625491.0, 691068.0);
  setExcelRowAtivo(21, 'Empréstimos e Financiamentos Concedidos', 2173.0, 2363.0);
  setExcelRowAtivo(22, 'Empréstimos e Financiamentos Concedidos', 2237.0, 2429.0);
  setExcelRowAtivo(23, '(-) Perdas Estimadas em Empréstimos e Financ. Concedidos', -64.0, -66.0);
  setExcelRowAtivo(24, 'Demais Créditos e Valores a Curto Prazo', 623318.0, 688705.0);
  setExcelRowAtivo(25, 'Demais Créditos e Valores a Curto Prazo', 633642.0, 699032.0);
  setExcelRowAtivo(26, '(-) Perdas Estimadas em Demais Créditos a Curto Prazo', -10324.0, -10326.0);
  setExcelRowAtivo(27, 'Investimentos e Aplicações Temporárias a Curto Prazo', '-', '-');
  setExcelRowAtivo(28, 'Estoques', 13411394.0, 13276826.0);
  setExcelRowAtivo(31, 'Ativo Não Circulante', 207379558.0, 210380619.0);
  setExcelRowAtivo(32, 'Ativo Realizável a Longo Prazo', 33572.0, 29698.0);
  setExcelRowAtivo(33, 'Créditos a Longo Prazo', 31954.0, 28080.0);
  setExcelRowAtivo(34, 'Empréstimos e Financiamentos Concedidos', 21652.0, 17751.0);
  setExcelRowAtivo(35, 'Empréstimos e Financiamentos Concedidos', 33973.0, 30094.0);
  setExcelRowAtivo(36, '(-) Perdas Estimadas em Empréstimos e Financ. Concedidos', -12321.0, -12343.0);
  setExcelRowAtivo(37, 'Demais Créditos e Valores a Longo Prazo', 10302.0, 10329.0);
  setExcelRowAtivo(38, 'Demais Créditos e Valores a Longo Prazo', 93794.0, 94014.0);
  setExcelRowAtivo(39, '(-) Perdas Estimadas em Demais Créditos a Longo Prazo', -83492.0, -83685.0);
  setExcelRowAtivo(40, 'Estoques a Longo Prazo', 1618.0, 1618.0);
  setExcelRowAtivo(41, 'Investimentos', '-', '-');
  setExcelRowAtivo(64, 'Imobilizado', 206604799.0, 209610750.0);
  setExcelRowAtivo(66, 'Bens Móveis', 36842451.0, 37030736.0);
  setExcelRowAtivo(68, 'Bens Móveis', 39267969.0, 39379345.0);
  setExcelRowAtivo(70, '(-) Depreciação/Amortização/Exaustão Acum. de Bens Móveis', -2425518.0, -2348609.0);
  setExcelRowAtivo(74, 'Bens Imóveis', 169762348.0, 172580014.0);
  setExcelRowAtivo(76, 'Bens Imóveis', 169782055.0, 172594926.0);
  setExcelRowAtivo(78, '(-) Depreciação/Amortização/Exaustão Acum. de Bens Imóveis', -19708.0, -14911.0);
  setExcelRowAtivo(82, 'Intangível', 741187.0, 740171.0);
  setExcelRowAtivo(119, 'TOTAL DO ATIVO', 239797037.65, 241454942.65);

  // 2. PASSIVO E PATRIMÔNIO LÍQUIDO (Col L, O, P)
  setExcelRowPassivo(16, 'Passivo Circulante', 16031512.55, 17853954.92);
  setExcelRowPassivo(17, 'Obrigações Trabalh., Previd. e Assist. a Pagar a Curto Prazo', 2168354.77, 1294404.33);
  setExcelRowPassivo(18, 'Empréstimos e Financiamentos a Curto Prazo', 1507268.0, 1507108.82);
  setExcelRowPassivo(19, 'Fornecedores e Contas a Pagar a Curto Prazo', 169329.86, 144701.23);
  setExcelRowPassivo(20, 'Obrigações Fiscais a Curto Prazo', '-', '-');
  setExcelRowPassivo(21, 'Transferências Fiscais a Curto Prazo', 5445.69, '-');
  setExcelRowPassivo(22, 'Provisões a Curto Prazo', 10666396.75, 14043211.32);
  setExcelRowPassivo(23, 'Demais Obrigações a Curto Prazo', 1514717.48, 864529.22);
  setExcelRowPassivo(31, 'Passivo Não Circulante', 202912955.08, 202913351.83);
  setExcelRowPassivo(32, 'Obrigações Trabalh., Previd. e Assist. a Pagar a Longo Prazo', '-', '-');
  setExcelRowPassivo(33, 'Empréstimos e Financiamentos a Longo Prazo', 16584480.94, 16584480.94);
  setExcelRowPassivo(34, 'Fornecedores e Contas a Pagar a Longo Prazo', '-', '-');
  setExcelRowPassivo(35, 'Obrigações Fiscais a Longo Prazo', '-', '-');
  setExcelRowPassivo(36, 'Transferências Fiscais a Longo Prazo', '-', '-');
  setExcelRowPassivo(37, 'Provisões a Longo Prazo', 186326780.49, 186326780.49);
  setExcelRowPassivo(38, 'Demais Obrigações a Longo Prazo', 1529.03, 1563.12);
  setExcelRowPassivo(39, 'Resultado Diferido', 164.62, 527.28);
  setExcelRowPassivo(40, 'Total do Passivo Exigível', 218944467.63, 220767306.75);
  setExcelRowPassivo(41, 'PATRIMÔNIO LÍQUIDO', '', ''); // Linha 41 é título sem saldo
  setExcelRowPassivo(45, 'Patrimônio Social', '-', '-');
  setExcelRowPassivo(51, 'Ajustes de Avaliação Patrimonial', -22908399.73, -22908400.0);
  setExcelRowPassivo(55, 'Demais Reservas', 99736060.0, 102339620.0);
  setExcelRowPassivo(57, 'Resultados Acumulados', -55975091.0, -58743584.0);
  setExcelRowPassivo(67, 'TOTAL DO PATRIMÔNIO LÍQUIDO', 20852570.02, 20687635.9);
  setExcelRowPassivo(119, 'TOTAL DO PASSIVO E PATRIMÔNIO LÍQUIDO', 239797037.65, 241454942.65);

  // 3. QUADRO FINANCEIRO E PERMANENTE (Linhas 128, 129, 130)
  // Ativo Financeiro (I128/K128), Ativo Permanente (I129/K129)
  setCell(127, 1, 'Ativo Financeiro');
  setCell(127, 8, 18380595.0);
  setCell(127, 10, 17106430.0);

  setCell(128, 1, 'Ativo Permanente');
  setCell(128, 8, 221416442.0);
  setCell(128, 10, 224348512.0);

  // Passivo Financeiro (O128/P128), Passivo Permanente (O129/P129)
  setCell(127, 11, 'Passivo Financeiro');
  setCell(127, 14, 6496799.0);
  setCell(127, 15, 5039533.0);

  setCell(128, 11, 'Passivo Permanente');
  setCell(128, 14, 216233237.0);
  setCell(128, 15, 219096138.0);

  // Saldo Patrimonial (O130/P130)
  setCell(129, 11, 'Saldo Patrimonial');
  setCell(129, 14, 17067002.0);
  setCell(129, 15, 17319272.0);

  // 4. QUADRO DE COMPENSAÇÕES (Linhas 137, 139..142)
  // Atos Ativos em G=6 e H=7; Atos Passivos em M=12 e N=13
  setCell(136, 1, 'Atos Potenciais Ativos');
  setCell(136, 6, 16101873.0); // G137
  setCell(136, 7, 16504357.0); // H137

  setCell(136, 11, 'Atos Potenciais Passivos');
  setCell(136, 12, 47338561.0); // M137
  setCell(136, 13, 47191004.0); // N137

  setCell(138, 1, 'Garantias e Contragarantias Recebidas a Executar');
  setCell(138, 6, 6328736.0); // G139
  setCell(138, 7, 6634886.0); // H139

  setCell(138, 11, 'Garantias e Contragarantias Concedidas a Executar');
  setCell(138, 12, '-'); // M139
  setCell(138, 13, '-'); // N139

  setCell(139, 1, 'Direitos Conveniados e Outros Instrumentos Congêneres');
  setCell(139, 6, 346171.0); // G140
  setCell(139, 7, 462772.0); // H140

  setCell(139, 11, 'Obrigações Conveniadas e Outros Instrum Congêneres');
  setCell(139, 12, 717680.0); // M140
  setCell(139, 13, 748526.0); // N140

  setCell(140, 1, 'Direitos Contratuais');
  setCell(140, 6, 9426965.0); // G141
  setCell(140, 7, 9406699.0); // H141

  setCell(140, 11, 'Obrigações Contratuais');
  setCell(140, 12, 46620881.0); // M141
  setCell(140, 13, 46442478.0); // N141

  setCell(141, 1, 'Outros Atos Potenciais Ativos a Executar');
  setCell(141, 6, '-'); // G142
  setCell(141, 7, '-'); // H142

  // 5. SUPERÁVIT/DÉFICIT FINANCEIRO (Linhas 149..156, Col B desc, Col L=11 valor)
  setCell(148, 1, 'Recursos Não Vinculados');
  setCell(148, 11, -1084025.0); // L149

  setCell(149, 1, 'Recursos Vinculados');
  setCell(149, 11, 12968051.0); // L150

  setCell(150, 1, 'Educação');
  setCell(150, 11, -63.0); // L151

  setCell(151, 1, 'Seguridade Social (Exceto Previdência)');
  setCell(151, 11, -19523.0); // L152

  setCell(152, 1, 'Previdência Social (RPPS)');
  setCell(152, 11, -9045.0); // L153

  setCell(153, 1, 'Dívida Pública');
  setCell(153, 11, -199.0); // L154

  setCell(154, 1, 'Fundos, Órgãos e Programas');
  setCell(154, 11, 12996881.0); // L155

  setCell(155, 1, 'TOTAL');
  setCell(155, 11, 11884026.0); // L156

  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'RelatorioBalancoPatrimonial_896');

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(wbout);
}
