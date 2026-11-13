import * as XLSX from 'xlsx';
import { AccountingBalance, TableData, TableRow } from '../types';

export interface ParseResult {
  balances: AccountingBalance[];
  currentPeriod: string;
  previousPeriod: string;
  currentDate?: string;
  previousDate?: string;
}

export const formatDateDDMMYYYY = (val?: string, rawVal?: any): string => {
  if (typeof rawVal === 'number' && rawVal > 20000 && rawVal < 70000) {
    try {
      const parsedDate = XLSX.SSF.parse_date_code(rawVal);
      if (parsedDate && parsedDate.y && parsedDate.m && parsedDate.d) {
        const dd = String(parsedDate.d).padStart(2, '0');
        const mm = String(parsedDate.m).padStart(2, '0');
        const yyyy = String(parsedDate.y);
        return `${dd}/${mm}/${yyyy}`;
      }
    } catch (e) {
      // fallback to string parsing
    }
  }

  const str = String(rawVal || val || '').trim();
  if (!str) return '';

  // 1. dd/mm/yyyy or dd-mm-yyyy or dd.mm.yyyy
  const dmyMatch = str.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
  if (dmyMatch) {
    return `${dmyMatch[1].padStart(2, '0')}/${dmyMatch[2].padStart(2, '0')}/${dmyMatch[3]}`;
  }

  // 2. dd/mm/yy or dd-mm-yy
  const dmy2Match = str.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2})/);
  if (dmy2Match) {
    return `${dmy2Match[1].padStart(2, '0')}/${dmy2Match[2].padStart(2, '0')}/20${dmy2Match[3]}`;
  }

  // 3. yyyy-mm-dd
  const ymdMatch = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymdMatch) {
    return `${ymdMatch[3].padStart(2, '0')}/${ymdMatch[2].padStart(2, '0')}/${ymdMatch[1]}`;
  }

  // 4. By month name and year
  const lower = str.toLowerCase();
  const monthsMap: Record<string, number> = {
    'jan': 1, 'janeiro': 1,
    'fev': 2, 'fevereiro': 2,
    'mar': 3, 'marco': 3, 'março': 3,
    'abr': 4, 'abril': 4,
    'mai': 5, 'maio': 5,
    'jun': 6, 'junho': 6,
    'jul': 7, 'julho': 7,
    'ago': 8, 'agosto': 8,
    'set': 9, 'setembro': 9,
    'out': 10, 'outubro': 10,
    'nov': 11, 'novembro': 11,
    'dez': 12, 'dezembro': 12
  };

  let monthNum = 0;
  for (const [key, num] of Object.entries(monthsMap)) {
    if (lower.includes(key)) {
      monthNum = num;
      break;
    }
  }

  const nums = str.match(/\d+/g) || [];
  let yearNum = 0;
  const y4 = nums.find(n => n.length === 4);
  if (y4) {
    yearNum = parseInt(y4, 10);
  } else if (nums.length > 0) {
    const lastNum = nums[nums.length - 1];
    if (lastNum.length === 2) {
      yearNum = 2000 + parseInt(lastNum, 10);
    }
  }

  if (!monthNum && nums.length >= 2) {
    const pot = nums.find((n, i) => i < nums.length - 1 && parseInt(n, 10) >= 1 && parseInt(n, 10) <= 12);
    if (pot) monthNum = parseInt(pot, 10);
  }

  if (monthNum && yearNum) {
    const lastDay = new Date(yearNum, monthNum, 0).getDate();
    return `${String(lastDay).padStart(2, '0')}/${String(monthNum).padStart(2, '0')}/${yearNum}`;
  }

  return str;
};

export const formatPeriod = (period: string) => {
  if (!period) return period;
  const p = period.trim();
  const lower = p.toLowerCase();
  
  if (lower === 'atual' || lower === 'anterior' || lower === 'saldo atual' || lower === 'saldo anterior') return p;
  
  const months: Record<string, string> = {
    'janeiro': 'jan', 'fevereiro': 'fev', 'março': 'mar', 'abril': 'abr',
    'maio': 'mai', 'junho': 'jun', 'julho': 'jul', 'agosto': 'ago',
    'setembro': 'set', 'outubro': 'out', 'novembro': 'nov', 'dezembro': 'dez',
    'jan': 'jan', 'fev': 'fev', 'mar': 'mar', 'abr': 'abr', 'mai': 'mai',
    'jun': 'jun', 'jul': 'jul', 'ago': 'ago', 'set': 'set', 'out': 'out',
    'nov': 'nov', 'dez': 'dez'
  };

  const numericalMonths: Record<string, string> = {
    '01': 'jan', '02': 'fev', '03': 'mar', '04': 'abr', '05': 'mai', '06': 'jun',
    '07': 'jul', '08': 'ago', '09': 'set', '10': 'out', '11': 'nov', '12': 'dez',
    '1': 'jan', '2': 'fev', '3': 'mar', '4': 'abr', '5': 'mai', '6': 'jun',
    '7': 'jul', '8': 'ago', '9': 'set'
  };

  let foundMonth = '';
  // 1. Check for textual months
  for (const [full, short] of Object.entries(months)) {
    if (lower.includes(full)) {
      foundMonth = short;
      break;
    }
  }

  // 2. Extract all numbers
  const nums = p.match(/\d+/g) || [];
  
  // 3. Find month numerically if not found textually
  if (!foundMonth && nums.length > 0) {
    if (nums.length >= 3) {
      // 31/03/2026 -> Usually index 1 is month
      const potential = nums.find((n, i) => i > 0 && i < nums.length - 1 && parseInt(n) >= 1 && parseInt(n) <= 12);
      if (potential) foundMonth = numericalMonths[potential] || numericalMonths[String(parseInt(potential))];
      else {
        // Fallback to any number <= 12 that isn't the year
        const fallback = nums.find(n => n.length <= 2 && parseInt(n) >= 1 && parseInt(n) <= 12);
        if (fallback) foundMonth = numericalMonths[fallback] || numericalMonths[String(parseInt(fallback))];
      }
    } else if (nums.length === 2) {
      // 03/2026 -> First one is often month
      const potential = nums.find(n => parseInt(n) >= 1 && parseInt(n) <= 12);
      if (potential) foundMonth = numericalMonths[potential] || numericalMonths[String(parseInt(potential))];
    } else if (nums.length === 1 && nums[0].length <= 2) {
       // Just a month number? Unlikely but possible
       foundMonth = numericalMonths[nums[0]] || numericalMonths[String(parseInt(nums[0]))];
    }
  }

  // 4. Find Year
  let foundYear = '';
  const year4 = nums.find(n => n.length === 4);
  const year2 = nums.find(n => n.length === 2 && (!foundMonth || numericalMonths[n] !== foundMonth));

  if (year4) {
    foundYear = year4.slice(-2);
  } else if (year2) {
    foundYear = year2;
  } else if (nums.length > 0 && nums[nums.length-1].length >= 2) {
    foundYear = nums[nums.length-1].slice(-2);
  }

  if (foundMonth && foundYear) {
    return `${foundMonth}/${foundYear}`;
  }

  return period;
};

export const parseExcelBalances = (file: File): Promise<ParseResult> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true }) as any[][];
        
        if (rows.length < 2) {
          resolve({ balances: [], currentPeriod: 'Atual', previousPeriod: 'Anterior' });
          return;
        }

        // Find header row and periods
        let headerIdx = -1;
        let currentPeriod = 'Atual';
        let previousPeriod = 'Anterior';
        
        // Strict mapping:
        // Coluna A (index 0): Conta
        // Coluna B (index 1): Nome da Conta / Denominação
        // Coluna D (index 3): Entidade CCor
        // Coluna E (index 4): Saldo Atual
        // Coluna F (index 5): Saldo Anterior
        const colConta = 0;
        const colNome = 1;
        const colEntidade = 3;
        const colAtual = 4;
        const colAnterior = 5;

        for (let i = 0; i < Math.min(rows.length, 15); i++) {
          const rowLower = rows[i].map(c => String(c || '').toLowerCase().trim());
          if (rowLower.some(c => c.includes('conta') || c.includes('código') || c.includes('codigo'))) {
            headerIdx = i;
            break;
          }
        }

        if (headerIdx === -1) headerIdx = 1; // Fallback

        // Periods from Coluna E (index 4) and Coluna F (index 5)
        let rawC2 = rows[headerIdx]?.[colAtual];
        let rawC3 = rows[headerIdx]?.[colAnterior];
        let c2 = formatPeriod(String(rawC2 || '').trim());
        let c3 = formatPeriod(String(rawC3 || '').trim());

        // If periods were not in header, scan rows above
        if ((c2 === 'Atual' || c2 === 'Saldo Atual' || !c2.includes('/')) && headerIdx > 0) {
          for (let j = headerIdx - 1; j >= 0; j--) {
            const pot = formatPeriod(String(rows[j]?.[colAtual] || '').trim());
            if (pot.includes('/')) {
              c2 = pot;
              rawC2 = rows[j]?.[colAtual];
              break;
            }
          }
        }
        if ((c3 === 'Anterior' || c3 === 'Saldo Anterior' || !c3.includes('/')) && headerIdx > 0) {
          for (let j = headerIdx - 1; j >= 0; j--) {
            const pot = formatPeriod(String(rows[j]?.[colAnterior] || '').trim());
            if (pot.includes('/')) {
              c3 = pot;
              rawC3 = rows[j]?.[colAnterior];
              break;
            }
          }
        }

        currentPeriod = c2 || 'Atual';
        previousPeriod = c3 || 'Anterior';

        const parseVal = (val: any): number => {
          if (val === null || val === undefined) return 0;
          if (typeof val === 'number') {
            return isNaN(val) ? 0 : val;
          }
          
          let str = String(val).trim();
          if (str === '' || str === '-' || str === '.') return 0;
          
          let isNegative = false;
          if (str.endsWith('-') || str.startsWith('-') || (str.startsWith('(') && str.endsWith(')'))) {
            isNegative = true;
          }
          
          // Remove currency symbols, parentheses, trailing/leading minus, spaces
          str = str.replace(/[R$\s()\-]/g, '');
          
          if (str.includes(',') && str.includes('.')) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
              str = str.replace(/\./g, '').replace(',', '.');
            } else {
              str = str.replace(/,/g, '');
            }
          } else if (str.includes(',')) {
            str = str.replace(',', '.');
          }
          
          const num = parseFloat(str);
          if (isNaN(num)) return 0;
          return isNegative ? -num : num;
        };

        const balances: AccountingBalance[] = rows.slice(headerIdx + 1).map((row) => {
          const rawConta = String(row[colConta] || '').trim();
          const conta = rawConta.replace(/\D/g, '');
          const nomeConta = String(row[colNome] || '').trim();
          const entidadeCCor = String(row[colEntidade] || '').trim();
          
          // STRICT RULE: Coluna E (index 4) = Período Atual, Coluna F (index 5) = Período Anterior
          const saldoAtual = parseVal(row[colAtual]);
          const saldoAnterior = parseVal(row[colAnterior]);

          return { conta, nomeConta, entidadeCCor, saldoAtual, saldoAnterior };
        }).filter(b => b.conta.length > 0 || (b.nomeConta.length > 0 && (b.saldoAtual !== 0 || b.saldoAnterior !== 0)));

        const currentDate = formatDateDDMMYYYY(currentPeriod, rawC2);
        const previousDate = formatDateDDMMYYYY(previousPeriod, rawC3);

        resolve({ balances, currentPeriod, previousPeriod, currentDate, previousDate });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const aggregateBalancesByAccount = (
  balancesList: AccountingBalance[]
): { conta: string; nomeConta: string; saldoAtual: number; saldoAnterior: number }[] => {
  const map = new Map<string, { conta: string; nomeConta: string; saldoAtual: number; saldoAnterior: number }>();
  
  for (const b of balancesList) {
    const key = b.conta || b.nomeConta;
    const existing = map.get(key);
    if (existing) {
      existing.saldoAtual += b.saldoAtual;
      existing.saldoAnterior += b.saldoAnterior;
    } else {
      map.set(key, {
        conta: b.conta,
        nomeConta: b.nomeConta,
        saldoAtual: b.saldoAtual,
        saldoAnterior: b.saldoAnterior
      });
    }
  }
  
  return Array.from(map.values());
};

const calculateTableData = (
  title: string,
  balances: AccountingBalance[],
  rules: { label: string; matcher: (b: AccountingBalance) => boolean }[],
  currentPeriod: string,
  previousPeriod: string
): TableData => {
  const rows: TableRow[] = rules.map(rule => {
    const filtered = balances.filter(rule.matcher);
    
    const atual = filtered.reduce((acc, b) => acc + b.saldoAtual, 0);
    const anterior = filtered.reduce((acc, b) => acc + b.saldoAnterior, 0);
    const variacao = Math.abs(anterior) > 0.001 ? ((atual - anterior) / Math.abs(anterior)) * 100 : 0;
    
    return {
      rubrica: rule.label,
      atual,
      anterior,
      variacao,
      av: 0 
    };
  });

  const totalAtual = rows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior = rows.reduce((acc, r) => acc + r.anterior, 0);
  const totalVariacao = Math.abs(totalAnterior) > 0.001 ? ((totalAtual - totalAnterior) / Math.abs(totalAnterior)) * 100 : 0;

  const toSentenceCase = (str: string) => {
    if (!str) return str;
    const cleaned = str.trim();
    return cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
  };

  const rowsWithAV = rows.map(r => ({
    ...r,
    rubrica: r.rubrica, // Maintain original casing for specific requirements
    av: Math.abs(totalAtual) > 0.001 ? (r.atual / totalAtual) * 100 : 0
  })).sort((a, b) => b.atual - a.atual);

  return {
    title,
    rows: rowsWithAV,
    total: {
      rubrica: 'TOTAL GERAL',
      atual: totalAtual,
      anterior: totalAnterior,
      variacao: totalVariacao,
      av: Math.abs(totalAtual) > 0.001 ? 100 : 0
    },
    currentPeriod,
    previousPeriod
  };
};

export const generateReports = (result: ParseResult): TableData[] => {
  const { balances, currentPeriod, previousPeriod } = result;

  const formatCurrencySimple = (value: number) => {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: 'BRL',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(value);
  };

  const formatCurrencyThousands = (value: number) => {
    return new Intl.NumberFormat('pt-BR', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    }).format(value / 1000);
  };

  const toSentenceCase = (str: string) => {
    if (!str) return str;
    const cleaned = str.trim();
    return cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
  };

  const formattedDate = currentPeriod;

  // --- TABELA 1 ---
  const table1Detail = calculateTableData(
    "Tabela 1: Caixa e Equivalentes de Caixa – Composição",
    balances,
    [
      { label: "Moeda Estrangeira", matcher: (b) => b.conta.startsWith('11121') || b.nomeConta.toLowerCase().includes('estrangeira') },
      { label: "Moeda Nacional", matcher: (b) => (b.conta.startsWith('111') && !b.conta.startsWith('11121')) || b.nomeConta.toLowerCase().includes('moeda nacional') }
    ],
    currentPeriod,
    previousPeriod
  );

  // --- TABELA 2 ---
  const table2Detail = calculateTableData(
    "Tabela 2: Caixa e Equivalentes de Caixa em Moeda Nacional – Composição",
    balances,
    [
      { label: "Aplicações financeiras", matcher: (b) => b.conta.startsWith('1111150') || b.conta.startsWith('111110206') || b.nomeConta.toLowerCase().includes('aplicaç') },
      { label: "Bancos", matcher: (b) => b.conta.startsWith('1111119') || b.nomeConta.toLowerCase().includes('banco') },
      { label: "Garantias", matcher: (b) => b.conta.startsWith('1113102') || b.nomeConta.toLowerCase().includes('garantia') },
      { label: "Recursos - Limite de saque", matcher: (b) => b.conta.startsWith('111122001') || b.conta.startsWith('111122003') || b.nomeConta.toLowerCase().includes('limite de saque') }
    ],
    currentPeriod,
    previousPeriod
  );

  // --- TABELA 3 ---
  const table3Detail = calculateTableData(
    "Tabela 3: Aplicações Financeiras – Composição",
    balances,
    [
      { label: "Recursos da conta única aplicados", matcher: (b) => b.conta.startsWith('111110206') || b.nomeConta.toLowerCase().includes('conta única') || b.nomeConta.toLowerCase().includes('conta unica') },
      { label: "Poupança", matcher: (b) => b.conta.startsWith('111115005') || b.nomeConta.toLowerCase().includes('poupança') || b.nomeConta.toLowerCase().includes('poupanca') },
      { label: "Fundos de aplicação financeira", matcher: (b) => b.conta.startsWith('111115006') || b.nomeConta.toLowerCase().includes('fundos de aplicação') },
      { label: "Certificados de depósitos bancários - CDB", matcher: (b) => b.conta.startsWith('111115010') || b.nomeConta.toLowerCase().includes('cdb') || b.nomeConta.toLowerCase().includes('depósitos bancários') }
    ],
    currentPeriod,
    previousPeriod
  );

  // --- TABELA 4 ---
  const table4Raw = calculateTableData(
    "Tabela 4: Créditos a Curto Prazo - composição",
    balances,
    [
      { label: "Empréstimos e financiamentos concedidos a receber", matcher: (b) => b.conta.startsWith('1124101') || b.conta.startsWith('112410201') || (b.nomeConta.toLowerCase().includes('empréstimo') && !b.nomeConta.toLowerCase().includes('ajuste')) },
      { label: "Ajuste para Perdas em Emprest. e Finaciament. Concedidos", matcher: (b) => b.conta.startsWith('1129104') || (b.nomeConta.toLowerCase().includes('ajuste') && b.nomeConta.toLowerCase().includes('emprest')) },
      { label: "Adiantamentos Concedidos", matcher: (b) => b.conta.startsWith('1131') || b.nomeConta.toLowerCase().includes('adiantamento') },
      { label: "Créditos por Dano ao Patrimônio", matcher: (b) => b.conta.startsWith('1134') || (b.nomeConta.toLowerCase().includes('dano') && !b.nomeConta.toLowerCase().includes('ajuste')) },
      { label: "Ajuste p/ Perda de Créditos por dano patrimônio", matcher: (b) => b.conta.startsWith('1139') || (b.nomeConta.toLowerCase().includes('ajuste') && b.nomeConta.toLowerCase().includes('dano')) },
      { label: "Depósitos e valores vinculados a receber", matcher: (b) => b.conta.startsWith('1135') || b.nomeConta.toLowerCase().includes('depósito') || b.nomeConta.toLowerCase().includes('vinculado') },
      { label: "Rendimentos das Aplicações Financeiras", matcher: (b) => b.conta.startsWith('1138148') || b.nomeConta.toLowerCase().includes('rendimento') },
      { label: "Outros Créditos a Curto Prazo", matcher: (b) => (b.conta.startsWith('1138') && !b.conta.startsWith('1138148')) || (b.conta.startsWith('113') && !b.conta.startsWith('1131') && !b.conta.startsWith('1134') && !b.conta.startsWith('1135') && !b.conta.startsWith('1138') && !b.conta.startsWith('1139')) || b.nomeConta.toLowerCase().includes('outros créditos') }
    ],
    currentPeriod,
    previousPeriod
  );

  // --- TABELA 5 --- (Estoques)
  const table5Detail = calculateTableData(
    "Tabela 5: Estoques - composição",
    balances,
    [
      { label: "Materiais de Consumo", matcher: (b) => b.conta.startsWith("1156101") || b.nomeConta.toLowerCase().includes("consumo") },
      { label: "Estoques sobressalentes a reparar", matcher: (b) => b.conta.startsWith("1156104") || b.nomeConta.toLowerCase().includes("a reparar") },
      { label: "Estoques sobressalentes em reparo", matcher: (b) => b.conta.startsWith("1156103") || b.nomeConta.toLowerCase().includes("em reparo") },
      { label: "Estoques sobressalentes a alienar", matcher: (b) => b.conta.startsWith("1156102") || b.nomeConta.toLowerCase().includes("a alienar") },
      { label: "Estoque Interno de Material para distrib.", matcher: (b) => b.conta.startsWith("1158102") || b.nomeConta.toLowerCase().includes("distrib") },
      { label: "Demais", matcher: (b) => b.conta.startsWith("115") && !b.conta.startsWith("1156101") && !b.conta.startsWith("1156104") && !b.conta.startsWith("1156103") && !b.conta.startsWith("1156102") && !b.conta.startsWith("1158102") }
    ],
    currentPeriod,
    previousPeriod
  );

  // --- TABELA 6 --- (Longo Prazo)
  const table6Raw = calculateTableData(
    "Tabela 6: Ativo Realizável a Longo Prazo - composição",
    balances,
    [
      { label: "Empréstimos e financiamentos concedidos", matcher: (b) => b.conta.startsWith("1211103") || (b.conta.startsWith("1211") && !b.nomeConta.toLowerCase().includes("ajuste")) },
      { label: "Ajuste para Perdas em Emprest. e Finaciament. Concedidos", matcher: (b) => b.conta.startsWith("121119904") || (b.conta.startsWith("1211") && b.nomeConta.toLowerCase().includes("ajuste")) },
      { label: "Demais créditos e valores a longo prazo", matcher: (b) => (b.conta.startsWith("1212") && !b.conta.startsWith("121219902")) || (b.conta.startsWith("1212") && !b.nomeConta.toLowerCase().includes("ajuste")) },
      { label: "Ajuste para Perdas de demais créditos e valores", matcher: (b) => b.conta.startsWith("121219902") || (b.conta.startsWith("1212") && b.nomeConta.toLowerCase().includes("ajuste")) },
      { label: "Estoques", matcher: (b) => b.conta.startsWith("1214") || (b.conta.startsWith("121") && b.nomeConta.toLowerCase().includes("estoque")) }
    ],
    currentPeriod,
    previousPeriod
  );

  // --- ASSEMBLE BALANCE SHEET (Summary Table) ---
  const balanceSheetRows: TableRow[] = [
    { ...table1Detail.total, rubrica: "Caixa e Equivalentes de Caixa", nota: "1", av: 0 },
    { ...table4Raw.total, rubrica: "Créditos a Curto Prazo", nota: "2", av: 0 },
    { ...table5Detail.total, rubrica: "Estoques", nota: "3", av: 0 },
    { ...table6Raw.total, rubrica: "Ativo Realizável a Longo Prazo", nota: "4", av: 0 }
  ];

  const totalAtivoAtual = balanceSheetRows.reduce((acc, r) => acc + r.atual, 0);
  const table0 = {
    title: "Ativo - Resumo",
    rows: balanceSheetRows.map(r => ({
      ...r,
      av: Math.abs(totalAtivoAtual) > 0 ? (r.atual / totalAtivoAtual) * 100 : 0
    })).sort((a, b) => b.atual - a.atual),
    total: {
      rubrica: "TOTAL DO ATIVO",
      atual: totalAtivoAtual,
      anterior: balanceSheetRows.reduce((acc, r) => acc + r.anterior, 0),
      variacao: 0, 
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "A demonstração do Ativo compreende os bens e direitos que são controlados pela entidade como resultado de eventos passados e dos quais se espera que fluam benefícios econômicos futuros ou potencial de serviços.",
    preCommentary: "Abaixo apresentamos o resumo do Ativo circulante e não circulante, com referências às suas respectivas Notas Explicativas."
  };
  table0.total.variacao = Math.abs(table0.total.anterior) > 0 ? ((table0.total.atual - table0.total.anterior) / Math.abs(table0.total.anterior)) * 100 : 0;

  // Now restore all previous details and commentaries...
  const table1 = table1Detail;
  table1.noteTitle = "Nota Explicativa 1: Caixa e Equivalente de Caixa";
  table1.preCommentary = "A conta Caixa e Equivalentes de Caixa registra os recursos financeiros disponíveis para aplicação nas operações da entidade. Compreende o numerário e outros bens e direitos com alta capacidade de conversibilidade em moeda, estando segmentada entre Moeda nacional e Moeda estrangeira.";

  const moedaNacional = table1.rows.find(r => r.rubrica.toLowerCase() === "moeda nacional");
  if (moedaNacional) {
    const trend = moedaNacional.variacao >= 0 ? "crescimento" : "diminuição";
    const absVar = Math.abs(moedaNacional.variacao).toFixed(2);
    const av = moedaNacional.av.toFixed(2);
    table1.commentary = `A conta Caixa e Equivalentes de Caixa registra os recursos financeiros disponíveis para aplicação nas operações da entidade. Compreende o numerário e outros bens e direitos com alta capacidade de conversibilidade em moeda, estando segmentada entre Moeda nacional e Moeda estrangeira.
A moeda nacional apresentou ${trend} de ${absVar}%, consolidando-se como o principal componente da conta, com ${av}% de representatividade no total. Esse item é composto por Aplicações financeiras, Recursos do limite de saque da conta única e Bancos.`;
  }

  const table2 = table2Detail;
  const mostRepresentativeT2 = [...table2.rows].sort((a, b) => b.av - a.av)[0];
  if (mostRepresentativeT2) {
    table2.commentary = `No item analisado, destacam-se as ${mostRepresentativeT2.rubrica.toLowerCase()} que representam cerca de ${mostRepresentativeT2.av.toFixed(2)}% do montante. As aplicações financeiras são efetuadas e gerenciadas, majoritariamente, pelo Fundo Aeronáutico, conforme legislação: Medida Provisória nº 2.170-36, de 23 de agosto de 2001, Portaria nº 345 MF, de 29 de dezembro de 1998, Decreto nº 93.872, de 23 dezembro de 1986 e Instrução Normativa STN nº 4, de 30 de agosto de 2004, sendo o montante aproximado de 1% aplicado e gerenciado pela CFIAe. As aplicações financeiras concentram-se, principalmente, em CDB - Certificados de depósitos bancários.`;
  }

  const table3 = table3Detail;
  table3.commentary = `A conta “CDB – Certificados de depósitos bancários” contempla aplicações junto às principais instituições financeiras do país, com reconhecida solidez. Os vencimentos são programados de forma a garantir liquidez diária, o que assegura a compatibilidade com o fluxo de caixa necessário às operações do COMAER, sem risco de liquidez em caso de resgate antecipado. 
O item “CTU – Recursos da conta única aplicados” refere-se ao capital de giro utilizado pela Setorial Financeira do COMAER para atender às demandas das Unidades Gestoras Executoras, com recursos provenientes do Fundo Aeronáutico. Aproximadamente 1,3% desse valor é aplicado pela CFIAe.
O item “Fundo de aplicação financeira” representa aplicações em dois fundos específicos, destinados à manutenção e ao desenvolvimento de projetos estratégicos.`;

  // --- RE-PROCESS TABLE 4 (Note 2) ---
  const parentMap: Record<string, string[]> = {
    "Empréstimos e financiamentos concedidos a receber": ["Ajuste para Perdas em Emprest. e Finaciament. Concedidos"],
    "Créditos por Dano ao Patrimônio": ["Ajuste p/ Perda de Créditos por dano patrimônio"]
  };
  const adjustmentToParent: Record<string, string> = {
    "Ajuste para Perdas em Emprest. e Finaciament. Concedidos": "Empréstimos e financiamentos concedidos a receber",
    "Ajuste p/ Perda de Créditos por dano patrimônio": "Créditos por Dano ao Patrimônio"
  };
  const primaryRows = table4Raw.rows.filter(r => !adjustmentToParent[r.rubrica]);
  primaryRows.sort((a, b) => b.atual - a.atual);
  const finalRows4: TableRow[] = [];
  primaryRows.forEach(pRow => {
    finalRows4.push(pRow);
    const adjustments = parentMap[pRow.rubrica];
    if (adjustments) {
      adjustments.forEach(adjName => {
        const adjRow = table4Raw.rows.find(r => r.rubrica === adjName);
        if (adjRow) finalRows4.push(adjRow);
      });
      if (pRow.rubrica === "Empréstimos e financiamentos concedidos a receber") {
         const empAdj = table4Raw.rows.find(r => r.rubrica === "Ajuste para Perdas em Emprest. e Finaciament. Concedidos");
         finalRows4.push({ rubrica: "Subtotal Empréstimos e Financiamentos (Líquido)", atual: pRow.atual + (empAdj?.atual || 0), anterior: pRow.anterior + (empAdj?.anterior || 0), variacao: 0, av: 0, isSubtotal: true });
      } else if (pRow.rubrica === "Créditos por Dano ao Patrimônio") {
         const credAdj = table4Raw.rows.find(r => r.rubrica === "Ajuste p/ Perda de Créditos por dano patrimônio");
         finalRows4.push({ rubrica: "Subtotal Créditos por Dano ao Patrimônio (Líquido)", atual: pRow.atual + (credAdj?.atual || 0), anterior: pRow.anterior + (credAdj?.anterior || 0), variacao: 0, av: 0, isSubtotal: true });
      }
    }
  });

  const table4 = { ...table4Raw, rows: finalRows4 };
  table4.noteTitle = "Nota Explicativa 2: Créditos a Curto Prazo";
  table4.preCommentary = `Os créditos a curto prazo compreendem os valores a receber realizáveis até 12 meses da data das demonstrações contábeis. Em ${currentPeriod}, esse grupo totalizou aproximadamente ${formatCurrencySimple(table4.total.atual)}, sendo composto, principalmente, pelas rubricas ${primaryRows.slice(0,2).map(r=>r.rubrica).join(" e ")}, conforme detalhado na Tabela 4 abaixo:`;
  const empAdj4 = table4Raw.rows.find(r => r.rubrica === "Ajuste para Perdas em Emprest. e Finaciament. Concedidos");
  const totalLiqEmprestimos4 = (table4Raw.rows.find(r => r.rubrica === "Empréstimos e financiamentos concedidos a receber")?.atual || 0) + (empAdj4?.atual || 0);

  table4.commentary = `A rubrica Adiantamentos concedidos refere-se a valores antecipados no âmbito de contratações internacionais destinadas à aquisição de bens e serviços estratégicos. 
  
O item Créditos por Dano ao Patrimônio contempla o ajuste para perdas segundo critérios do TCU.

A rubrica Rendimentos das Aplicações Financeiras representa rendimentos pré-fixados.

O saldo líquido de empréstimos totalizou ${formatCurrencySimple(totalLiqEmprestimos4)}.`;

  // --- RE-PROCESS TABLE 5 (Note 3) ---
  const table5 = table5Detail;
  table5.noteTitle = "Nota Explicativa 3: Estoques";
  table5.preCommentary = `Em ${formattedDate}, o COMAER apresentou no subgrupo Estoques um saldo de ${formatCurrencySimple(table5.total.atual)}, representando um ${table5.total.variacao >= 0 ? "crescimento" : "diminuição"} de ${Math.abs(table5.total.variacao).toFixed(2)}% em relação ao exercício anterior.`;
  table5.commentary = `A composição dos estoques reflete as principais rubricas de materiais de consumo e sobressalentes operacionais.`;

  // --- RE-PROCESS TABLE 6 (Note 4) ---
  const dem6 = table6Raw.rows.find(r => r.rubrica === "Demais créditos e valores a longo prazo");
  const demAdj6 = table6Raw.rows.find(r => r.rubrica === "Ajuste para Perdas de demais créditos e valores");
  if (dem6 && demAdj6) {
    dem6.atual += Math.abs(demAdj6.atual);
    dem6.anterior += Math.abs(demAdj6.anterior);
    dem6.variacao = Math.abs(dem6.anterior) > 0.001 ? ((dem6.atual - dem6.anterior) / Math.abs(dem6.anterior)) * 100 : 0;
  }
  const emp6 = table6Raw.rows.find(r => r.rubrica === "Empréstimos e financiamentos concedidos");
  const empAdj6 = table6Raw.rows.find(r => r.rubrica === "Ajuste para Perdas em Emprest. e Finaciament. Concedidos");
  const est6 = table6Raw.rows.find(r => r.rubrica === "Estoques");

  const subEmp6Atual = (emp6?.atual || 0) + (empAdj6?.atual || 0);
  const subEmp6Anterior = (emp6?.anterior || 0) + (empAdj6?.anterior || 0);
  const subDem6Atual = (dem6?.atual || 0) + (demAdj6?.atual || 0);
  const subDem6Anterior = (dem6?.anterior || 0) + (demAdj6?.anterior || 0);

  const totalAtual6 = subEmp6Atual + subDem6Atual;
  const totalAnterior6 = subEmp6Anterior + subDem6Anterior;

  const groups6: any[] = [];
  if (emp6) {
    groups6.push({ 
      value: emp6.atual, 
      rows: [
        emp6, 
        ...(empAdj6 ? [empAdj6] : []), 
        { 
          rubrica: "Subtotal Empréstimos (Líquido)", 
          atual: subEmp6Atual, 
          anterior: subEmp6Anterior, 
          variacao: Math.abs(subEmp6Anterior) > 0.001 ? ((subEmp6Atual - subEmp6Anterior) / Math.abs(subEmp6Anterior)) * 100 : 0, 
          av: 0, 
          isSubtotal: true 
        }
      ] 
    });
  }
  if (dem6) {
    groups6.push({ 
      value: dem6.atual, 
      rows: [
        dem6, 
        ...(demAdj6 ? [demAdj6] : []), 
        { 
          rubrica: "Subtotal Demais Créditos (Líquido)", 
          atual: subDem6Atual, 
          anterior: subDem6Anterior, 
          variacao: Math.abs(subDem6Anterior) > 0.001 ? ((subDem6Atual - subDem6Anterior) / Math.abs(subDem6Anterior)) * 100 : 0, 
          av: 0, 
          isSubtotal: true 
        }
      ] 
    });
  }
  if (est6) groups6.push({ value: est6.atual, rows: [est6] });
  groups6.sort((a,b) => b.value - a.value);

  // Logic for T6 Table Commentary
  const top2Names = groups6.slice(0, 2).map(g => g.rows[0].rubrica).join(" e ");
  const top2Sum = formatCurrencySimple(groups6.slice(0, 2).reduce((acc, g) => acc + g.value, 0));
  const mostRepGroup = groups6[0]?.rows[0].rubrica || "";
  const secondRepGroup = groups6[1]?.rows[0].rubrica || "";
  const demAdjValue = formatCurrencySimple(Math.abs(demAdj6?.atual || 0));

  const table6: TableData = {
    ...table6Raw,
    rows: groups6.flatMap(g => g.rows).map(r => ({
      ...r,
      av: Math.abs(totalAtual6) > 0.001 ? (r.atual / totalAtual6) * 100 : 0
    })),
    total: {
      rubrica: "TOTAL REALIZÁVEL A LONGO PRAZO",
      atual: totalAtual6,
      anterior: totalAnterior6,
      variacao: Math.abs(totalAnterior6) > 0.001 ? ((totalAtual6 - totalAnterior6) / Math.abs(totalAnterior6)) * 100 : 0,
      av: 100
    },
    noteTitle: "Nota Explicativa 4: Ativo Realizável a Longo Prazo",
    preCommentary: `Em ${formattedDate}, o Ativo Realizável a Longo Prazo totalizou ${formatCurrencySimple(totalAtual6)}.`,
    commentary: `As rubricas ${top2Names} representam, aproximadamente, ${top2Sum} do saldo do Ativo Realizável a Longo Prazo. 

A rubrica ${mostRepGroup} trata-se de créditos a receber por desfalques/desvios ou por danos ou perdas. O ajuste para perdas apresentou o montante de ${demAdjValue}. Esse montante decorre da aplicação da mesma metodologia utilizada no cálculo da PECLD dos Créditos por Danos ao Patrimônio Público. Dessa forma, a redução do saldo do Ativo Realizável a Longo Prazo decorre, preponderantemente, do reforço dos ajustes para perdas, não refletindo, necessariamente, diminuição proporcional dos créditos brutos registrados.

A rubrica ${secondRepGroup} refere-se a empréstimos e financiamentos concedidos, bem como créditos do Fundo de Compensação de Variações Salariais (FCVS), operado pela CEF, relativos aos contratos habitacionais encerrados no âmbito do Sistema Financeiro Habitacional (SFH). Todos esses valores são gerenciados pela CFIAe.

A rubrica estoques representa valores que necessitam de confirmação de saldos junto à CFIAe.`
  };

  // --- TABELA 7 --- (Imobilizado)
  const bensMoveisBal = balances.filter(b => b.conta.startsWith("12311") || (b.conta.startsWith("1231") && !b.nomeConta.toLowerCase().includes("deprec")));
  const depMoveisBal = balances.filter(b => b.conta.startsWith("1238101") || (b.nomeConta.toLowerCase().includes("deprec") && b.nomeConta.toLowerCase().includes("móve")));
  const bensImoveisBal = balances.filter(b => b.conta.startsWith("12321") || (b.conta.startsWith("1232") && !b.nomeConta.toLowerCase().includes("deprec")));
  const depImoveisBal = balances.filter(b => b.conta.startsWith("1238102") || (b.nomeConta.toLowerCase().includes("deprec") && b.nomeConta.toLowerCase().includes("imóve")));

  const calcRow = (label: string, filtered: AccountingBalance[]) => {
    const atual = filtered.reduce((acc, b) => acc + b.saldoAtual, 0);
    const anterior = filtered.reduce((acc, b) => acc + b.saldoAnterior, 0);
    const variacao = Math.abs(anterior) > 0.001 ? ((atual - anterior) / Math.abs(anterior)) * 100 : 0;
    return { rubrica: label, atual, anterior, variacao, av: 0 };
  };

  const bmRow = calcRow("Bens móveis", bensMoveisBal);
  const dmRaw = calcRow("Depreciação acumulada de bens móveis", depMoveisBal);
  const dmRow: TableRow = {
    ...dmRaw,
    atual: -Math.abs(dmRaw.atual),
    anterior: -Math.abs(dmRaw.anterior),
    variacao: Math.abs(dmRaw.anterior) > 0.001 ? ((-Math.abs(dmRaw.atual) - (-Math.abs(dmRaw.anterior))) / Math.abs(dmRaw.anterior)) * 100 : 0
  };

  const biRow = calcRow("Bens imóveis", bensImoveisBal);
  const diRaw = calcRow("Depreciação acumulada de bens imóveis", depImoveisBal);
  const diRow: TableRow = {
    ...diRaw,
    atual: -Math.abs(diRaw.atual),
    anterior: -Math.abs(diRaw.anterior),
    variacao: Math.abs(diRaw.anterior) > 0.001 ? ((-Math.abs(diRaw.atual) - (-Math.abs(diRaw.anterior))) / Math.abs(diRaw.anterior)) * 100 : 0
  };

  const subBensMoveis: TableRow = {
    rubrica: "Subtotal Bens Móveis (Líquido)",
    atual: bmRow.atual + dmRow.atual,
    anterior: bmRow.anterior + dmRow.anterior,
    variacao: 0,
    av: 0,
    isSubtotal: true
  };
  subBensMoveis.variacao = Math.abs(subBensMoveis.anterior) > 0.001 ? ((subBensMoveis.atual - subBensMoveis.anterior) / Math.abs(subBensMoveis.anterior)) * 100 : 0;

  const subBensImoveis: TableRow = {
    rubrica: "Subtotal Bens Imóveis (Líquido)",
    atual: biRow.atual + diRow.atual,
    anterior: biRow.anterior + diRow.anterior,
    variacao: 0,
    av: 0,
    isSubtotal: true
  };
  subBensImoveis.variacao = Math.abs(subBensImoveis.anterior) > 0.001 ? ((subBensImoveis.atual - subBensImoveis.anterior) / Math.abs(subBensImoveis.anterior)) * 100 : 0;

  const totalAtual7 = subBensMoveis.atual + subBensImoveis.atual;
  const totalAnterior7 = subBensMoveis.anterior + subBensImoveis.anterior;

  // Calculate AV against totalAtual7
  subBensMoveis.av = Math.abs(totalAtual7) > 0.001 ? (subBensMoveis.atual / totalAtual7) * 100 : 0;
  bmRow.av = Math.abs(totalAtual7) > 0.001 ? (bmRow.atual / totalAtual7) * 100 : 0;
  dmRow.av = Math.abs(totalAtual7) > 0.001 ? (dmRow.atual / totalAtual7) * 100 : 0;
  subBensImoveis.av = Math.abs(totalAtual7) > 0.001 ? (subBensImoveis.atual / totalAtual7) * 100 : 0;
  biRow.av = Math.abs(totalAtual7) > 0.001 ? (biRow.atual / totalAtual7) * 100 : 0;
  diRow.av = Math.abs(totalAtual7) > 0.001 ? (diRow.atual / totalAtual7) * 100 : 0;

  const t7Rows: TableRow[] = [
    subBensMoveis,
    bmRow,
    dmRow,
    subBensImoveis,
    biRow,
    diRow
  ];

  const table7: TableData = {
    title: "Tabela 7: Imobilizado - composição",
    rows: t7Rows,
    total: {
      rubrica: "TOTAL IMOBILIZADO",
      atual: totalAtual7,
      anterior: totalAnterior7,
      variacao: Math.abs(totalAnterior7) > 0.001 ? ((totalAtual7 - totalAnterior7) / Math.abs(totalAnterior7)) * 100 : 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 5: Imobilizado",
    preCommentary: `Ativo Imobilizado registrou ${formatCurrencySimple(totalAtual7)} em ${currentPeriod}. Os ativos são compostos por bens móveis e imóveis, registrados pelo custo de aquisição, construção ou produção, líquidos da depreciação, amortização ou exaustão acumuladas, conforme aplicável. Na tabela a seguir, é apresentada a composição do Subgrupo Imobilizado:`,
    commentary: `**Bens Móveis**\nA composição dos Bens Móveis, com saldo total de ${formatCurrencySimple(bmRow.atual)}, apresentou ${bmRow.variacao >= 0 ? "acréscimo" : "diminuição"} de ${Math.abs(bmRow.variacao).toFixed(2)}% (AH). Os saldos dos bens estão distribuídos conforme detalhamento a seguir:`
  };

  // --- TABELA 8 --- (Bens Móveis - composição)
  const aggregatedBensMoveis = aggregateBalancesByAccount(bensMoveisBal).sort((a,b) => b.saldoAtual - a.saldoAtual);
  const first8 = aggregatedBensMoveis.slice(0, 8);
  const others = aggregatedBensMoveis.slice(8);
  
  const t8Rows: TableRow[] = first8.map(b => {
    let cleanName = b.nomeConta.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim();
    if (cleanName.length > 0) {
      cleanName = cleanName.charAt(0).toUpperCase() + cleanName.slice(1).toLowerCase();
    } else {
      cleanName = b.nomeConta;
    }
    return {
      rubrica: cleanName,
      atual: b.saldoAtual,
      anterior: b.saldoAnterior,
      variacao: Math.abs(b.saldoAnterior) > 0.001 ? ((b.saldoAtual - b.saldoAnterior) / Math.abs(b.saldoAnterior)) * 100 : 0,
      av: 0
    };
  });

  if (others.length > 0) {
    const othersAtual = others.reduce((acc, b) => acc + b.saldoAtual, 0);
    const othersAnterior = others.reduce((acc, b) => acc + b.saldoAnterior, 0);
    t8Rows.push({
      rubrica: "Demais",
      atual: othersAtual,
      anterior: othersAnterior,
      variacao: Math.abs(othersAnterior) > 0.001 ? ((othersAtual - othersAnterior) / Math.abs(othersAnterior)) * 100 : 0,
      av: 0
    });
  }

  t8Rows.push(dmRow); // Depreciação acumulada de bens móveis

  const totalAtual8 = t8Rows.reduce((acc, r) => acc + r.atual, 0);
  const table8: TableData = {
    title: "Tabela 8: Bens Móveis - composição",
    rows: t8Rows.map(r => ({ ...r, av: Math.abs(totalAtual8) > 0.001 ? (r.atual / totalAtual8) * 100 : 0 })),
    total: {
      rubrica: "TOTAL BENS MÓVEIS (LÍQUIDO)",
      atual: totalAtual8,
      anterior: t8Rows.reduce((acc, r) => acc + r.anterior, 0),
      variacao: 0,
      av: 100
    },
    currentPeriod,
    previousPeriod
  };
  table8.total.variacao = Math.abs(table8.total.anterior) > 0.001 ? ((table8.total.atual - table8.total.anterior) / Math.abs(table8.total.anterior)) * 100 : 0;

  // Commentary for T8
  const elaboracao = t8Rows.find(r => r.rubrica.toLowerCase().includes("elaboração"))?.atual || 0;
  const aeronaves = t8Rows.find(r => r.rubrica.toLowerCase().includes("aeronaves"))?.atual || 0;
  const reparo = t8Rows.find(r => r.rubrica.toLowerCase().includes("reparo"))?.atual || 0;
  const sumTop3 = elaboracao + aeronaves + reparo;
  const sumPercent = totalAtual8 !== 0 ? (sumTop3 / totalAtual8) * 100 : 0;

  table8.commentary = `A grande concentração dos Bens Móveis do COMAER está nas rubricas Bens Móveis em Elaboração, Aeronaves e Bens Móveis em Reparo, que juntas representam aproximadamente ${sumPercent.toFixed(2)}% do total.

**Bens Imóveis**
A composição do grupo de Bens Imóveis, com saldo total de ${formatCurrencySimple(biRow.atual)}, apresentou ${biRow.variacao >= 0 ? "acréscimo" : "diminuição"} de ${Math.abs(biRow.variacao).toFixed(2)}% (AH). Os saldos dos bens estão distribuídos conforme detalhamento a seguir:`;

  // --- TABELA 9 --- (Bens Imóveis - composição)
  const aggregatedBensImoveis = aggregateBalancesByAccount(bensImoveisBal).sort((a,b) => b.saldoAtual - a.saldoAtual);
  const first5BI = aggregatedBensImoveis.slice(0, 5);
  const othersBI = aggregatedBensImoveis.slice(5);

  const cleanBI = (name: string) => {
    let clean = name.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim();
    if (clean.length > 0) {
      return clean.charAt(0).toUpperCase() + clean.slice(1).toLowerCase();
    }
    return name;
  };

  const t9Rows: TableRow[] = first5BI.map(b => ({
    rubrica: cleanBI(b.nomeConta),
    atual: b.saldoAtual,
    anterior: b.saldoAnterior,
    variacao: Math.abs(b.saldoAnterior) > 0.001 ? ((b.saldoAtual - b.saldoAnterior) / Math.abs(b.saldoAnterior)) * 100 : 0,
    av: 0
  }));

  if (othersBI.length > 0) {
    const othersAtual = othersBI.reduce((acc, b) => acc + b.saldoAtual, 0);
    const othersAnterior = othersBI.reduce((acc, b) => acc + b.saldoAnterior, 0);
    t9Rows.push({
      rubrica: "Demais",
      atual: othersAtual,
      anterior: othersAnterior,
      variacao: Math.abs(othersAnterior) > 0.001 ? ((othersAtual - othersAnterior) / Math.abs(othersAnterior)) * 100 : 0,
      av: 0
    });
  }

  t9Rows.push(diRow); // Depreciação acumulada de bens imóveis

  const totalAtual9 = t9Rows.reduce((acc, r) => acc + r.atual, 0);
  const table9: TableData = {
    title: "Tabela 9: Bens Imóveis - composição",
    rows: t9Rows.map(r => ({ ...r, av: Math.abs(totalAtual9) > 0.001 ? (r.atual / totalAtual9) * 100 : 0 })),
    total: {
      rubrica: "TOTAL BENS IMÓVEIS (LÍQUIDO)",
      atual: totalAtual9,
      anterior: t9Rows.reduce((acc, r) => acc + r.anterior, 0),
      variacao: 0,
      av: 100
    },
    currentPeriod,
    previousPeriod
  };
  table9.total.variacao = Math.abs(table9.total.anterior) > 0.001 ? ((table9.total.atual - table9.total.anterior) / Math.abs(table9.total.anterior)) * 100 : 0;

  // Commentary for T9
  const top3BI = [...table9.rows].filter(r => r.rubrica !== "TOTAL BENS IMÓVEIS (LÍQUIDO)" && !r.isSubtotal).sort((a,b) => b.av - a.av).slice(0, 3);
  const top3BINames = top3BI.map(r => r.rubrica).join(", ");
  const top3BIAVSum = top3BI.reduce((acc, r) => acc + r.av, 0).toFixed(2);

  const mostRepBI = top3BI[0];
  const secondRepBI = top3BI[1];

  table9.commentary = `A composição do grupo permanece concentrada nos itens ${top3BINames}, que, em conjunto, representam ${top3BIAVSum}% do total.
 
Os ${mostRepBI?.rubrica || 'Bens imóveis'}, principal componente do grupo, apresentaram ${mostRepBI?.variacao >= 0 ? "aumento" : "redução"} de ${Math.abs(mostRepBI?.variacao || 0).toFixed(2)}% em relação ao exercício anterior. Os ${secondRepBI?.rubrica || 'Bens imóveis de uso especial'} registraram ${secondRepBI?.variacao >= 0 ? "aumento" : "redução"} de ${Math.abs(secondRepBI?.variacao || 0).toFixed(2)}%.`;

  // --- TABELA 10 --- (Intangível - composição)
  const intRules = [
    { label: "Software - vida útil definida", matcher: (b: AccountingBalance) => b.conta.startsWith("124110101") || b.nomeConta.toLowerCase().includes("definida") },
    { label: "Software - vida útil indefinida", matcher: (b: AccountingBalance) => b.conta.startsWith("124110201") || b.nomeConta.toLowerCase().includes("indefinida") },
    { label: "Software em desenvolvimento", matcher: (b: AccountingBalance) => b.conta.startsWith("124110202") || b.nomeConta.toLowerCase().includes("desenvolvimento") },
    { label: "Demais", matcher: (b: AccountingBalance) => b.conta.startsWith("124") && !b.conta.startsWith("124110101") && !b.conta.startsWith("124110201") && !b.conta.startsWith("124110202") && !b.conta.startsWith("1248") },
    { label: "Amortização acumulada", matcher: (b: AccountingBalance) => b.conta.startsWith("1248") || b.nomeConta.toLowerCase().includes("amortiza") }
  ];

  const t10RawRows: TableRow[] = intRules.map(rule => {
    const filtered = balances.filter(rule.matcher);
    let atual = filtered.reduce((acc, b) => acc + b.saldoAtual, 0);
    let anterior = filtered.reduce((acc, b) => acc + b.saldoAnterior, 0);
    if (rule.label === "Amortização acumulada") {
      atual = -Math.abs(atual);
      anterior = -Math.abs(anterior);
    }
    const variacao = Math.abs(anterior) > 0.001 ? ((atual - anterior) / Math.abs(anterior)) * 100 : 0;
    return { rubrica: rule.label, atual, anterior, variacao, av: 0 };
  }).filter(r => r.rubrica !== "Demais" || r.atual !== 0 || r.anterior !== 0);

  const totalAtual10 = t10RawRows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior10 = t10RawRows.reduce((acc, r) => acc + r.anterior, 0);

  const t10Rows: TableRow[] = t10RawRows.map(r => ({
    ...r,
    av: Math.abs(totalAtual10) > 0.001 ? (r.atual / totalAtual10) * 100 : 0
  })).sort((a, b) => {
    if (a.rubrica === "Amortização acumulada") return 1;
    if (b.rubrica === "Amortização acumulada") return -1;
    return b.atual - a.atual;
  });

  const table10: TableData = {
    title: "Tabela 10: Bens Intangíveis - composição",
    rows: t10Rows.map(r => ({ ...r, av: Math.abs(totalAtual10) > 0.001 ? (r.atual / totalAtual10) * 100 : 0 })),
    total: {
      rubrica: "TOTAL INTANGÍVEL (LÍQUIDO)",
      atual: totalAtual10,
      anterior: totalAnterior10,
      variacao: Math.abs(totalAnterior10) > 0.001 ? ((totalAtual10 - totalAnterior10) / Math.abs(totalAnterior10)) * 100 : 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 6: Intangível",
    preCommentary: `Em ${formattedDate}, o Órgão apresentou um saldo de ${formatCurrencySimple(totalAtual10)} no Subgrupo Intangível e sua composição é formada por softwares, conforme consta da tabela abaixo:`
  };

  const mostRepIntangible = [...table10.rows].filter(r => r.rubrica !== "Amortização acumulada").sort((a,b) => b.av - a.av)[0];
  let t10Commentary = "";
  if (mostRepIntangible) {
    t10Commentary = `O item ${mostRepIntangible.rubrica} representa a parcela mais significativa do total, com ${mostRepIntangible.av.toFixed(2)}% de representatividade, sendo composto majoritariamente por sistemas em desenvolvimento relacionados ao controle do espaço aéreo nacional, sob a responsabilidade da Comissão de Implantação do Sistema de Controle do Espaço Aéreo – CISCEA.`;
  }

  // Check for variation > 15%
  const highVarItems = table10.rows.filter(r => Math.abs(r.variacao) > 15);
  if (highVarItems.length > 0) {
    const varTexts = highVarItems.map(r => `Em relação à rubrica ${r.rubrica}, a variação de ${r.variacao.toFixed(2)}%, ocorreu em função de...`);
    t10Commentary += `\n\n${varTexts.join("\n")}`;
  }
  table10.commentary = t10Commentary;

  // --- TABELA 11 --- (Obrigações Trabalhistas - composição)
  const laborBalances = balances.filter(b => b.conta.startsWith("211"));
  const cleanLaborName = (name: string) => {
    let clean = name.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim();
    if (clean.length > 0) {
      return clean.charAt(0).toUpperCase() + clean.slice(1).toLowerCase();
    }
    return name;
  };

  const aggregatedLabor = aggregateBalancesByAccount(laborBalances).sort((a, b) => b.saldoAtual - a.saldoAtual);
  const laborFirst3 = aggregatedLabor.slice(0, 3);
  const laborOthers = aggregatedLabor.slice(3);

  const t11Rows: TableRow[] = laborFirst3.map(b => ({
    rubrica: cleanLaborName(b.nomeConta),
    atual: b.saldoAtual,
    anterior: b.saldoAnterior,
    variacao: Math.abs(b.saldoAnterior) > 0.001 ? ((b.saldoAtual - b.saldoAnterior) / Math.abs(b.saldoAnterior)) * 100 : 0,
    av: 0
  }));

  if (laborOthers.length > 0) {
    const othersAtual = laborOthers.reduce((acc, b) => acc + b.saldoAtual, 0);
    const othersAnterior = laborOthers.reduce((acc, b) => acc + b.saldoAnterior, 0);
    t11Rows.push({
      rubrica: "Demais",
      atual: othersAtual,
      anterior: othersAnterior,
      variacao: Math.abs(othersAnterior) > 0.001 ? ((othersAtual - othersAnterior) / Math.abs(othersAnterior)) * 100 : 0,
      av: 0
    });
  }

  const totalAtual11 = t11Rows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior11 = t11Rows.reduce((acc, r) => acc + r.anterior, 0);

  const table11: TableData = {
    title: "Tabela 11: Obrigações Trabalhistas, Previdenciárias e Assistenciais - composição",
    rows: t11Rows.map(r => ({ ...r, av: Math.abs(totalAtual11) > 0.001 ? (r.atual / totalAtual11) * 100 : 0 })),
    total: {
      rubrica: "TOTAL OBRIGAÇÕES",
      atual: totalAtual11,
      anterior: totalAnterior11,
      variacao: Math.abs(totalAnterior11) > 0.001 ? ((totalAtual11 - totalAnterior11) / Math.abs(totalAnterior11)) * 100 : 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 7: Obrigações Trabalhistas, Previdenciárias e Assistenciais",
    preCommentary: `O saldo do item consiste em obrigações de curto prazo, compreendendo valores a pagar referentes a salários, remunerações, benefícios, décimo terceiro salário, férias e benefícios assistenciais. A composição detalhada das obrigações está apresentada na tabela a seguir:`
  };

  const mostRepLabor = [...table11.rows].sort((a,b) => b.av - a.av)[0];
  let t11Commentary = "";
  if (mostRepLabor) {
    t11Commentary = `O componente mais representativo do total é o de ${mostRepLabor.rubrica} que corresponde a aproximadamente ${mostRepLabor.av.toFixed(2)}%. Esse valor reflete .... Destaca-se que o passivo de décimo terceiro salário é registrado por competência, em atendimento aos procedimentos contábeis estabelecidos pelo Manual SIAFI, Macrofunção 02.11.42 – Folha de Pagamento, da Secretaria do Tesouro Nacional (STN), assegurando a adequada mensuração e reconhecimento das obrigações trabalhistas no momento de sua ocorrência, ainda que o desembolso ocorra futuramente.`;
  }

  const highVarLabor = table11.rows.filter(r => Math.abs(r.variacao) > 15);
  if (highVarLabor.length > 0) {
    const varTexts = highVarLabor.map(r => `Em relação à rubrica ${r.rubrica}, a variação de ${r.variacao.toFixed(2)}%, ocorreu em função de...`);
    t11Commentary += `\n\n${varTexts.join("\n")}`;
  }
  table11.commentary = t11Commentary;

  // --- TABELA 12 --- (Empréstimos e Financiamentos)
  const loanBalances = balances.filter(b => b.conta.startsWith("212") || b.conta.startsWith("222"));
  
  const cleanLoanName = (code: string, name: string) => {
    if (code === "212410201") return "Financiamentos do ativo permanente - curto prazo";
    if (code === "222410101") return "Financiamentos do ativo permanente - longo prazo";
    return name.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim();
  };

  const loanMap = new Map<string, { rubrica: string; atual: number; anterior: number }>();
  for (const b of loanBalances) {
    const rubrica = cleanLoanName(b.conta, b.nomeConta);
    const existing = loanMap.get(rubrica) || { rubrica, atual: 0, anterior: 0 };
    existing.atual += b.saldoAtual;
    existing.anterior += b.saldoAnterior;
    loanMap.set(rubrica, existing);
  }

  const t12Rows: TableRow[] = Array.from(loanMap.values()).map(r => ({
    rubrica: r.rubrica,
    atual: r.atual,
    anterior: r.anterior,
    variacao: Math.abs(r.anterior) > 0.001 ? ((r.atual - r.anterior) / Math.abs(r.anterior)) * 100 : 0,
    av: 0
  })).sort((a, b) => b.atual - a.atual);

  const totalAtual12 = t12Rows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior12 = t12Rows.reduce((acc, r) => acc + r.anterior, 0);

  const table12: TableData = {
    title: "Tabela 12: Empréstimos e Financiamentos (CP e LP) - composição",
    rows: t12Rows.map(r => ({ ...r, av: Math.abs(totalAtual12) > 0.001 ? (r.atual / totalAtual12) * 100 : 0 })),
    total: {
      rubrica: "TOTAL EMPRÉSTIMOS E FINANCIAMENTOS",
      atual: totalAtual12,
      anterior: totalAnterior12,
      variacao: Math.abs(totalAnterior12) > 0.001 ? ((totalAtual12 - totalAnterior12) / Math.abs(totalAnterior12)) * 100 : 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 8: Empréstimos e Financiamentos",
    preCommentary: `O COMAER apresentou saldo de ${formatCurrencySimple(totalAtual12)} referente a financiamentos de curto e longo prazo, concentrados integralmente no Órgão Comando da Aeronáutica. O item é composto exclusivamente por financiamentos externos, contratados para apoio aos projetos estratégicos da Força Aérea Brasileira.`
  };

  let t12Commentary = `Os contratos de financiamento foram realizados com as seguintes instituições:

a) Société Générale (EX9863600), por meio de sindicato de bancos, no âmbito do Projeto H‑X BR, destinado à aquisição de 50 helicópteros de médio porte e suporte logístico; e

b) Swedish Export Credit Corporation – AB SEK (EX9867800), agência de crédito sueca, no âmbito do Projeto F‑X2, voltado à aquisição de 28 aeronaves de caça monoposto e 8 aeronaves biposto.

Os valores são atualizados por variação cambial e com base nas orientações da Secretaria do Tesouro Nacional (STN), por meio de informações fornecidas mensalmente pela Coordenação-Geral de Controle da Dívida Pública (CODIV/GECOD), com base nos dados do Sistema Integrado da Dívida Pública – SID, conforme previsto na Macrofunção 02.03.10 – Obrigações Internas e Externas, do Manual SIAFI.`;

  const highVarLoans = table12.rows.filter(r => Math.abs(r.variacao) > 15);
  if (highVarLoans.length > 0) {
    const varTexts = highVarLoans.map(r => `Em relação à conta ${r.rubrica}, a variação de ${r.variacao.toFixed(2)}%, ocorreu em função de...`);
    t12Commentary += `\n\n${varTexts.join("\n\n")}`;
  }
  table12.commentary = t12Commentary;
  
  // --- TABELA 14 --- (Demais Obrigações a Curto Prazo)
  const otherLiabilityBalances = balances.filter(b => b.conta.startsWith("218"));
  
  const cleanOtherLiabilityName = (name: string) => {
    let clean = name.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim().toLowerCase();
    if (clean.length > 0) {
      return clean.charAt(0).toUpperCase() + clean.slice(1);
    }
    return name;
  };

  const aggregatedOtherLia = aggregateBalancesByAccount(otherLiabilityBalances).sort((a, b) => b.saldoAtual - a.saldoAtual);
  const otherLiaFirst3 = aggregatedOtherLia.slice(0, 3);
  const otherLiaRemainder = aggregatedOtherLia.slice(3);

  const t14Rows: TableRow[] = otherLiaFirst3.map(b => ({
    rubrica: cleanOtherLiabilityName(b.nomeConta),
    atual: b.saldoAtual,
    anterior: b.saldoAnterior,
    variacao: Math.abs(b.saldoAnterior) > 0.001 ? ((b.saldoAtual - b.saldoAnterior) / Math.abs(b.saldoAnterior)) * 100 : 0,
    av: 0
  }));

  if (otherLiaRemainder.length > 0) {
    const othersAtual = otherLiaRemainder.reduce((acc, b) => acc + b.saldoAtual, 0);
    const othersAnterior = otherLiaRemainder.reduce((acc, b) => acc + b.saldoAnterior, 0);
    t14Rows.push({
      rubrica: "Demais",
      atual: othersAtual,
      anterior: othersAnterior,
      variacao: Math.abs(othersAnterior) > 0.001 ? ((othersAtual - othersAnterior) / Math.abs(othersAnterior)) * 100 : 0,
      av: 0
    });
  }

  const totalAtual14 = t14Rows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior14 = t14Rows.reduce((acc, r) => acc + r.anterior, 0);

  const table14: TableData = {
    title: "Tabela 14: Demais Obrigações a Curto Prazo – composição",
    rows: t14Rows.map(r => ({ ...r, av: Math.abs(totalAtual14) > 0.001 ? (r.atual / totalAtual14) * 100 : 0 })),
    total: {
      rubrica: "TOTAL DEMAIS OBRIGAÇÕES",
      atual: totalAtual14,
      anterior: totalAnterior14,
      variacao: Math.abs(totalAnterior14) > 0.001 ? ((totalAtual14 - totalAnterior14) / Math.abs(totalAnterior14)) * 100 : 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 10: Demais Obrigações a Curto Prazo",
    preCommentary: `Em ${currentPeriod}, o COMAER apresentou um saldo de ${formatCurrencySimple(totalAtual14)} relacionado a demais obrigações de curto prazo. A composição do item e suas variações são apresentadas a seguir:`
  };

  const sortedByAV14 = [...table14.rows].sort((a, b) => b.av - a.av);
  const mostRep14 = sortedByAV14[0];
  const secondRep14 = sortedByAV14[1];

  let t14Commentary = "";
  if (mostRep14) {
    t14Commentary = `A rubrica “${mostRep14.rubrica}”, que representa ${mostRep14.av.toFixed(2)}% do total, permanece como o principal componente do grupo. O saldo decorre da sistemática de execução descentralizada adotada a partir de 2019, conforme orientações da STN. Nessa dinâmica, a UG descentralizadora registra o ativo financeiro enquanto a UG recebedora contabiliza a obrigação até a devida comprovação da aplicação dos recursos.`;
  }
  if (secondRep14) {
    t14Commentary += `\n\nO item “${secondRep14.rubrica}”, responsável por ${secondRep14.av.toFixed(2)}% do total …….`;
  }

  const highVar14 = table14.rows.filter(r => Math.abs(r.variacao) > 15);
  if (highVar14.length > 0) {
    const varTexts = highVar14.map(r => `Em relação à conta ${r.rubrica}, a variação de ${r.variacao.toFixed(2)}%, ocorreu em função de...`);
    t14Commentary += `\n\n${varTexts.join("\n\n")}`;
  }
  table14.commentary = t14Commentary;

  // --- TABELA 13 --- (Fornecedores e Contas a Pagar – Por Contratado)
  const supplierBalances = balances.filter(b => b.conta.startsWith("213") || b.nomeConta.toLowerCase().includes("fornecedor"));
  const totalAtual13 = supplierBalances.reduce((acc, b) => acc + b.saldoAtual, 0);
  const totalAnterior13 = supplierBalances.reduce((acc, b) => acc + b.saldoAnterior, 0);

  const supplierMap = new Map<string, { atual: number; anterior: number }>();
  for (const b of supplierBalances) {
    const rawName = b.entidadeCCor || b.nomeConta || "Outros Fornecedores";
    let cleanName = rawName.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim();
    if (!cleanName) cleanName = "Outros Fornecedores";
    const cur = supplierMap.get(cleanName) || { atual: 0, anterior: 0 };
    cur.atual += b.saldoAtual;
    cur.anterior += b.saldoAnterior;
    supplierMap.set(cleanName, cur);
  }

  const supplierGroupedList: TableRow[] = [];
  for (const [rubrica, val] of supplierMap.entries()) {
    const av = Math.abs(totalAtual13) > 0.001 ? (val.atual / totalAtual13) * 100 : 0;
    supplierGroupedList.push({
      rubrica,
      atual: val.atual,
      anterior: val.anterior,
      variacao: 0,
      av
    });
  }

  const supplierMajor: TableRow[] = [];
  let supplierDemaisAtual = 0;
  let supplierDemaisAnterior = 0;
  let supplierDemaisAV = 0;

  supplierGroupedList.forEach(item => {
    if (item.av >= 5) {
      supplierMajor.push(item);
    } else {
      supplierDemaisAtual += item.atual;
      supplierDemaisAnterior += item.anterior;
      supplierDemaisAV += item.av;
    }
  });

  supplierMajor.sort((a, b) => b.atual - a.atual);

  const finalT13Rows: TableRow[] = [...supplierMajor];
  if (supplierDemaisAtual !== 0 || supplierDemaisAV !== 0 || supplierMajor.length === 0) {
    finalT13Rows.push({
      rubrica: "Demais",
      atual: supplierDemaisAtual,
      anterior: supplierDemaisAnterior,
      variacao: 0,
      av: supplierDemaisAV
    });
  }

  const table13: TableData = {
    title: "Tabela 13: Fornecedores e Contas a Pagar – Por Contratado",
    rows: finalT13Rows,
    total: {
      rubrica: "TOTAL FORNECEDORES",
      atual: totalAtual13,
      anterior: totalAnterior13,
      variacao: 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    onlyAV: true,
    noteTitle: "Nota Explicativa 9: Fornecedores a Pagar",
    preCommentary: `Em ${currentPeriod}, o COMAER apresentou um saldo de ${formatCurrencySimple(totalAtual13)} relacionado a Fornecedores e Contas a Pagar no curto prazo conforme detalhamento abaixo:`,
    commentary: supplierMajor.length > 0 
      ? `Entre os principais credores e fornecedores do período, destacam-se: ${supplierMajor.slice(0, 3).map(s => `${s.rubrica} (${s.av.toFixed(2)}%)`).join(', ')}. As contas e fornecedores com participação inferior a 5% foram consolidados na rubrica "Demais".`
      : undefined
  };
  
  // --- TABELA 15 --- (Provisões Atuariais)
  const actuaryCodes = ["217910600", "217910700", "217910800", "217910900", "227910600", "227910700", "227910800", "227910900"];
  const actuaryBalances = balances.filter(b => actuaryCodes.includes(b.conta));

  const getActuaryLabel = (code: string) => {
    const map: Record<string, string> = {
      "217910600": "Pensões Militares",
      "217910700": "Militares Inativos",
      "217910800": "Pensões Especiais",
      "217910900": "Anistiados",
      "227910600": "Pensões Militares",
      "227910700": "Militares Inativos",
      "227910800": "Pensões Especiais",
      "227910900": "Anistiados"
    };
    return map[code] || code;
  };

  const cpActuary = actuaryBalances.filter(b => b.conta.startsWith("21"));
  const lpActuary = actuaryBalances.filter(b => b.conta.startsWith("22"));

  const cpActuaryMap = new Map<string, { rubrica: string; atual: number; anterior: number }>();
  for (const b of cpActuary) {
    const rubrica = getActuaryLabel(b.conta);
    const existing = cpActuaryMap.get(rubrica) || { rubrica, atual: 0, anterior: 0 };
    existing.atual += b.saldoAtual;
    existing.anterior += b.saldoAnterior;
    cpActuaryMap.set(rubrica, existing);
  }

  const lpActuaryMap = new Map<string, { rubrica: string; atual: number; anterior: number }>();
  for (const b of lpActuary) {
    const rubrica = getActuaryLabel(b.conta);
    const existing = lpActuaryMap.get(rubrica) || { rubrica, atual: 0, anterior: 0 };
    existing.atual += b.saldoAtual;
    existing.anterior += b.saldoAnterior;
    lpActuaryMap.set(rubrica, existing);
  }

  const cpRows: TableRow[] = Array.from(cpActuaryMap.values()).map(r => ({
    rubrica: r.rubrica,
    atual: r.atual,
    anterior: r.anterior,
    variacao: Math.abs(r.anterior) > 0.001 ? ((r.atual - r.anterior) / Math.abs(r.anterior)) * 100 : 0,
    av: 0
  })).sort((a, b) => b.atual - a.atual);

  const lpRows: TableRow[] = Array.from(lpActuaryMap.values()).map(r => ({
    rubrica: r.rubrica,
    atual: r.atual,
    anterior: r.anterior,
    variacao: Math.abs(r.anterior) > 0.001 ? ((r.atual - r.anterior) / Math.abs(r.anterior)) * 100 : 0,
    av: 0
  })).sort((a, b) => b.atual - a.atual);

  const subTotalCPAtual = cpRows.reduce((acc, r) => acc + r.atual, 0);
  const subTotalCPAnterior = cpRows.reduce((acc, r) => acc + r.anterior, 0);
  const subTotalLPAtual = lpRows.reduce((acc, r) => acc + r.atual, 0);
  const subTotalLPAnterior = lpRows.reduce((acc, r) => acc + r.anterior, 0);

  const totalAtual15 = subTotalCPAtual + subTotalLPAtual;
  const totalAnterior15 = subTotalCPAnterior + subTotalLPAnterior;

  const t15Rows: TableRow[] = [
    ...cpRows.map(r => ({
      ...r,
      av: Math.abs(totalAtual15) > 0.001 ? (r.atual / totalAtual15) * 100 : 0
    })),
    { 
      rubrica: "Curto Prazo (Subtotal)", 
      atual: subTotalCPAtual, 
      anterior: subTotalCPAnterior, 
      variacao: Math.abs(subTotalCPAnterior) > 0.001 ? ((subTotalCPAtual - subTotalCPAnterior) / Math.abs(subTotalCPAnterior)) * 100 : 0, 
      av: Math.abs(totalAtual15) > 0.001 ? (subTotalCPAtual / totalAtual15) * 100 : 0, 
      isSubtotal: true 
    },
    ...lpRows.map(r => ({
      ...r,
      av: Math.abs(totalAtual15) > 0.001 ? (r.atual / totalAtual15) * 100 : 0
    })),
    { 
      rubrica: "Longo Prazo (Subtotal)", 
      atual: subTotalLPAtual, 
      anterior: subTotalLPAnterior, 
      variacao: Math.abs(subTotalLPAnterior) > 0.001 ? ((subTotalLPAtual - subTotalLPAnterior) / Math.abs(subTotalLPAnterior)) * 100 : 0, 
      av: Math.abs(totalAtual15) > 0.001 ? (subTotalLPAtual / totalAtual15) * 100 : 0, 
      isSubtotal: true 
    }
  ];

  const table15: TableData = {
    title: "Tabela 15: Provisões de Benefícios com Militares – composição",
    rows: t15Rows,
    total: {
      rubrica: "TOTAL GERAL",
      atual: totalAtual15,
      anterior: totalAnterior15,
      variacao: Math.abs(totalAnterior15) > 0.001 ? ((totalAtual15 - totalAnterior15) / Math.abs(totalAnterior15)) * 100 : 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 11: Provisões Atuariais",
    preCommentary: `As Provisões Atuariais do Comando da Aeronáutica compreendem, substancialmente, os registros relacionados às Pensões Militares e aos Benefícios com Militares Veteranos, reconhecidos em conformidade com as disposições da NBC TSP 15 – Benefícios a Empregados e demais normativos aplicáveis ao setor público.
O processo de contabilização das provisões atuariais observa os procedimentos estabelecidos na Nota Técnica nº 8/CGCONT/DEORF/SEORI/SG/MD/2025, do Ministério da Defesa, a qual padroniza os registros contábeis relativos aos benefícios de militares veteranos e pensionistas no âmbito das Forças Armadas.
Os valores registrados refletem estimativas atuariais elaboradas com base em premissas biométricas, demográficas, financeiras e econômicas, considerando metodologia de cálculo atuarial apropriada para mensuração das obrigações de longo prazo assumidas pela Administração Pública.
No primeiro trimestre de 2026, não houve alterações relevantes nas premissas atuariais utilizadas em relação ao encerramento do exercício de 2025 que impactassem significativamente os saldos das provisões atuariais do COMAER.
A variação observada nas provisões de curto prazo decorre, principalmente, da padronização promovida pela referida Nota Técnica do Ministério da Defesa, a qual orienta que, para fins de pagamento mensal dos benefícios, os valores sejam reclassificados da conta contábil de provisões para a conta de salários e remunerações, conforme a exigibilidade dos dispêndios no curto prazo.
A composição das provisões atuariais encontra-se demonstrada na Tabela 15`,
    commentary: `As informações complementares referentes às metodologias atuariais, premissas adotadas, critérios de mensuração, estudos técnicos, memórias de cálculo e demais detalhamentos das provisões atuariais encontram-se apresentadas de forma completa nas Demonstrações Contábeis do exercício encerrado em 31 de dezembro de 2025.`
  };

  // Re-calculate Table 0 summary with all finalized groups
  const ativoItems = [
    { ...table1.total, rubrica: "Caixa e Equivalentes de Caixa", nota: "1", av: 0 },
    { ...table4.total, rubrica: "Créditos a Curto Prazo", nota: "2", av: 0 },
    { ...table5.total, rubrica: "Estoques", nota: "3", av: 0 },
    { ...table6.total, rubrica: "Ativo Realizável a Longo Prazo", nota: "4", av: 0 },
    { ...table7.total, rubrica: "Imobilizado", nota: "5", av: 0 },
    { ...table10.total, rubrica: "Intangível", nota: "6", av: 0 }
  ];

  const passivoItems = [
    { ...table11.total, rubrica: "Obrigações Trabalhistas", nota: "7", av: 0 },
    { ...table12.total, rubrica: "Empréstimos e Financiamentos", nota: "8", av: 0 },
    { ...table13.total, rubrica: "Fornecedores a Pagar", nota: "9", av: 0 },
    { ...table14.total, rubrica: "Demais Obrigações a Curto Prazo", nota: "10", av: 0 },
    { ...table15.total, rubrica: "Provisões Atuariais", nota: "11", av: 0 }
  ];

  const totalAtivoFinalAtual = ativoItems.reduce((acc, r) => acc + r.atual, 0);
  const totalAtivoFinalAnterior = ativoItems.reduce((acc, r) => acc + r.anterior, 0);
  const totalPassivoFinalAtual = passivoItems.reduce((acc, r) => acc + r.atual, 0);
  const totalPassivoFinalAnterior = passivoItems.reduce((acc, r) => acc + r.anterior, 0);

  const totalGlobalAtual = totalAtivoFinalAtual + totalPassivoFinalAtual;

  table0.rows = [
    ...ativoItems.map(r => ({ ...r, av: Math.abs(totalAtivoFinalAtual) > 0.001 ? (r.atual / totalAtivoFinalAtual) * 100 : 0 })),
    { 
      rubrica: "TOTAL DO ATIVO", 
      atual: totalAtivoFinalAtual, 
      anterior: totalAtivoFinalAnterior, 
      variacao: Math.abs(totalAtivoFinalAnterior) > 0.001 ? ((totalAtivoFinalAtual - totalAtivoFinalAnterior) / Math.abs(totalAtivoFinalAnterior)) * 100 : 0, 
      av: 100, 
      isSubtotal: true 
    },
    ...passivoItems.map(r => ({ ...r, av: Math.abs(totalPassivoFinalAtual) > 0.001 ? (r.atual / totalPassivoFinalAtual) * 100 : 0 })),
    { 
      rubrica: "TOTAL DO PASSIVO", 
      atual: totalPassivoFinalAtual, 
      anterior: totalPassivoFinalAnterior, 
      variacao: Math.abs(totalPassivoFinalAnterior) > 0.001 ? ((totalPassivoFinalAtual - totalPassivoFinalAnterior) / Math.abs(totalPassivoFinalAnterior)) * 100 : 0, 
      av: 100, 
      isSubtotal: true 
    }
  ];

  table0.total = {
    rubrica: "TOTAL GERAL (ATIVO + PASSIVO)",
    atual: totalGlobalAtual,
    anterior: totalAtivoFinalAnterior + totalPassivoFinalAnterior,
    variacao: 0,
    av: 100
  };
  table0.total.variacao = Math.abs(table0.total.anterior) > 0.001 ? ((table0.total.atual - table0.total.anterior) / Math.abs(table0.total.anterior)) * 100 : 0;

  // --- TABELA 17 --- (Demais Reservas - Composição)
  const t17Rules = [
    { 
      label: "Reavaliação de Bens Imóveis", 
      matcher: (b: AccountingBalance) => b.conta === "236110100" || (b.conta.startsWith("23611") && b.nomeConta.toLowerCase().includes("imóve"))
    },
    { 
      label: "Reavaliação de Bens Móveis", 
      matcher: (b: AccountingBalance) => b.conta === "236110200" || (b.conta === "236110300" && !b.nomeConta.toLowerCase().includes("intang")) || (b.conta.startsWith("23611") && b.nomeConta.toLowerCase().includes("móve") && !b.nomeConta.toLowerCase().includes("imóve"))
    },
    { 
      label: "Reavaliação de Bens Intangíveis", 
      matcher: (b: AccountingBalance) => b.conta === "236110400" || (b.conta === "236110300" && b.nomeConta.toLowerCase().includes("intang")) || (b.conta.startsWith("23611") && b.nomeConta.toLowerCase().includes("intang"))
    }
  ];

  const t17Rows: TableRow[] = t17Rules.map(rule => {
    const filtered = balances.filter(rule.matcher);
    const atual = filtered.reduce((acc, b) => acc + b.saldoAtual, 0);
    const anterior = filtered.reduce((acc, b) => acc + b.saldoAnterior, 0);
    return {
      rubrica: rule.label,
      atual,
      anterior,
      variacao: 0,
      av: 0
    };
  });

  const totalAtual17 = t17Rows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior17 = t17Rows.reduce((acc, r) => acc + r.anterior, 0);

  const table17: TableData = {
    title: "Tabela 17: Demais Reservas – Composição",
    rows: t17Rows,
    total: {
      rubrica: "TOTAL GERAL",
      atual: totalAtual17,
      anterior: totalAnterior17,
      variacao: 0,
      av: 0
    },
    currentPeriod,
    previousPeriod,
    hideAV: true,
    hideAH: true,
    noteTitle: "Nota Explicativa 14: Demais Reservas",
    preCommentary: "Demais Reservas são compostas por saldos de reavaliação de bens do ativo imobilizado e intangível, com destaque para a reserva de reavaliação de bens imóveis, que representa a maior parte do montante registrado na rubrica. A composição detalhada encontra-se a seguir:"
  };

  // --- TABELA 18 --- (Obrigações Contratuais - Composição)
  const t18Rules = [
    { label: "Fornecimento de Bens", matcher: (b: AccountingBalance) => b.conta.startsWith("812310401") || b.conta.startsWith("812311101") || b.nomeConta.toLowerCase().includes("fornecimento de bens") },
    { label: "Serviços", matcher: (b: AccountingBalance) => b.conta.startsWith("812310201") || b.conta.startsWith("812310901") || b.nomeConta.toLowerCase().includes("serviço") },
    { label: "Demais", matcher: (b: AccountingBalance) => b.conta.startsWith("81231") && !b.conta.startsWith("812310401") && !b.conta.startsWith("812311101") && !b.conta.startsWith("812310201") && !b.conta.startsWith("812310901") }
  ];

  const t18RawRows = t18Rules.map(rule => {
    const filtered = balances.filter(rule.matcher);
    const atual = filtered.reduce((acc, b) => acc + b.saldoAtual, 0);
    const anterior = filtered.reduce((acc, b) => acc + b.saldoAnterior, 0);
    const variacao = Math.abs(anterior) > 0.001 ? ((atual - anterior) / Math.abs(anterior)) * 100 : 0;
    return { rubrica: rule.label, atual, anterior, variacao, av: 0 };
  });

  const totalAtual18 = t18RawRows.reduce((acc, r) => acc + r.atual, 0);
  const totalAnterior18 = t18RawRows.reduce((acc, r) => acc + r.anterior, 0);
  const totalVariacao18 = Math.abs(totalAnterior18) > 0.001 ? ((totalAtual18 - totalAnterior18) / Math.abs(totalAnterior18)) * 100 : 0;

  const t18Rows = t18RawRows.map(r => ({
    ...r,
    av: Math.abs(totalAtual18) > 0.001 ? (r.atual / totalAtual18) * 100 : 0
  })).sort((a, b) => b.atual - a.atual);

  const table18: TableData = {
    title: "Tabela 18: Obrigações Contratuais – Composição",
    rows: t18Rows,
    total: {
      rubrica: "TOTAL GERAL",
      atual: totalAtual18,
      anterior: totalAnterior18,
      variacao: totalVariacao18,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    noteTitle: "Nota Explicativa 15: Obrigações Contratuais",
    preCommentary: `Em ${currentPeriod}, o COMAER apresentou saldo de ${formatCurrencySimple(totalAtual18)} em Obrigações Contratuais a executar, representando compromissos firmados por meio de contratos administrativos ainda vigentes. A composição dessas obrigações, segregada por natureza do contrato, está apresentada na tabela seguir:`,
    commentary: `O item Fornecimento de Bens corresponde, majoritariamente, às apropriações de contratos em execução pelas Unidades Gestoras do COMAER, sobretudo àquelas situadas no exterior. Entre esses contratos, destacam-se os firmados com empresas estratégicas para a Defesa Nacional, como SAAB AB, EMBRAER S.A. e EUROCOPTER, no âmbito de projetos como o PROJETO GRIPEN.
 
O controle e a execução desses contratos seguem os normativos da Administração Pública Federal, com acompanhamento orçamentário e financeiro por meio do SIAFI e dos sistemas corporativos do COMAER, garantindo a transparência e o cumprimento dos compromissos institucionais com a Defesa Nacional.
 
Na tabela a seguir, apresentam-se os contratados significativos e o saldo a executar em ${currentPeriod}:`
  };

  // --- TABELA 19 --- (Obrigações Contratuais – Por Contratado)
  const t19AccountCodes = [
    "812310101","812310201","812310301","812310401","812310501",
    "812310601","812310701","812310801","812310901","812311001",
    "812311101","812311201","812311301","812311401"
  ];

  const t19Balances = balances.filter(b => t19AccountCodes.some(c => b.conta === c || b.conta.startsWith(c)));
  const totalAtual19 = t19Balances.reduce((acc, b) => acc + b.saldoAtual, 0);

  const t19Map = new Map<string, { atual: number; anterior: number }>();
  for (const b of t19Balances) {
    const rawName = b.entidadeCCor || b.nomeConta || "Outros Contratados";
    let cleanName = rawName.replace(/^[^a-zA-Z0-9À-ÿ]+/, '').trim();
    if (!cleanName) cleanName = "Outros Contratados";
    const cur = t19Map.get(cleanName) || { atual: 0, anterior: 0 };
    cur.atual += b.saldoAtual;
    cur.anterior += b.saldoAnterior;
    t19Map.set(cleanName, cur);
  }

  const t19GroupedList: TableRow[] = [];
  for (const [rubrica, val] of t19Map.entries()) {
    const av = Math.abs(totalAtual19) > 0.001 ? (val.atual / totalAtual19) * 100 : 0;
    t19GroupedList.push({
      rubrica,
      atual: val.atual,
      anterior: val.anterior,
      variacao: 0,
      av
    });
  }

  const t19Major: TableRow[] = [];
  let t19DemaisAtual = 0;
  let t19DemaisAnterior = 0;
  let t19DemaisAV = 0;

  t19GroupedList.forEach(item => {
    if (item.av >= 5) {
      t19Major.push(item);
    } else {
      t19DemaisAtual += item.atual;
      t19DemaisAnterior += item.anterior;
      t19DemaisAV += item.av;
    }
  });

  t19Major.sort((a, b) => b.atual - a.atual);

  const finalT19Rows: TableRow[] = [...t19Major];
  if (t19DemaisAtual !== 0 || t19DemaisAV !== 0 || t19Major.length === 0) {
    finalT19Rows.push({
      rubrica: "Demais",
      atual: t19DemaisAtual,
      anterior: t19DemaisAnterior,
      variacao: 0,
      av: t19DemaisAV
    });
  }

  const table19: TableData = {
    title: "Tabela 19: Obrigações Contratuais – Por Contratado",
    rows: finalT19Rows,
    total: {
      rubrica: "TOTAL GERAL",
      atual: totalAtual19,
      anterior: 0,
      variacao: 0,
      av: 100
    },
    currentPeriod,
    previousPeriod,
    onlyAV: true,
    noteTitle: "Nota Explicativa 15: Obrigações Contratuais – Por Contratado",
    preCommentary: `Apresenta-se a seguir a composição por contratado (Entidade CCor) das obrigações contratuais a executar em ${currentPeriod}:`
  };

  const resolvedCurrentDate = result.currentDate || formatDateDDMMYYYY(currentPeriod);
  const resolvedPreviousDate = result.previousDate || formatDateDDMMYYYY(previousPeriod);

  const allTables: TableData[] = [
    table0, table1, table2, table3, table4, table5, table6, table7, table8, table9, table10,
    table11, table12, table13, table14, table15, table17, table18, table19
  ];
  return allTables.map(t => ({
    ...t,
    currentDate: t.currentDate || resolvedCurrentDate,
    previousDate: t.previousDate || resolvedPreviousDate,
  }));
};

