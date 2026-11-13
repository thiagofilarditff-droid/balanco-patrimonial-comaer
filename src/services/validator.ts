import { BalancoData, StatementItem, ValidationRuleResult, ValidationStatus } from '../types/balanco';
import { roundCommercial } from '../utils/numberFormat';

export interface ValidationSummary {
  rules: ValidationRuleResult[];
  conformeCount: number;
  avisoCount: number;
  erroCount: number;
  pendenteCount: number;
  canExport: boolean;
  blockingReasons: string[];
}

/**
 * Calculates conservative tolerance for n items rounded with precision p (p = 0.01 thousand = 10 reais)
 * limit = (n + 1) * p / 2
 */
function calcTolerance(n: number, p = 0.01): number {
  return Number((((n + 1) * p) / 2).toFixed(4));
}

function getItemVal(item: StatementItem | undefined, isCurrent: boolean): number {
  if (!item) return 0;
  const val = isCurrent ? item.valCurrent : item.valPrevious;
  return val !== null && val !== undefined ? val : 0;
}

function getItemRef(item: StatementItem | undefined, isCurrent: boolean): string {
  if (!item) return '-';
  const ref = isCurrent ? item.cellRefCurrent : item.cellRefPrevious;
  return ref || '-';
}

export function validateBalanco(data: BalancoData): ValidationSummary {
  const rules: ValidationRuleResult[] = [];
  const blockingReasons: string[] = [];

  const { metadata, ativo, passivoPL, financeiroPermanente, compensacoes, superavitDeficit } = data;
  const currYear = metadata.colCurrent || '2026';
  const prevYear = metadata.colPrevious || '2025';

  const findAtivo = (id: string) => ativo.find(x => x.id === id);
  const findPassivo = (id: string) => passivoPL.find(x => x.id === id);
  const findFP = (id: string) => financeiroPermanente.find(x => x.id === id);
  const findComp = (id: string) => compensacoes.find(x => x.id === id);
  const findSD = (id: string) => superavitDeficit.find(x => x.id === id);

  // 1. CHECAGEM DE RUBRICAS NÃO MAPEADAS
  const allItems = [...ativo, ...passivoPL, ...financeiroPermanente, ...compensacoes, ...superavitDeficit];
  const unmappedItems = allItems.filter(item => item.isUnmapped);
  if (unmappedItems.length > 0) {
    unmappedItems.forEach(item => {
      rules.push({
        id: `unmapped_${item.id}`,
        quadro: item.compositeKey ? item.compositeKey.split(' > ')[0] : 'Estrutura',
        exercicio: currYear,
        regra: `Localização de Rubrica: ${item.label}`,
        celulasOrigem: item.cellRefCurrent || 'Não localizada',
        valorInformado: 0,
        valorCalculado: 0,
        diferenca: 0,
        unidade: 'status',
        tolerancia: 0,
        situacao: 'erro',
        explicacao: `Pendente de mapeamento: rubrica "${item.label}" ${item.compositeKey ? `(${item.compositeKey})` : ''} não foi localizada.`,
      });
    });
    blockingReasons.push(`Existem ${unmappedItems.length} rubricas com status "Pendente de mapeamento".`);
  }

  const checkEquation = (
    quadro: string,
    exercicio: string,
    regra: string,
    informedItem: StatementItem | undefined,
    componentItems: (StatementItem | undefined)[],
    operation: 'sum' | 'custom',
    customCalcFn?: () => number,
    customRefText?: string
  ) => {
    const isCurrent = exercicio === currYear;
    const informedVal = getItemVal(informedItem, isCurrent);
    const n = componentItems.length;
    // Tolerância contábil conservadora para relatórios em R$ milhares: (n + 1) * p / 2, onde p = 0,01 milhar = R$ 10,00
    const tolerance = calcTolerance(n, 0.01);

    let calculatedVal = 0;
    if (operation === 'sum') {
      calculatedVal = componentItems.reduce((acc, curr) => acc + getItemVal(curr, isCurrent), 0);
    } else if (customCalcFn) {
      calculatedVal = customCalcFn();
    }

    const diff = Number((informedVal - calculatedVal).toFixed(4));
    const absDiff = Math.abs(diff);

    let situacao: ValidationStatus = 'conforme';
    let explicacao = 'Valor informado confere exatamente com a soma dos componentes subordinados.';

    if (informedVal === 0 && calculatedVal === 0 && (!informedItem || getItemVal(informedItem, isCurrent) === 0)) {
      situacao = 'conforme';
      explicacao = 'Sem movimentação ou saldo neutro.';
    } else if (absDiff === 0) {
      situacao = 'conforme';
      explicacao = 'Valor informado confere exatamente com a soma dos componentes subordinados (diferença zero).';
    } else if (absDiff <= tolerance) {
      situacao = 'aviso';
      explicacao = `Diferença compatível com arredondamento da origem (${diff > 0 ? '+' : ''}${diff.toFixed(2)} ≤ limite de ±${tolerance.toFixed(4)} R$ mil).`;
    } else {
      situacao = 'erro';
      explicacao = `Divergência de ${diff > 0 ? '+' : ''}${diff.toFixed(2)} R$ mil excede a tolerância de ±${tolerance.toFixed(4)} R$ mil.`;
    }

    const originRefs =
      customRefText ||
      (operation === 'sum' && componentItems.length > 0
        ? `${getItemRef(informedItem, isCurrent)} = ${componentItems.map(c => getItemRef(c, isCurrent)).join(' + ')}`
        : `${getItemRef(informedItem, isCurrent)} vs [${componentItems.map(c => getItemRef(c, isCurrent)).join(', ')}]`);

    const breakdown = componentItems
      .filter((c): c is StatementItem => !!c)
      .map(c => {
        let natureza: 'bruto' | 'redutora' | 'subtotal' | 'parcela' = 'parcela';
        if (c.isRedutora || c.label.startsWith('(-)')) {
          natureza = 'redutora';
        } else if (c.compositeKey?.includes('(Bruto)') || c.id.endsWith('_bruto')) {
          natureza = 'bruto';
        } else if (c.isSubtotal) {
          natureza = 'subtotal';
        }
        return {
          id: c.id,
          label: c.label,
          cellRef: getItemRef(c, isCurrent),
          valorOriginal: getItemVal(c, isCurrent),
          natureza,
        };
      });

    rules.push({
      id: `${quadro}_${exercicio}_${regra.replace(/\s+/g, '_').toLowerCase()}`,
      quadro,
      exercicio,
      regra,
      celulasOrigem: originRefs,
      valorInformado: informedVal,
      valorCalculado: calculatedVal,
      diferenca: diff,
      unidade: 'R$ mil',
      tolerancia: tolerance,
      situacao,
      explicacao,
      componentBreakdown: breakdown,
    });
  };

  // Run validations for both exercises (colCurrent and colPrevious)
  [currYear, prevYear].forEach(yr => {
    const isCurr = yr === currYear;

    // 1. ATIVO: Ativo Circulante = Caixa + Créditos CP + Investimentos e Aplicações CP + Estoques + demais componentes
    const circComponents: StatementItem[] = [
      findAtivo('at_cx_eq'),
      findAtivo('at_cred_cp'),
      findAtivo('at_inv_cp'),
      findAtivo('at_estoques'),
    ].filter((c): c is StatementItem => !!c);

    // Considerar todos os componentes diretamente subordinados ao grupo (nível 1), mesmo que não constem no modelo fixo
    data.ativo.forEach(it => {
      if (
        it.level === 1 &&
        it.compositeKey?.startsWith('Ativo > Ativo Circulante') &&
        !circComponents.some(c => c.id === it.id)
      ) {
        circComponents.push(it);
      }
    });

    checkEquation(
      'Ativo',
      yr,
      'Ativo Circulante = Componentes',
      findAtivo('at_circ'),
      circComponents,
      'sum'
    );

    // 1.1. ATIVO: Composição de Investimentos e Aplicações Temporárias CP (se possuir bruto e redutoras)
    const invBruto = findAtivo('at_inv_cp_bruto');
    const invPerdas = findAtivo('at_inv_cp_perdas');
    if (
      invBruto &&
      invPerdas &&
      (getItemVal(invBruto, isCurr) !== 0 || getItemVal(invPerdas, isCurr) !== 0)
    ) {
      checkEquation(
        'Ativo',
        yr,
        'Investimentos e Aplicações CP = Bruto + Perdas',
        findAtivo('at_inv_cp'),
        [invBruto, invPerdas],
        'sum'
      );
    }

    // 2. ATIVO: Créditos CP Líquidos = Empréstimos Líq + Demais Créditos Líq
    checkEquation(
      'Ativo',
      yr,
      'Créditos a Curto Prazo = Componentes Líquidos',
      findAtivo('at_cred_cp'),
      [findAtivo('at_cred_cp_emp_liq'), findAtivo('at_cred_cp_demais_liq')],
      'sum'
    );

    // 3. ATIVO: Empréstimos CP Líquido = Bruto + Perdas Estimadas
    checkEquation(
      'Ativo',
      yr,
      'Empréstimos CP Líquido = Bruto + Perdas',
      findAtivo('at_cred_cp_emp_liq'),
      [findAtivo('at_cred_cp_emp_bruto'), findAtivo('at_cred_cp_emp_perdas')],
      'sum'
    );

    // 4. ATIVO: Demais Créditos CP Líquido = Bruto + Perdas
    checkEquation(
      'Ativo',
      yr,
      'Demais Créditos e Valores CP = Bruto + Perdas',
      findAtivo('at_cred_cp_demais_liq'),
      [findAtivo('at_cred_cp_demais_bruto'), findAtivo('at_cred_cp_demais_perdas')],
      'sum'
    );

    // 5. ATIVO: Ativo Não Circulante = ARLP + Investimentos + Imobilizado + Intangível
    checkEquation(
      'Ativo',
      yr,
      'Ativo Não Circulante = Componentes',
      findAtivo('at_nao_circ'),
      [findAtivo('at_rlp'), findAtivo('at_invest'), findAtivo('at_imob'), findAtivo('at_intangivel')],
      'sum'
    );

    // 6. ATIVO: Imobilizado = Bens Móveis Líquido + Bens Imóveis Líquido
    checkEquation(
      'Ativo',
      yr,
      'Imobilizado = Móveis Líq. + Imóveis Líq.',
      findAtivo('at_imob'),
      [findAtivo('at_imob_moveis_liq'), findAtivo('at_imob_imoveis_liq')],
      'sum'
    );

    // 7. ATIVO: Bens Móveis Líquido = Bruto + Depreciação Acumulada
    checkEquation(
      'Ativo',
      yr,
      'Bens Móveis Líq. = Bruto + Depreciação',
      findAtivo('at_imob_moveis_liq'),
      [findAtivo('at_imob_moveis_bruto'), findAtivo('at_imob_moveis_deprec')],
      'sum'
    );

    // 8. ATIVO: Bens Imóveis Líquido = Bruto + Depreciação Acumulada
    checkEquation(
      'Ativo',
      yr,
      'Bens Imóveis Líq. = Bruto + Depreciação',
      findAtivo('at_imob_imoveis_liq'),
      [findAtivo('at_imob_imoveis_bruto'), findAtivo('at_imob_imoveis_deprec')],
      'sum'
    );

    // 9. ATIVO: Total do Ativo = Ativo Circulante + Ativo Não Circulante
    checkEquation(
      'Ativo',
      yr,
      'Total do Ativo = Circulante + Não Circulante',
      findAtivo('at_total'),
      [findAtivo('at_circ'), findAtivo('at_nao_circ')],
      'sum'
    );

    // 10. PASSIVO: Passivo Circulante = soma dos 7 componentes (O17..O23 / P17..P23)
    checkEquation(
      'Passivo e PL',
      yr,
      'Passivo Circulante = Componentes',
      findPassivo('pass_circ'),
      [
        findPassivo('pass_trab'),
        findPassivo('pass_emp_cp'),
        findPassivo('pass_fornec'),
        findPassivo('pass_fisc_cp'),
        findPassivo('pass_transf_fisc'),
        findPassivo('pass_prov_cp'),
        findPassivo('pass_demais_cp'),
      ],
      'sum'
    );

    // 11. PASSIVO: Passivo Não Circulante = soma dos 8 componentes (O32..O39 / P32..P39)
    checkEquation(
      'Passivo e PL',
      yr,
      'Passivo Não Circulante = Componentes',
      findPassivo('pass_nao_circ'),
      [
        findPassivo('pass_trab_lp'),
        findPassivo('pass_emp_lp'),
        findPassivo('pass_fornec_lp'),
        findPassivo('pass_fisc_lp'),
        findPassivo('pass_transf_fisc_lp'),
        findPassivo('pass_prov_lp'),
        findPassivo('pass_demais_lp'),
        findPassivo('pass_res_dif'),
      ],
      'sum'
    );

    // 12. PASSIVO: Total do Passivo Exigível = Passivo Circulante + Passivo Não Circulante
    checkEquation(
      'Passivo e PL',
      yr,
      'Total Passivo Exigível = Circulante + Não Circulante',
      findPassivo('pass_exig_total'),
      [findPassivo('pass_circ'), findPassivo('pass_nao_circ')],
      'sum'
    );

    // 13. PASSIVO: Patrimônio Líquido = componentes
    checkEquation(
      'Passivo e PL',
      yr,
      'Patrimônio Líquido = Soma Componentes',
      findPassivo('pl_grupo'),
      [
        findPassivo('pl_social'),
        findPassivo('pl_ajustes_aval'),
        findPassivo('pl_reservas'),
        findPassivo('pl_res_acum'),
      ],
      'sum'
    );

    // 14. PASSIVO: Total Passivo e PL = Passivo Exigível + Patrimônio Líquido
    checkEquation(
      'Passivo e PL',
      yr,
      'Total Passivo e PL = Exigível + PL',
      findPassivo('pass_total'),
      [findPassivo('pass_exig_total'), findPassivo('pl_grupo')],
      'sum'
    );

    // 15. EQUILÍBRIO FUNDAMENTAL: Total do Ativo = Total do Passivo e PL
    checkEquation(
      'Balanço Geral',
      yr,
      'Total do Ativo = Total Passivo e PL',
      findAtivo('at_total'),
      [findPassivo('pass_total')],
      'custom',
      () => getItemVal(findPassivo('pass_total'), isCurr),
      `${getItemRef(findAtivo('at_total'), isCurr)} vs ${getItemRef(findPassivo('pass_total'), isCurr)}`
    );

    // 16. QUADRO FINANCEIRO E PERMANENTE: Ativo (I) = Ativo Fin + Ativo Perm
    checkEquation(
      'Quadro Fin. e Perm.',
      yr,
      'Ativo (I) = Financeiro + Permanente',
      findFP('fp_at_i'),
      [findFP('fp_at_fin'), findFP('fp_at_perm')],
      'sum'
    );

    // 17. QUADRO FINANCEIRO E PERMANENTE: Ativo (I) = Total Ativo
    checkEquation(
      'Quadro Fin. e Perm.',
      yr,
      'Ativo (I) = Total do Ativo Principal',
      findFP('fp_at_i'),
      [findAtivo('at_total')],
      'custom',
      () => getItemVal(findAtivo('at_total'), isCurr),
      `${getItemRef(findFP('fp_at_i'), isCurr)} vs ${getItemRef(findAtivo('at_total'), isCurr)}`
    );

    // 18. QUADRO FINANCEIRO E PERMANENTE: Passivo (II) = Passivo Fin + Passivo Perm
    checkEquation(
      'Quadro Fin. e Perm.',
      yr,
      'Passivo (II) = Financeiro + Permanente',
      findFP('fp_pass_ii'),
      [findFP('fp_pass_fin'), findFP('fp_pass_perm')],
      'sum'
    );

    // 19. QUADRO FINANCEIRO E PERMANENTE: Saldo Patrimonial = I - II
    checkEquation(
      'Quadro Fin. e Perm.',
      yr,
      'Saldo Patrimonial = Ativo (I) - Passivo (II)',
      findFP('fp_saldo_patr'),
      [findFP('fp_at_i'), findFP('fp_pass_ii')],
      'custom',
      () => getItemVal(findFP('fp_at_i'), isCurr) - getItemVal(findFP('fp_pass_ii'), isCurr),
      `${getItemRef(findFP('fp_saldo_patr'), isCurr)} = ${getItemRef(findFP('fp_at_i'), isCurr)} - ${getItemRef(findFP('fp_pass_ii'), isCurr)}`
    );

    // 20. COMPENSAÇÕES: Atos Potenciais Ativos = soma dos componentes
    checkEquation(
      'Compensações',
      yr,
      'Atos Potenciais Ativos = Componentes',
      findComp('comp_at_pot'),
      [
        findComp('comp_at_gar'),
        findComp('comp_at_conv'),
        findComp('comp_at_contrat'),
        findComp('comp_at_outros'),
      ],
      'sum'
    );

    // 21. COMPENSAÇÕES: Atos Potenciais Passivos = soma dos componentes
    checkEquation(
      'Compensações',
      yr,
      'Atos Potenciais Passivos = Componentes',
      findComp('comp_pass_pot'),
      [
        findComp('comp_pass_gar'),
        findComp('comp_pass_conv'),
        findComp('comp_pass_contrat'),
      ],
      'sum'
    );
  });

  // 22. DEMONSTRATIVO DO SUPERÁVIT / DÉFICIT FINANCEIRO (Exercício Corrente)
  // Regra 1: Recursos Vinculados = soma de suas cinco destinações
  const sdVinculadosDestinations: StatementItem[] = [
    findSD('sd_educ'),
    findSD('sd_seg_soc'),
    findSD('sd_prev_soc'),
    findSD('sd_div_pub'),
    findSD('sd_fundos'),
    findSD('sd_extraorc'),
  ].filter((x): x is StatementItem => !!x);

  checkEquation(
    'Superávit/Déficit',
    currYear,
    'Recursos Vinculados = Soma das Destinações',
    findSD('sd_rec_vinc'),
    sdVinculadosDestinations,
    'sum'
  );

  // Regra 2: TOTAL = Recursos Não Vinculados + Recursos Vinculados
  const sdTotalComponents: StatementItem[] = [
    findSD('sd_rec_nao_vinc'),
    findSD('sd_rec_vinc'),
  ].filter((x): x is StatementItem => !!x);

  checkEquation(
    'Superávit/Déficit',
    currYear,
    'TOTAL = Recursos Não Vinculados + Recursos Vinculados',
    findSD('sd_total'),
    sdTotalComponents,
    'sum'
  );

  // 23. FIDELIDADE DOS ARREDONDAMENTOS INDIVIDUAIS
  let fidelityErrors = 0;
  allItems.forEach(item => {
    if (item.valCurrent !== null && item.valCurrent !== undefined && !item.isUnmapped) {
      const expected = roundCommercial(item.valCurrent);
      if (expected === null) fidelityErrors++;
    }
  });

  rules.push({
    id: 'fidelidade_arredondamento',
    quadro: 'Conferência Geral',
    exercicio: currYear,
    regra: 'Fidelidade dos Valores Diagramados',
    celulasOrigem: 'Todas as células diagramadas',
    valorInformado: allItems.length,
    valorCalculado: allItems.length - fidelityErrors,
    diferenca: fidelityErrors,
    unidade: 'itens',
    tolerancia: 0,
    situacao: fidelityErrors === 0 ? 'conforme' : 'erro',
    explicacao: fidelityErrors === 0
      ? 'Todos os valores diagramados correspondem exatamente ao arredondamento comercial da célula original.'
      : `${fidelityErrors} itens com divergência entre valor original e diagramado.`,
  });

  // Avaliação de bloqueio
  const conformeCount = rules.filter(r => r.situacao === 'conforme').length;
  const avisoCount = rules.filter(r => r.situacao === 'aviso').length;
  const erroCount = rules.filter(r => r.situacao === 'erro').length;
  const pendenteCount = rules.filter(r => r.situacao === 'pendente').length;

  if (findAtivo('at_total')?.valCurrent === null) {
    blockingReasons.push('Total do Ativo ausente ou não identificado.');
  }
  if (findPassivo('pass_total')?.valCurrent === null) {
    blockingReasons.push('Total do Passivo e Patrimônio Líquido ausente ou não identificado.');
  }
  if (erroCount > 0) {
    blockingReasons.push(`Existem ${erroCount} divergências ou pendências que impedem a exportação.`);
  }

  const canExport = blockingReasons.length === 0;

  return {
    rules,
    conformeCount,
    avisoCount,
    erroCount,
    pendenteCount,
    canExport,
    blockingReasons,
  };
}
