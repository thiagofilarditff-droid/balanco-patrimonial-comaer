import React, { useState } from 'react';
import { dashboardModel, visibleBalance } from '../services/dashboardModel';
import { displayReportPeriod } from '../utils/reportPeriod';
import { BalancoData, StatementItem } from '../types/balanco';
import { formatThousands, formatDecimal } from '../utils/numberFormat';
import {
  Wallet,
  CreditCard,
  TrendingUp,
  Package,
  Clock,
  PieChart,
  Building2,
  FileSpreadsheet,
  Users,
  Banknote,
  Truck,
  Landmark,
  Layers,
  ShieldAlert,
  Calendar,
  Briefcase,
  Scale,
  Coins,
  ChevronDown,
  ChevronRight,
  Printer,
  FileText,
  SlidersHorizontal,
  Info,
} from 'lucide-react';

interface TabDashboardProps {
  data: BalancoData | null;
  onGoToImport?: () => void;
  onGoToDiagram?: () => void;
  onGoToValidations?: () => void;
}

type ExerciseView = 'current' | 'previous' | 'comparative';
type DetailLevel = 'synthetic' | 'analytic';

export const TabDashboard: React.FC<TabDashboardProps> = ({
  data,
  onGoToImport,
  onGoToDiagram,
  onGoToValidations,
}) => {
  const [exerciseView, setExerciseView] = useState<ExerciseView>('current');
  const [detailLevel, setDetailLevel] = useState<DetailLevel>('synthetic');
  const [useDecimals, setUseDecimals] = useState<boolean>(false);
  const [expandedSections, setExpandedSections] = useState<Record<string, boolean>>({
    at_cred_cp: false,
    at_imob: false,
    at_rlp: false,
  });

  if (!data) {
    return (
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 text-center max-w-xl mx-auto my-8">
        <div className="w-16 h-16 rounded-full bg-blue-100 text-blue-800 flex items-center justify-center mx-auto mb-4">
          <FileSpreadsheet className="w-8 h-8" />
        </div>
        <h3 className="text-lg font-bold text-slate-900 mb-2">
          Dashboard Patrimonial Aguardando Carga de Dados
        </h3>
        <p className="text-sm text-slate-600 mb-6">
          Carregue o arquivo Excel do Balanço Patrimonial na aba <strong>1. Importação</strong> para gerar o resumo visual executivo.
        </p>
        {onGoToImport && (
          <button
            onClick={onGoToImport}
            className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-blue-900 hover:bg-blue-950 text-white font-semibold text-sm shadow-sm transition-all cursor-pointer"
          >
            <span>Ir para Painel de Importação</span>
          </button>
        )}
      </div>
    );
  }

  const currYear = data.metadata.colCurrent || '2026';
  const prevYear = data.metadata.colPrevious || '2025';
  const unidade = data.metadata.unidadeNormalizada || 'R$ milhares';
  const orgaoName = data.metadata.orgaoName || 'Comando da Aeronáutica';
  const orgaoCode = data.metadata.orgaoCode || '52111';
  const periodo = displayReportPeriod(data.metadata.periodo, String(data.metadata.exercicio));
  const grouped = dashboardModel(data);

  // Toggle expansion of sub-components (analytical breakdown)
  const toggleExpand = (id: string) => {
    setExpandedSections(prev => ({
      ...prev,
      [id]: !prev[id],
    }));
  };

  // Format value according to selected format
  const formatVal = (val: number | null | undefined): string => {
    if (val === null || val === undefined || isNaN(val)) return '-';
    if (useDecimals) {
      return formatDecimal(val, 2);
    }
    return formatThousands(val);
  };

  // Helper to get raw value for selected exercise
  const getItemValue = (item: StatementItem | undefined, year: 'current' | 'previous'): number | null => {
    if (!item) return null;
    return year === 'current' ? item.valCurrent : item.valPrevious;
  };

  // Resilient item finders without depending on fixed rows
  const normalize = (str: string) =>
    str
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, ' ')
      .trim();

  const findItem = (list: StatementItem[], id: string, labelKeywords: string[]): StatementItem | undefined => {
    const byId = list.find(it => it.id === id);
    if (byId) return byId;

    return list.find(it => {
      const norm = normalize(it.label);
      return labelKeywords.every(kw => norm.includes(normalize(kw)));
    });
  };

  // --- ATIVO ITEMS ---
  const atCirc = findItem(data.ativo, 'at_circ', ['ativo', 'circulante']);
  const atCxEq = findItem(data.ativo, 'at_cx_eq', ['caixa', 'equivalentes']);
  const atCredCp = findItem(data.ativo, 'at_cred_cp', ['creditos', 'curto prazo']);
  const atCredCpEmpLiq = findItem(data.ativo, 'at_cred_cp_emp_liq', ['emprestimos', 'curto prazo']);
  const atCredCpEmpBruto = findItem(data.ativo, 'at_cred_cp_emp_bruto', ['emprestimos']);
  const atCredCpEmpPerdas = findItem(data.ativo, 'at_cred_cp_emp_perdas', ['perdas', 'emprestimos']);
  const atCredCpDemaisLiq = findItem(data.ativo, 'at_cred_cp_demais_liq', ['demais creditos']);
  const atCredCpDemaisBruto = findItem(data.ativo, 'at_cred_cp_demais_bruto', ['demais creditos']);
  const atCredCpDemaisPerdas = findItem(data.ativo, 'at_cred_cp_demais_perdas', ['perdas', 'demais creditos']);

  const atInvCp = findItem(data.ativo, 'at_inv_cp', ['investimentos', 'curto prazo']);
  const atEstoques = findItem(data.ativo, 'at_estoques', ['estoques']);

  const atNaoCirc = findItem(data.ativo, 'at_nao_circ', ['ativo', 'nao circulante']);
  const atRlp = findItem(data.ativo, 'at_rlp', ['realizavel', 'longo prazo']);
  const atRlpCredLp = findItem(data.ativo, 'at_rlp_cred_lp', ['creditos', 'longo prazo']);
  const atRlpEmpLiq = findItem(data.ativo, 'at_rlp_emp_liq', ['emprestimos']);
  const atRlpDemaisLiq = findItem(data.ativo, 'at_rlp_demais_liq', ['demais creditos']);
  const atRlpEstoques = findItem(data.ativo, 'at_rlp_estoques_lp', ['estoques']);

  const atInvest = findItem(data.ativo, 'at_invest', ['investimentos']);
  const atImob = findItem(data.ativo, 'at_imob', ['imobilizado']);
  const atImobMoveisLiq = findItem(data.ativo, 'at_imob_moveis_liq', ['bens moveis']);
  const atImobMoveisBruto = findItem(data.ativo, 'at_imob_moveis_bruto', ['bens moveis']);
  const atImobMoveisDeprec = findItem(data.ativo, 'at_imob_moveis_deprec', ['depreciacao', 'moveis']);
  const atImobImoveisLiq = findItem(data.ativo, 'at_imob_imoveis_liq', ['bens imoveis']);
  const atImobImoveisBruto = findItem(data.ativo, 'at_imob_imoveis_bruto', ['bens imoveis']);
  const atImobImoveisDeprec = findItem(data.ativo, 'at_imob_imoveis_deprec', ['depreciacao', 'imoveis']);

  const atIntangivel = findItem(data.ativo, 'at_intangivel', ['intangivel']);
  const atTotal = findItem(data.ativo, 'at_total', ['total', 'ativo']);

  // --- PASSIVO E PL ITEMS ---
  const passCirc = findItem(data.passivoPL, 'pass_circ', ['passivo', 'circulante']);
  const passTrab = findItem(data.passivoPL, 'pass_trab', ['trabalh', 'curto prazo']);
  const passEmpCp = findItem(data.passivoPL, 'pass_emp_cp', ['emprestimos', 'curto prazo']);
  const passFornec = findItem(data.passivoPL, 'pass_fornec', ['fornecedores', 'curto prazo']);
  const passFiscCp = findItem(data.passivoPL, 'pass_fisc_cp', ['fiscais', 'curto prazo']);
  const passTransfFisc = findItem(data.passivoPL, 'pass_transf_fisc', ['transferencias fiscais']);
  const passProvCp = findItem(data.passivoPL, 'pass_prov_cp', ['provisoes', 'curto prazo']);
  const passDemaisCp = findItem(data.passivoPL, 'pass_demais_cp', ['demais obrigacoes', 'curto prazo']);

  const passNaoCirc = findItem(data.passivoPL, 'pass_nao_circ', ['passivo', 'nao circulante']);
  const passTrabLp = findItem(data.passivoPL, 'pass_trab_lp', ['trabalh', 'longo prazo']);
  const passEmpLp = findItem(data.passivoPL, 'pass_emp_lp', ['emprestimos', 'longo prazo']);
  const passFornecLp = findItem(data.passivoPL, 'pass_fornec_lp', ['fornecedores', 'longo prazo']);
  const passFiscLp = findItem(data.passivoPL, 'pass_fisc_lp', ['fiscais', 'longo prazo']);
  const passProvLp = findItem(data.passivoPL, 'pass_prov_lp', ['provisoes', 'longo prazo']);
  const passDemaisLp = findItem(data.passivoPL, 'pass_demais_lp', ['demais obrigacoes', 'longo prazo']);
  const passResDif = findItem(data.passivoPL, 'pass_res_dif', ['resultado diferido']);

  const plGrupo = findItem(data.passivoPL, 'pl_grupo', ['patrimonio liquido']);
  const plSocial = findItem(data.passivoPL, 'pl_social', ['patrimonio social']);
  const plAjustesAval = findItem(data.passivoPL, 'pl_ajustes_aval', ['ajustes de avaliacao']);
  const plReservas = findItem(data.passivoPL, 'pl_reservas', ['demais reservas']);
  const plResAcum = findItem(data.passivoPL, 'pl_res_acum', ['resultados acumulados']);
  const passTotal = findItem(data.passivoPL, 'pass_total', ['total', 'passivo', 'patrimonio liquido']);

  // Check if item has balance in at least one year
  const hasBalance = (item: StatementItem | undefined): boolean => {
    if (!item) return false;
    return !!visibleBalance(item);
  };

  // Helper row component
  const renderDetailRow = (
    icon: React.ReactNode,
    label: string,
    item: StatementItem | undefined,
    options?: {
      hasSubItems?: boolean;
      subItemsExpanded?: boolean;
      onToggleExpand?: () => void;
      isSubItem?: boolean;
      isDeduction?: boolean;
      isBold?: boolean;
      themeColor?: 'blue' | 'green';
    }
  ) => {
    if (!hasBalance(item)) return null;
    const isSub = options?.isSubItem;
    const isDed = options?.isDeduction;
    const theme = options?.themeColor || 'blue';

    return (
      <div
        className={`flex items-center justify-between py-2 px-3 sm:px-4 text-sm transition-colors ${
          isSub
            ? 'bg-slate-50/70 pl-8 sm:pl-10 text-xs border-t border-slate-100/60'
            : 'border-b border-slate-100 hover:bg-slate-50/80'
        }`}
      >
        <div className="flex items-center gap-2.5 min-w-0 pr-2">
          {options?.hasSubItems && (
            <button
              onClick={options.onToggleExpand}
              className="text-slate-400 hover:text-slate-700 p-0.5 rounded cursor-pointer transition-colors"
              title={options.subItemsExpanded ? 'Recolher detalhes' : 'Expandir detalhes analíticos'}
            >
              {options.subItemsExpanded ? (
                <ChevronDown className="w-3.5 h-3.5 text-slate-600" />
              ) : (
                <ChevronRight className="w-3.5 h-3.5 text-slate-400" />
              )}
            </button>
          )}

          {!options?.hasSubItems && isSub && (
            <span className="w-3.5 text-slate-300 text-center select-none">•</span>
          )}

          <div
            className={`flex-shrink-0 ${
              isSub
                ? 'text-slate-400'
                : theme === 'blue'
                ? 'text-blue-700'
                : 'text-emerald-700'
            }`}
          >
            {icon}
          </div>

          <span
            className={`truncate ${
              options?.isBold
                ? 'font-semibold text-slate-900'
                : isDed
                ? 'text-rose-700 italic font-medium'
                : isSub
                ? 'text-slate-600 font-normal'
                : 'text-slate-800 font-normal'
            }`}
          >
            {label}
          </span>
        </div>

        {/* Values column */}
        <div className="flex items-center gap-3 flex-shrink-0 tabular-nums">
          {(exerciseView === 'current' || exerciseView === 'comparative') && (
            <div className="text-right min-w-[90px] sm:min-w-[110px]">
              {exerciseView === 'comparative' && (
                <span className="text-[10px] text-slate-400 block -mb-0.5 font-normal">
                  {currYear}
                </span>
              )}
              <span
                className={`font-mono text-sm ${
                  options?.isBold
                    ? 'font-bold text-slate-900'
                    : isDed
                    ? 'text-rose-700 font-medium'
                    : 'text-slate-800'
                }`}
              >
                {formatVal(getItemValue(item, 'current'))}
              </span>
            </div>
          )}

          {(exerciseView === 'previous' || exerciseView === 'comparative') && (
            <div className="text-right min-w-[90px] sm:min-w-[110px]">
              {exerciseView === 'comparative' && (
                <span className="text-[10px] text-slate-400 block -mb-0.5 font-normal">
                  {prevYear}
                </span>
              )}
              <span
                className={`font-mono text-sm ${
                  options?.isBold
                    ? 'font-bold text-slate-600'
                    : isDed
                    ? 'text-rose-600/80 font-medium'
                    : 'text-slate-600'
                }`}
              >
                {formatVal(getItemValue(item, 'previous'))}
              </span>
            </div>
          )}
        </div>
      </div>
    );
  };

  // Group Header component
  const renderGroupHeader = (
    title: string,
    item: StatementItem | undefined,
    theme: 'blue' | 'green'
  ) => {
    if (!hasBalance(item)) return null;
    const isBlue = theme === 'blue';
    const bgClass = isBlue ? 'bg-blue-50/90 text-blue-950 border-blue-100' : 'bg-emerald-50/90 text-emerald-950 border-emerald-100';

    return (
      <div
        className={`flex items-center justify-between py-2.5 px-4 font-bold border-y ${bgClass}`}
      >
        <span className="text-xs sm:text-sm tracking-wide uppercase font-bold">
          {title}
        </span>

        <div className="flex items-center gap-3 tabular-nums">
          {(exerciseView === 'current' || exerciseView === 'comparative') && (
            <div className="text-right min-w-[90px] sm:min-w-[110px]">
              {exerciseView === 'comparative' && (
                <span className="text-[10px] opacity-70 block -mb-0.5 font-normal">
                  {currYear}
                </span>
              )}
              <span className="font-mono text-sm font-bold">
                {formatVal(getItemValue(item, 'current'))}
              </span>
            </div>
          )}

          {(exerciseView === 'previous' || exerciseView === 'comparative') && (
            <div className="text-right min-w-[90px] sm:min-w-[110px]">
              {exerciseView === 'comparative' && (
                <span className="text-[10px] opacity-70 block -mb-0.5 font-normal">
                  {prevYear}
                </span>
              )}
              <span className="font-mono text-sm font-bold opacity-80">
                {formatVal(getItemValue(item, 'previous'))}
              </span>
            </div>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-6">
      {/* Controls Bar */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 flex flex-wrap items-center justify-between gap-4">
        {/* Exercise View Selector */}
        <div className="flex items-center gap-2">
          <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
            Exercício:
          </span>
          <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100">
            <button
              onClick={() => setExerciseView('current')}
              className={`px-3 py-1 text-xs font-medium rounded-md transition-all cursor-pointer ${
                exerciseView === 'current'
                  ? 'bg-white text-blue-900 shadow-sm font-bold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {currYear} (Atual)
            </button>
            <button
              onClick={() => setExerciseView('previous')}
              className={`px-3 py-1 text-xs font-medium rounded-md transition-all cursor-pointer ${
                exerciseView === 'previous'
                  ? 'bg-white text-blue-900 shadow-sm font-bold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {prevYear} (Anterior)
            </button>
            <button
              onClick={() => setExerciseView('comparative')}
              className={`px-3 py-1 text-xs font-medium rounded-md transition-all cursor-pointer ${
                exerciseView === 'comparative'
                  ? 'bg-white text-blue-900 shadow-sm font-bold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Comparativo (Ambos)
            </button>
          </div>
        </div>

        {/* View Options: Synthetic / Analytic & Decimals */}
        <div className="flex flex-wrap items-center gap-3">
          {/* Detail Level Toggle */}
          <div className="flex items-center gap-2">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Nível:
            </span>
            <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100">
              <button
                onClick={() => setDetailLevel('synthetic')}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-all cursor-pointer ${
                  detailLevel === 'synthetic'
                    ? 'bg-white text-slate-900 shadow-sm font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Sintético (Líquidos)
              </button>
              <button
                onClick={() => setDetailLevel('analytic')}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-all cursor-pointer ${
                  detailLevel === 'analytic'
                    ? 'bg-white text-slate-900 shadow-sm font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Analítico (com Deduções)
              </button>
            </div>
          </div>

          {/* Decimals Toggle */}
          <button
            onClick={() => setUseDecimals(!useDecimals)}
            className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-lg border text-xs font-medium transition-colors cursor-pointer ${
              useDecimals
                ? 'bg-blue-50 border-blue-200 text-blue-800 font-semibold'
                : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
            }`}
            title="Alternar entre números inteiros arredondados e precisão com 2 casas decimais"
          >
            <SlidersHorizontal className="w-3.5 h-3.5" />
            <span>{useDecimals ? 'Precisão Real (,00)' : 'Arredondado (Milhares)'}</span>
          </button>

          {/* Quick Print Button */}
          <button
            onClick={() => window.print()}
            className="inline-flex items-center gap-1.5 px-3 py-1 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 transition-colors cursor-pointer"
            title="Imprimir visualização do Dashboard"
          >
            <Printer className="w-3.5 h-3.5" />
            <span>Imprimir</span>
          </button>
        </div>
      </div>

      {/* Main Dashboard Presentation Card (Fundo Branco, Duas Colunas) */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 sm:p-8">
        {/* Header da Apresentação */}
        <div className="mb-4">
          <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-1">
            <h2 className="text-2xl sm:text-3xl font-extrabold text-[#1E3A8A] tracking-tight">
              Balanço Patrimonial
            </h2>
            <div className="text-xs sm:text-sm font-semibold text-slate-500">
              {orgaoName} • UG: {orgaoCode}
            </div>
          </div>

          {/* Período abaixo do título */}
          <div className="text-xs sm:text-sm text-slate-600 font-medium mt-1">
            Período: <span className="font-semibold text-slate-900">{periodo}</span> • Exercício: <span className="font-semibold text-[#1E3A8A]">{currYear}</span> vs <span className="font-medium text-slate-500">{prevYear}</span>
          </div>

          {/* Linha horizontal fina abaixo do cabeçalho */}
          <div className="w-full border-b border-slate-200 mt-3 sm:mt-4" />
        </div>

        {/* Grid com Duas Colunas Lado a Lado (Em telas menores empilha) */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-start">
          {/* ================= COLUNA ESQUERDA: ATIVO (AZUL ESCURO) ================= */}
          <div className="border border-slate-200 rounded-lg overflow-hidden shadow-xs flex flex-col">
            {/* Cabeçalho do Bloco: Fundo escuro azul e letras brancas */}
            <div className="bg-[#1E3A8A] text-white py-3 px-4 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="font-extrabold tracking-wider text-sm sm:text-base uppercase">
                  ATIVO
                </span>
              </div>
              <span className="text-[11px] sm:text-xs text-blue-200 font-medium">
                ({unidade})
              </span>
            </div>

            {/* GRUPO 1: ATIVO CIRCULANTE */}
            <div className="flex flex-col">
              {renderGroupHeader('Ativo Circulante', atCirc, 'blue')}

              {/* Detalhe: Caixa e Equivalentes */}
              {renderDetailRow(
                <Wallet className="w-4 h-4" />,
                'Caixa e Equivalentes de Caixa',
                atCxEq,
                { themeColor: 'blue' }
              )}

              {/* Detalhe: Créditos a Curto Prazo (Líquido) */}
              {renderDetailRow(
                <CreditCard className="w-4 h-4" />,
                'Créditos a Curto Prazo (Líquido)',
                atCredCp,
                {
                  hasSubItems: detailLevel === 'analytic',
                  subItemsExpanded: expandedSections['at_cred_cp'] || detailLevel === 'analytic',
                  onToggleExpand: () => toggleExpand('at_cred_cp'),
                  themeColor: 'blue',
                }
              )}

              {/* Deduções e Brutos subordinados (se analítico ou expandido) */}
              {(detailLevel === 'analytic' || expandedSections['at_cred_cp']) && (
                <div className="bg-slate-50/50">
                  {renderDetailRow(
                    <CreditCard className="w-3.5 h-3.5" />,
                    'Empréstimos Concedidos (Bruto)',
                    atCredCpEmpBruto,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <CreditCard className="w-3.5 h-3.5" />,
                    '(-) Perdas Estimadas em Empréstimos Concedidos',
                    atCredCpEmpPerdas,
                    { isSubItem: true, isDeduction: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <CreditCard className="w-3.5 h-3.5" />,
                    'Demais Créditos e Valores a Curto Prazo (Bruto)',
                    atCredCpDemaisBruto,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <CreditCard className="w-3.5 h-3.5" />,
                    '(-) Perdas Estimadas em Demais Créditos',
                    atCredCpDemaisPerdas,
                    { isSubItem: true, isDeduction: true, themeColor: 'blue' }
                  )}
                </div>
              )}

              {/* Detalhe: Investimentos e Aplicações Temporárias a Curto Prazo */}
              {(hasBalance(atInvCp) || atInvCp) && (
                renderDetailRow(
                  <TrendingUp className="w-4 h-4" />,
                  'Investimentos e Aplicações Temporárias a Curto Prazo',
                  atInvCp,
                  { themeColor: 'blue' }
                )
              )}

              {/* Detalhe: Estoques */}
              {renderDetailRow(
                <Package className="w-4 h-4" />,
                'Estoques',
                atEstoques,
                { themeColor: 'blue' }
              )}
            </div>

            {/* GRUPO 2: ATIVO NÃO CIRCULANTE */}
            <div className="flex flex-col">
              {renderGroupHeader('Ativo Não Circulante', atNaoCirc, 'blue')}

              {/* Detalhe: Ativo Realizável a Longo Prazo */}
              {renderDetailRow(
                <Clock className="w-4 h-4" />,
                'Ativo Realizável a Longo Prazo',
                atRlp,
                {
                  hasSubItems: detailLevel === 'analytic',
                  subItemsExpanded: expandedSections['at_rlp'] || detailLevel === 'analytic',
                  onToggleExpand: () => toggleExpand('at_rlp'),
                  themeColor: 'blue',
                }
              )}

              {(detailLevel === 'analytic' || expandedSections['at_rlp']) && (
                <div className="bg-slate-50/50">
                  {renderDetailRow(
                    <Clock className="w-3.5 h-3.5" />,
                    'Empréstimos Concedidos a LP (Líquido)',
                    atRlpEmpLiq,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <Clock className="w-3.5 h-3.5" />,
                    'Demais Créditos a Longo Prazo (Líquido)',
                    atRlpDemaisLiq,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <Package className="w-3.5 h-3.5" />,
                    'Estoques a Longo Prazo',
                    atRlpEstoques,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                </div>
              )}

              {/* Detalhe: Investimentos (Longo Prazo) */}
              {renderDetailRow(
                <PieChart className="w-4 h-4" />,
                'Investimentos',
                atInvest,
                { themeColor: 'blue' }
              )}

              {/* Detalhe: Imobilizado (Líquido) */}
              {renderDetailRow(
                <Building2 className="w-4 h-4" />,
                'Imobilizado (Líquido)',
                atImob,
                {
                  hasSubItems: detailLevel === 'analytic',
                  subItemsExpanded: expandedSections['at_imob'] || detailLevel === 'analytic',
                  onToggleExpand: () => toggleExpand('at_imob'),
                  themeColor: 'blue',
                }
              )}

              {(detailLevel === 'analytic' || expandedSections['at_imob']) && (
                <div className="bg-slate-50/50">
                  {renderDetailRow(
                    <Building2 className="w-3.5 h-3.5" />,
                    'Bens Móveis (Bruto)',
                    atImobMoveisBruto,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <Building2 className="w-3.5 h-3.5" />,
                    '(-) Depreciação Acumulada de Bens Móveis',
                    atImobMoveisDeprec,
                    { isSubItem: true, isDeduction: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <Building2 className="w-3.5 h-3.5" />,
                    'Bens Imóveis (Bruto)',
                    atImobImoveisBruto,
                    { isSubItem: true, themeColor: 'blue' }
                  )}
                  {renderDetailRow(
                    <Building2 className="w-3.5 h-3.5" />,
                    '(-) Depreciação Acumulada de Bens Imóveis',
                    atImobImoveisDeprec,
                    { isSubItem: true, isDeduction: true, themeColor: 'blue' }
                  )}
                </div>
              )}

              {/* Detalhe: Intangível */}
              {renderDetailRow(
                <Layers className="w-4 h-4" />,
                'Intangível',
                atIntangivel,
                { themeColor: 'blue' }
              )}
            </div>

            {/* FAIXA ESCURA DE TOTAL DO ATIVO */}
            <div className="bg-[#1E3A8A] text-white py-3 px-4 flex items-center justify-between font-bold text-sm sm:text-base mt-auto">
              <span className="uppercase tracking-wider">TOTAL DO ATIVO</span>

              <div className="flex items-center gap-3 tabular-nums">
                {(exerciseView === 'current' || exerciseView === 'comparative') && (
                  <div className="text-right min-w-[90px] sm:min-w-[110px]">
                    {exerciseView === 'comparative' && (
                      <span className="text-[10px] text-blue-200 block -mb-0.5 font-normal">
                        {currYear}
                      </span>
                    )}
                    <span className="font-mono text-base font-extrabold text-white">
                      {formatVal(getItemValue(atTotal, 'current'))}
                    </span>
                  </div>
                )}

                {(exerciseView === 'previous' || exerciseView === 'comparative') && (
                  <div className="text-right min-w-[90px] sm:min-w-[110px]">
                    {exerciseView === 'comparative' && (
                      <span className="text-[10px] text-blue-200 block -mb-0.5 font-normal">
                        {prevYear}
                      </span>
                    )}
                    <span className="font-mono text-base font-extrabold text-blue-100">
                      {formatVal(getItemValue(atTotal, 'previous'))}
                    </span>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* ================= COLUNA DIREITA: PASSIVO E PATRIMÔNIO LÍQUIDO (VERDE ESCURO) ================= */}
          <div className="border border-slate-200 rounded-lg overflow-hidden shadow-xs flex flex-col">
            {/* Cabeçalho do Bloco: Fundo escuro verde e letras brancas */}
            <div className="bg-[#14532D] text-white py-3 px-4 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="font-extrabold tracking-wider text-sm sm:text-base uppercase">
                  PASSIVO E PATRIMÔNIO LÍQUIDO
                </span>
              </div>
              <span className="text-[11px] sm:text-xs text-emerald-200 font-medium">
                ({unidade})
              </span>
            </div>

            {/* GRUPO 1: PASSIVO CIRCULANTE */}
            <div className="flex flex-col">
              {renderGroupHeader('Passivo Circulante', passCirc, 'green')}

              {/* Detalhe: Obrigações Trabalhistas */}
              {renderDetailRow(
                <Users className="w-4 h-4" />,
                'Obrigações Trabalhistas, Previdenciárias e Assistenciais a CP',
                passTrab,
                { themeColor: 'green' }
              )}

              {/* Detalhe: Empréstimos e Financiamentos CP */}
              {renderDetailRow(
                <Banknote className="w-4 h-4" />,
                'Empréstimos e Financiamentos a Curto Prazo',
                passEmpCp,
                { themeColor: 'green' }
              )}

              {/* Detalhe: Fornecedores e Contas a Pagar */}
              {renderDetailRow(
                <Truck className="w-4 h-4" />,
                'Fornecedores e Contas a Pagar a Curto Prazo',
                passFornec,
                { themeColor: 'green' }
              )}

              {/* Detalhe: Obrigações Fiscais CP */}
              {(hasBalance(passFiscCp) || passFiscCp) && (
                renderDetailRow(
                  <Landmark className="w-4 h-4" />,
                  'Obrigações Fiscais a Curto Prazo',
                  passFiscCp,
                  { themeColor: 'green' }
                )
              )}

              {/* Detalhe: Transferências Fiscais CP */}
              {renderDetailRow(
                <Coins className="w-4 h-4" />,
                'Transferências Fiscais a Curto Prazo',
                passTransfFisc,
                { themeColor: 'green' }
              )}

              {/* Detalhe: Provisões CP */}
              {renderDetailRow(
                <ShieldAlert className="w-4 h-4" />,
                'Provisões a Curto Prazo',
                passProvCp,
                { themeColor: 'green' }
              )}

              {/* Detalhe: Demais Obrigações CP */}
              {renderDetailRow(
                <Layers className="w-4 h-4" />,
                'Demais Obrigações a Curto Prazo',
                passDemaisCp,
                { themeColor: 'green' }
              )}
            </div>

            {/* Passivo não circulante: rubricas selecionadas e Demais */}
            <div className="flex flex-col">
              {renderGroupHeader('Passivo Não Circulante', passNaoCirc, 'green')}
              {grouped.passivoNaoCirculante.map(item => (
                <React.Fragment key={item.id}>
                  {renderDetailRow(
                    item.id === 'pass_emp_lp' ? <Banknote className="w-4 h-4" /> : item.id === 'pass_prov_lp' ? <ShieldAlert className="w-4 h-4" /> : <Layers className="w-4 h-4" />,
                    item.label, item, { themeColor: 'green' }
                  )}
                </React.Fragment>
              ))}
            </div>
            <div className="flex flex-col">
              {renderGroupHeader('Patrimônio Líquido', plGrupo, 'green')}
            </div>

            {/* FAIXA ESCURA DE TOTAL DO PASSIVO E PATRIMÔNIO LÍQUIDO */}
            <div className="bg-[#14532D] text-white py-3 px-4 flex items-center justify-between font-bold text-sm sm:text-base mt-auto">
              <span className="uppercase tracking-wider">TOTAL DO PASSIVO E PL</span>

              <div className="flex items-center gap-3 tabular-nums">
                {(exerciseView === 'current' || exerciseView === 'comparative') && (
                  <div className="text-right min-w-[90px] sm:min-w-[110px]">
                    {exerciseView === 'comparative' && (
                      <span className="text-[10px] text-emerald-200 block -mb-0.5 font-normal">
                        {currYear}
                      </span>
                    )}
                    <span className="font-mono text-base font-extrabold text-white">
                      {formatVal(getItemValue(passTotal, 'current'))}
                    </span>
                  </div>
                )}

                {(exerciseView === 'previous' || exerciseView === 'comparative') && (
                  <div className="text-right min-w-[90px] sm:min-w-[110px]">
                    {exerciseView === 'comparative' && (
                      <span className="text-[10px] text-emerald-200 block -mb-0.5 font-normal">
                        {prevYear}
                      </span>
                    )}
                    <span className="font-mono text-base font-extrabold text-emerald-100">
                      {formatVal(getItemValue(passTotal, 'previous'))}
                    </span>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Rodapé com a fonte dos dados */}
        <div className="mt-6 pt-3 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between text-xs text-slate-500 gap-2">
          <div className="flex items-center gap-1.5">
            <Info className="w-3.5 h-3.5 text-slate-400" />
            <span>
              Fonte: SIAFI — Balanço Patrimonial / Sistema Integrado de Administração Financeira
            </span>
          </div>
          <div className="flex items-center gap-4">
            {onGoToDiagram && (
              <button
                onClick={onGoToDiagram}
                className="text-blue-700 hover:text-blue-900 font-medium hover:underline cursor-pointer"
              >
                Ver Balanço Diagramado Completo →
              </button>
            )}
            {onGoToValidations && (
              <button
                onClick={onGoToValidations}
                className="text-emerald-700 hover:text-emerald-900 font-medium hover:underline cursor-pointer"
              >
                Ver Validações Contábeis →
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
