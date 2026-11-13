import React, { useState } from 'react';
import { BalancoData } from '../types/balanco';
import { validateBalanco, ValidationSummary } from '../services/validator';
import { formatDecimal } from '../utils/numberFormat';
import {
  CheckCircle2,
  AlertTriangle,
  XCircle,
  HelpCircle,
  Filter,
  ShieldCheck,
  Scale,
  FileCheck,
  ChevronDown,
  ChevronRight,
  ListOrdered,
} from 'lucide-react';

interface TabValidationsProps {
  data: BalancoData | null;
  onGoToImport?: () => void;
}

export const TabValidations: React.FC<TabValidationsProps> = ({ data, onGoToImport }) => {
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [quadroFilter, setQuadroFilter] = useState<string>('all');
  const [expandedRuleId, setExpandedRuleId] = useState<string | null>(null);

  if (!data) {
    return (
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 text-center max-w-xl mx-auto my-8">
        <div className="w-16 h-16 rounded-full bg-sky-100 text-sky-700 flex items-center justify-center mx-auto mb-4">
          <Scale className="w-8 h-8" />
        </div>
        <h3 className="text-lg font-bold text-slate-900 mb-2">
          Nenhuma Validação Ativa
        </h3>
        <p className="text-sm text-slate-600 mb-6">
          Carregue seu arquivo Excel na aba <strong>Importação</strong> e clique em <strong>"Gerar Diagramação dos Saldos"</strong> para inspecionar os laudos e tolerâncias contábeis.
        </p>
        {onGoToImport && (
          <button
            onClick={onGoToImport}
            className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-sky-700 hover:bg-sky-800 text-white font-semibold text-sm shadow-sm transition-all"
          >
            <span>Ir para Painel de Importação</span>
          </button>
        )}
      </div>
    );
  }

  const summary: ValidationSummary = validateBalanco(data);

  const filteredRules = summary.rules.filter(rule => {
    if (statusFilter !== 'all' && rule.situacao !== statusFilter) {
      return false;
    }
    if (quadroFilter !== 'all' && rule.quadro !== quadroFilter) {
      return false;
    }
    return true;
  });

  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'conforme':
        return (
          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
            Conforme
          </span>
        );
      case 'aviso':
        return (
          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-amber-100 text-amber-800 border border-amber-200">
            <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
            Aviso de Arredondamento
          </span>
        );
      case 'erro':
        return (
          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-rose-100 text-rose-800 border border-rose-200">
            <XCircle className="w-3.5 h-3.5 text-rose-600" />
            Erro Aritmético
          </span>
        );
      default:
        return (
          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-slate-100 text-slate-800 border border-slate-200">
            <HelpCircle className="w-3.5 h-3.5 text-slate-500" />
            Pendente
          </span>
        );
    }
  };

  return (
    <div className="space-y-6">
      {/* Metric Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card Conforme */}
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center gap-4">
          <div className="w-12 h-12 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold">
            <CheckCircle2 className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs font-semibold uppercase text-slate-500 tracking-wider">
              Conformes
            </span>
            <p className="text-2xl font-bold text-slate-900">{summary.conformeCount}</p>
            <span className="text-[11px] text-emerald-700">Equações exatas</span>
          </div>
        </div>

        {/* Card Avisos */}
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center gap-4">
          <div className="w-12 h-12 rounded-lg bg-amber-100 text-amber-700 flex items-center justify-center font-bold">
            <AlertTriangle className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs font-semibold uppercase text-slate-500 tracking-wider">
              Arredondamentos
            </span>
            <p className="text-2xl font-bold text-amber-800">{summary.avisoCount}</p>
            <span className="text-[11px] text-amber-700">Dentro da tolerância</span>
          </div>
        </div>

        {/* Card Erros */}
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center gap-4">
          <div className="w-12 h-12 rounded-lg bg-rose-100 text-rose-700 flex items-center justify-center font-bold">
            <XCircle className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs font-semibold uppercase text-slate-500 tracking-wider">
              Divergências Críticas
            </span>
            <p className="text-2xl font-bold text-rose-700">{summary.erroCount}</p>
            <span className="text-[11px] text-rose-600">
              {summary.erroCount > 0 ? 'Bloqueia exportação' : 'Nenhuma divergência'}
            </span>
          </div>
        </div>

        {/* Card Export Status */}
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center gap-4">
          <div
            className={`w-12 h-12 rounded-lg flex items-center justify-center font-bold ${
              summary.canExport
                ? 'bg-sky-100 text-sky-800'
                : 'bg-rose-100 text-rose-800'
            }`}
          >
            <ShieldCheck className="w-6 h-6" />
          </div>
          <div>
            <span className="text-xs font-semibold uppercase text-slate-500 tracking-wider">
              Status para Word
            </span>
            <p
              className={`text-lg font-bold ${
                summary.canExport ? 'text-emerald-700' : 'text-rose-700'
              }`}
            >
              {summary.canExport ? 'Aprovado' : 'Bloqueado'}
            </p>
            <span className="text-[11px] text-slate-500">
              {summary.canExport ? 'Pronto para download' : 'Pendências ativas'}
            </span>
          </div>
        </div>
      </div>

      {/* Tolerance & Methodology Explanatory Box */}
      <div className="bg-sky-50 border border-sky-200 rounded-xl p-4 text-xs text-sky-950 flex items-start gap-3">
        <Scale className="w-5 h-5 text-sky-700 shrink-0 mt-0.5" />
        <div className="space-y-1">
          <p className="font-bold text-sky-900">
            Critério Técnico de Auditoria Contábil e Tolerância de Arredondamento da Origem:
          </p>
          <p className="text-sky-800 leading-relaxed">
            A consistência aritmética avalia cada soma subordinada da origem SIAFI. Para uma soma de <em>n</em> parcelas
            e um total com precisão de duas casas decimais em milhares de reais (<em>p</em> = 0,01 milhar = R$ 10,00),
            adota-se o limite conservador: <strong>Tolerância = (n + 1) × p ÷ 2</strong>.
            Diferenças não nulas dentro desse limite são classificadas como "Aviso de arredondamento da origem", permitindo
            a exportação fidedigna sem distorcer os dados originais publicados.
          </p>
        </div>
      </div>

      {/* Filter and Table Container */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        {/* Table Controls */}
        <div className="p-4 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-slate-500" />
            <span className="text-xs font-bold uppercase tracking-wider text-slate-600">
              Filtros:
            </span>
            <select
              value={statusFilter}
              onChange={e => setStatusFilter(e.target.value)}
              className="text-xs bg-white border border-slate-300 rounded-lg px-2.5 py-1 text-slate-800 font-medium"
            >
              <option value="all">Todas as Situações ({summary.rules.length})</option>
              <option value="conforme">Apenas Conformes ({summary.conformeCount})</option>
              <option value="aviso">Apenas Avisos ({summary.avisoCount})</option>
              <option value="erro">Apenas Erros ({summary.erroCount})</option>
            </select>

            <select
              value={quadroFilter}
              onChange={e => setQuadroFilter(e.target.value)}
              className="text-xs bg-white border border-slate-300 rounded-lg px-2.5 py-1 text-slate-800 font-medium"
            >
              <option value="all">Todos os Quadros</option>
              <option value="Ativo">Ativo</option>
              <option value="Passivo e PL">Passivo e PL</option>
              <option value="Balanço Geral">Balanço Geral</option>
              <option value="Quadro Fin. e Perm.">Quadro Fin. e Perm.</option>
              <option value="Compensações">Compensações</option>
              <option value="Superávit/Déficit">Superávit/Déficit</option>
              <option value="Conferência Geral">Conferência Geral</option>
            </select>
          </div>

          <div className="text-xs text-slate-500 font-medium">
            Exibindo <strong>{filteredRules.length}</strong> de {summary.rules.length} verificações
          </div>
        </div>

        {/* Validation Rules Table */}
        <div className="overflow-x-auto">
          <table className="w-full text-xs text-left border-collapse">
            <thead>
              <tr className="bg-slate-100 text-slate-700 uppercase font-semibold text-[11px] border-b border-slate-200">
                <th className="py-2.5 px-3">Quadro / Exercício</th>
                <th className="py-2.5 px-3">Regra Validada</th>
                <th className="py-2.5 px-3">Células Origem</th>
                <th className="py-2.5 px-3 text-right">Valor Informado</th>
                <th className="py-2.5 px-3 text-right">Valor Calculado</th>
                <th className="py-2.5 px-3 text-right">Diferença</th>
                <th className="py-2.5 px-3 text-right">Tolerância</th>
                <th className="py-2.5 px-3 text-center">Situação</th>
                <th className="py-2.5 px-3">Explicação / Laudo Técnico</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredRules.map(rule => {
                const isExpanded = expandedRuleId === rule.id;
                const hasBreakdown = rule.componentBreakdown && rule.componentBreakdown.length > 0;

                return (
                  <React.Fragment key={rule.id}>
                    <tr
                      onClick={() => hasBreakdown && setExpandedRuleId(isExpanded ? null : rule.id)}
                      className={`hover:bg-slate-50 transition-colors ${hasBreakdown ? 'cursor-pointer' : ''} ${
                        rule.situacao === 'erro'
                          ? 'bg-rose-50/40'
                          : rule.situacao === 'aviso'
                          ? 'bg-amber-50/30'
                          : ''
                      }`}
                    >
                      {/* Quadro & Exercício */}
                      <td className="py-2.5 px-3 font-semibold text-slate-800 whitespace-nowrap">
                        <div className="flex items-center gap-1.5">
                          {hasBreakdown && (
                            <span className="text-slate-400">
                              {isExpanded ? (
                                <ChevronDown className="w-3.5 h-3.5" />
                              ) : (
                                <ChevronRight className="w-3.5 h-3.5" />
                              )}
                            </span>
                          )}
                          <div>
                            <div>{rule.quadro}</div>
                            <span className="text-[10px] text-slate-500 font-normal">
                              Exer. {rule.exercicio}
                            </span>
                          </div>
                        </div>
                      </td>

                      {/* Regra */}
                      <td className="py-2.5 px-3 font-medium text-slate-900">
                        <div className="flex items-center gap-1.5">
                          <span>{rule.regra}</span>
                          {hasBreakdown && (
                            <span className="inline-flex items-center gap-0.5 text-[10px] font-normal px-1.5 py-0.2 rounded bg-sky-100 text-sky-800">
                              <ListOrdered className="w-2.5 h-2.5" />
                              {rule.componentBreakdown!.length} parcelas
                            </span>
                          )}
                        </div>
                      </td>

                      {/* Células Origem */}
                      <td className="py-2.5 px-3 font-mono text-[11px] text-sky-800 max-w-[150px] truncate" title={rule.celulasOrigem}>
                        {rule.celulasOrigem}
                      </td>

                      {/* Valor Informado */}
                      <td className="py-2.5 px-3 text-right tabular-nums font-medium text-slate-900 whitespace-nowrap">
                        {formatDecimal(rule.valorInformado, 2)}
                      </td>

                      {/* Valor Calculado */}
                      <td className="py-2.5 px-3 text-right tabular-nums text-slate-700 whitespace-nowrap">
                        {formatDecimal(rule.valorCalculado, 2)}
                      </td>

                      {/* Diferença */}
                      <td className="py-2.5 px-3 text-right tabular-nums font-bold whitespace-nowrap">
                        <span
                          className={
                            Math.abs(rule.diferenca) === 0
                              ? 'text-emerald-700'
                              : Math.abs(rule.diferenca) <= rule.tolerancia
                              ? 'text-amber-700'
                              : 'text-rose-700'
                          }
                        >
                          {formatDecimal(rule.diferenca, 2)}
                        </span>
                      </td>

                      {/* Tolerância */}
                      <td className="py-2.5 px-3 text-right tabular-nums text-slate-500 whitespace-nowrap">
                        ±{rule.tolerancia.toFixed(2)}
                      </td>

                      {/* Situação */}
                      <td className="py-2.5 px-3 text-center whitespace-nowrap">
                        {getStatusBadge(rule.situacao)}
                      </td>

                      {/* Explicação */}
                      <td className="py-2.5 px-3 text-slate-600 max-w-[260px] text-[11px]">
                        {rule.explicacao}
                      </td>
                    </tr>

                    {/* Linha Expansível de Detalhamento das Parcelas */}
                    {isExpanded && hasBreakdown && (
                      <tr className="bg-sky-50/50 border-b border-sky-200">
                        <td colSpan={9} className="p-4">
                          <div className="bg-white rounded-lg border border-sky-200 p-3 shadow-xs">
                            <div className="flex items-center justify-between mb-2 pb-1.5 border-b border-slate-100">
                              <span className="text-xs font-bold text-sky-950 flex items-center gap-1.5">
                                <ListOrdered className="w-3.5 h-3.5 text-sky-700" />
                                Detalhamento das Rubricas Somadas — {rule.regra} ({rule.exercicio})
                              </span>
                              <span className="text-[11px] text-slate-500">
                                Soma das parcelas: <strong className="text-slate-800">{formatDecimal(rule.valorCalculado, 2)} R$ mil</strong>
                              </span>
                            </div>

                            <table className="w-full text-xs">
                              <thead>
                                <tr className="text-[11px] text-slate-500 border-b border-slate-100 uppercase">
                                  <th className="text-left py-1.5 px-2 font-semibold">Rubrica Contábil</th>
                                  <th className="text-center py-1.5 px-2 font-semibold w-32">Classificação</th>
                                  <th className="text-center py-1.5 px-2 font-semibold w-28">Célula Origem</th>
                                  <th className="text-right py-1.5 px-2 font-semibold w-36">Valor Original</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {rule.componentBreakdown!.map((comp, idx) => (
                                  <tr key={comp.id || idx} className="hover:bg-slate-50">
                                    <td className="py-1.5 px-2 text-slate-800 font-medium">
                                      {comp.label}
                                    </td>
                                    <td className="py-1.5 px-2 text-center">
                                      {comp.natureza === 'redutora' || comp.label.startsWith('(-)') || (comp.valorOriginal !== null && comp.valorOriginal < 0) ? (
                                        <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                                          Conta Redutora (-)
                                        </span>
                                      ) : comp.natureza === 'bruto' ? (
                                        <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-sky-50 text-sky-700 border border-sky-200">
                                          Valor Bruto
                                        </span>
                                      ) : comp.natureza === 'subtotal' ? (
                                        <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200">
                                          Subtotal Líquido
                                        </span>
                                      ) : (
                                        <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-slate-100 text-slate-600">
                                          Parcela
                                        </span>
                                      )}
                                    </td>
                                    <td className="py-1.5 px-2 text-center font-mono text-sky-800 text-[11px]">
                                      {comp.cellRef}
                                    </td>
                                    <td className="py-1.5 px-2 text-right tabular-nums font-mono">
                                      {comp.valorOriginal !== null && comp.valorOriginal !== undefined && comp.valorOriginal !== 0 ? (
                                        <span className={comp.valorOriginal < 0 ? 'text-rose-700 font-bold' : 'text-slate-900 font-medium'}>
                                          {formatDecimal(comp.valorOriginal, 2)}
                                        </span>
                                      ) : (
                                        <span className="text-slate-400 font-sans italic">-</span>
                                      )}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                              <tfoot>
                                <tr className="border-t-2 border-slate-200 font-bold bg-slate-50 text-[11px]">
                                  <td className="py-1.5 px-2 text-slate-900" colSpan={2}>
                                    Soma dos Componentes (Saldo Calculado)
                                  </td>
                                  <td className="py-1.5 px-2 text-center text-slate-500 font-mono text-[10px]">
                                    {rule.celulasOrigem}
                                  </td>
                                  <td className="py-1.5 px-2 text-right tabular-nums text-sky-900 font-mono">
                                    {formatDecimal(rule.valorCalculado, 2)}
                                  </td>
                                </tr>
                                <tr className="border-t border-slate-200 bg-sky-50/40 text-[11px]">
                                  <td className="py-1 px-2 text-sky-900 font-semibold" colSpan={2}>
                                    Saldo Líquido Informado na Origem:
                                  </td>
                                  <td className="py-1 px-2 text-center font-mono text-sky-800 text-[10px]">
                                    {rule.celulasOrigem.split(' = ')[0]}
                                  </td>
                                  <td className="py-1 px-2 text-right tabular-nums text-sky-950 font-bold font-mono">
                                    {formatDecimal(rule.valorInformado, 2)}
                                  </td>
                                </tr>
                                <tr className="border-t border-slate-200 bg-white text-[11px]">
                                  <td className="py-1 px-2 font-semibold" colSpan={2}>
                                    Diferença (Informado − Calculado):
                                  </td>
                                  <td className="py-1 px-2 text-center">
                                    {rule.situacao === 'conforme' ? (
                                      <span className="text-[10px] font-bold text-emerald-700 bg-emerald-100 px-1.5 py-0.5 rounded">
                                        Conforme
                                      </span>
                                    ) : rule.situacao === 'aviso' ? (
                                      <span className="text-[10px] font-bold text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">
                                        Arredondamento
                                      </span>
                                    ) : (
                                      <span className="text-[10px] font-bold text-rose-700 bg-rose-100 px-1.5 py-0.5 rounded">
                                        Divergente
                                      </span>
                                    )}
                                  </td>
                                  <td className="py-1 px-2 text-right tabular-nums font-bold font-mono">
                                    <span className={Math.abs(rule.diferenca) === 0 ? 'text-emerald-700' : 'text-amber-700'}>
                                      {formatDecimal(rule.diferenca, 2)}
                                    </span>
                                  </td>
                                </tr>
                              </tfoot>
                            </table>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
