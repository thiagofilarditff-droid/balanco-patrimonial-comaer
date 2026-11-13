import React, { useState } from 'react';
import { BalancoData, StatementItem } from '../types/balanco';
import { formatThousands, formatDecimal } from '../utils/numberFormat';
import { INITIAL_NE_MAPPINGS } from '../services/statementStructure';
import {
  FileText,
  RotateCcw,
  Check,
  AlertTriangle,
  Info,
  LayoutDashboard,
} from 'lucide-react';

interface TabDiagramProps {
  data: BalancoData | null;
  setData: React.Dispatch<React.SetStateAction<BalancoData | null>>;
  onGoToImport?: () => void;
  onGoToDashboard?: () => void;
}

export const TabDiagram: React.FC<TabDiagramProps> = ({
  data,
  setData,
  onGoToImport,
  onGoToDashboard,
}) => {
  const [viewMode, setViewMode] = useState<'diagram' | 'cells' | 'decimals'>('diagram');
  const [activeStatement, setActiveStatement] = useState<string>('all');
  const [editingNeId, setEditingNeId] = useState<string | null>(null);
  const [neInputVal, setNeInputVal] = useState<string>('');

  if (!data) {
    return (
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 text-center max-w-xl mx-auto my-8">
        <div className="w-16 h-16 rounded-full bg-sky-100 text-sky-700 flex items-center justify-center mx-auto mb-4">
          <FileText className="w-8 h-8" />
        </div>
        <h3 className="text-lg font-bold text-slate-900 mb-2">
          Nenhum Balanço Diagramado no Momento
        </h3>
        <p className="text-sm text-slate-600 mb-6">
          Faça a carga do arquivo Excel na aba <strong>1. Importação</strong> e clique no botão <strong>"Gerar Diagramação dos Saldos"</strong> para visualizar os cinco quadros aqui.
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

  const currYear = data.metadata.colCurrent || '2026';
  const prevYear = data.metadata.colPrevious || '2025';

  const handleStartEditNe = (item: StatementItem) => {
    setEditingNeId(item.id);
    setNeInputVal(item.neValue !== null && item.neValue !== undefined ? String(item.neValue) : '');
  };

  const handleSaveNe = (itemId: string, targetSection: 'ativo' | 'passivoPL' | 'compensacoes') => {
    const trimmed = neInputVal.trim();
    const parsedVal = trimmed === '' ? null : trimmed;

    setData(prev => {
      if (!prev) return null;
      const updatedList = prev[targetSection].map(item => {
        if (item.id === itemId) {
          return { ...item, neValue: parsedVal };
        }
        return item;
      });
      return {
        ...prev,
        [targetSection]: updatedList,
      };
    });
    setEditingNeId(null);
  };

  const handleResetAllNe = () => {
    setData(prev => {
      if (!prev) return null;
      const resetAtivo = prev.ativo.map(item => ({
        ...item,
        neValue: item.neKey ? INITIAL_NE_MAPPINGS[item.neKey] ?? null : null,
      }));
      const resetPassivo = prev.passivoPL.map(item => ({
        ...item,
        neValue: item.neKey ? INITIAL_NE_MAPPINGS[item.neKey] ?? null : null,
      }));
      const resetComp = prev.compensacoes.map(item => ({
        ...item,
        neValue: item.neKey ? INITIAL_NE_MAPPINGS[item.neKey] ?? null : null,
      }));
      return {
        ...prev,
        ativo: resetAtivo,
        passivoPL: resetPassivo,
        compensacoes: resetComp,
      };
    });
  };

  const renderValue = (
    val: number | null | undefined,
    cellRef?: string,
    isUnmapped?: boolean
  ) => {
    if (isUnmapped) {
      return (
        <span className="text-rose-600 font-bold text-xs bg-rose-50 px-1.5 py-0.5 rounded border border-rose-200">
          Pendente de mapeamento
        </span>
      );
    }
    if (viewMode === 'cells') {
      return (
        <span className="font-mono text-[11px] text-sky-800 bg-sky-50 px-1 py-0.5 rounded border border-sky-200">
          {cellRef || '-'}
        </span>
      );
    }
    if (viewMode === 'decimals') {
      return formatDecimal(val, 2);
    }
    return formatThousands(val);
  };

  // Render a 4-column statement table with strict official styles
  const renderFourColumnTable = (
    title: string,
    items: StatementItem[],
    sectionType: 'ativo' | 'passivoPL' | 'financeiroPermanente' | 'compensacoes',
    themeBgColor: string,
    showBottomBar = false
  ) => {
    const filteredItems = items.filter(item => {
      if (viewMode === 'cells') return true;
      if (
        item.omitIfZeroBalance &&
        (item.valCurrent === null || item.valCurrent === 0) &&
        (item.valPrevious === null || item.valPrevious === 0)
      ) {
        return false;
      }
      return true;
    });

    return (
      <div className="bg-white my-4 font-['Times_New_Roman',_serif] text-slate-950">
        {/* Top unit */}
        <div className="text-right text-[11px] text-slate-900 pb-0.5 italic font-normal">
          R$ milhares
        </div>

        <table className="w-full text-[13.5px] border-collapse">
          <thead>
            <tr style={{ backgroundColor: themeBgColor }} className="text-white">
              <th className="text-left font-bold py-1 px-3 w-[64%] tracking-wide">
                {title}
              </th>
              <th className="text-center font-bold py-1 px-2 w-[8%]">
                NE
              </th>
              <th className="text-right font-bold py-1 px-3 w-[14%]">
                {currYear}
              </th>
              <th className="text-right font-bold py-1 px-3 w-[14%]">
                {prevYear}
              </th>
            </tr>
          </thead>
          <tbody className="bg-white">
            {filteredItems.map(item => {
              const isTotal = !!item.isTotal;
              const isBold = !!item.isBold || isTotal;
              const isUnderline = !!item.isUnderline && !isTotal;
              const isUnderlineVal = !!item.isUnderlineValue && !isTotal;
              const indentPx = isTotal ? 0 : item.level * 16;

              return (
                <tr
                  key={item.id}
                  style={isTotal ? { backgroundColor: themeBgColor, color: '#FFFFFF' } : { backgroundColor: '#FFFFFF' }}
                  className={`${isTotal ? 'font-bold' : ''} border-b border-transparent hover:bg-sky-50/25 transition-colors`}
                >
                  {/* Descrição */}
                  <td
                    className="py-0.5 px-3 text-left align-middle"
                    style={{ paddingLeft: `${indentPx + 12}px` }}
                  >
                    <span
                      className={`${isBold ? 'font-bold' : 'font-normal'} ${
                        isUnderline ? 'underline decoration-slate-950' : ''
                      }`}
                    >
                      {item.label}
                    </span>
                    {viewMode === 'cells' && item.rowSource && (
                      <span className="ml-2 text-[10px] text-slate-400 font-sans">
                        (Linha {item.rowSource})
                      </span>
                    )}
                  </td>

                  {/* NE */}
                  <td className="py-0.5 px-2 text-center align-middle">
                    {!isTotal ? (
                      editingNeId === item.id ? (
                        <div className="inline-flex items-center gap-1 font-sans">
                          <input
                            type="text"
                            value={neInputVal}
                            onChange={e => setNeInputVal(e.target.value)}
                            className="w-8 text-center text-xs font-bold text-sky-800 bg-white border border-sky-400 rounded px-0.5 py-0.5"
                            autoFocus
                            onKeyDown={e => {
                              if (e.key === 'Enter') {
                                handleSaveNe(
                                  item.id,
                                  sectionType as 'ativo' | 'passivoPL' | 'compensacoes'
                                );
                              }
                            }}
                          />
                          <button
                            onClick={() =>
                              handleSaveNe(
                                item.id,
                                sectionType as 'ativo' | 'passivoPL' | 'compensacoes'
                              )
                            }
                            className="text-emerald-700 hover:text-emerald-900"
                          >
                            <Check className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => handleStartEditNe(item)}
                          className="group/ne inline-flex items-center justify-center text-[#0563C1] hover:text-[#004B87] underline font-medium cursor-pointer"
                          title="Clique para editar número da Nota Explicativa"
                        >
                          {item.neValue !== null && item.neValue !== undefined && String(item.neValue) !== '' ? (
                            <span>{item.neValue}</span>
                          ) : (
                            <span className="opacity-0 group-hover/ne:opacity-100 text-[10px] text-slate-400 no-underline font-sans">
                              +NE
                            </span>
                          )}
                        </button>
                      )
                    ) : null}
                  </td>

                  {/* Valor Atual */}
                  <td className="py-0.5 px-3 text-right tabular-nums align-middle">
                    <span
                      className={`${isBold ? 'font-bold' : 'font-normal'} ${
                        isUnderlineVal ? 'underline decoration-slate-950' : ''
                      }`}
                    >
                      {renderValue(item.valCurrent, item.cellRefCurrent, item.isUnmapped)}
                    </span>
                  </td>

                  {/* Valor Anterior */}
                  <td className="py-0.5 px-3 text-right tabular-nums align-middle">
                    <span
                      className={`${isBold ? 'font-bold' : 'font-normal'} ${
                        isUnderlineVal ? 'underline decoration-slate-950' : ''
                      }`}
                    >
                      {renderValue(item.valPrevious, item.cellRefPrevious, item.isUnmapped)}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* Faixa azul inferior se indicada (ex: Quadro de Compensações) */}
        {showBottomBar && (
          <div
            className="w-full h-2.5 my-0.5"
            style={{ backgroundColor: themeBgColor }}
          />
        )}

        {/* Rodapé Fonte */}
        <div className="text-left text-[11px] text-slate-800 pt-1 font-normal">
          Fonte: SIAFI.
        </div>
      </div>
    );
  };

  // Render Quadro 5: Superávit/Déficit
  const renderSuperavitTable = () => {
    return (
      <div className="bg-white my-4 font-['Times_New_Roman',_serif] text-slate-950">
        <div className="text-right text-[11px] text-slate-900 pb-0.5 italic font-normal">
          R$ milhares
        </div>

        <table className="w-full text-[13.5px] border-collapse">
          <thead>
            <tr style={{ backgroundColor: '#31549A' }} className="text-white">
              <th className="text-left font-bold py-1 px-3 w-[75%]">
                {/* Cabeçalho de descrição vazio por especificação */}
              </th>
              <th className="text-right font-bold py-1 px-3 w-[25%]">
                {currYear}
              </th>
            </tr>
          </thead>
          <tbody className="bg-white">
            {data.superavitDeficit.map(item => {
              const isTotal = !!item.isTotal;
              const isBold = !!item.isBold || isTotal;
              const indentPx = isTotal ? 0 : item.level * 18;

              return (
                <tr
                  key={item.id}
                  style={isTotal ? { backgroundColor: '#31549A', color: '#FFFFFF' } : { backgroundColor: '#FFFFFF' }}
                  className={`${isTotal ? 'font-bold' : ''} border-b border-transparent hover:bg-sky-50/25 transition-colors`}
                >
                  <td
                    className="py-0.5 px-3 text-left align-middle"
                    style={{ paddingLeft: `${indentPx + 12}px` }}
                  >
                    <span className={isBold ? 'font-bold' : 'font-normal'}>
                      {item.label}
                    </span>
                    {viewMode === 'cells' && item.rowSource && (
                      <span className="ml-2 text-[10px] text-slate-400 font-sans">
                        (Linha {item.rowSource})
                      </span>
                    )}
                  </td>
                  <td className="py-0.5 px-3 text-right tabular-nums align-middle">
                    <span className={isBold ? 'font-bold' : 'font-normal'}>
                      {renderValue(item.valCurrent, item.cellRefCurrent, item.isUnmapped)}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        <div className="text-left text-[11px] text-slate-800 pt-1 font-normal">
          Fonte: SIAFI.
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-6">
      {/* Controls Bar */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          {/* Section filter tabs */}
          <div className="flex flex-wrap items-center gap-1.5">
            {[
              { id: 'all', label: 'Todos os 5 Quadros' },
              { id: 'ativo', label: '1. Ativo' },
              { id: 'passivo', label: '2. Passivo e PL' },
              { id: 'fp', label: '3. Fin. e Permanente' },
              { id: 'comp', label: '4. Compensações' },
              { id: 'sd', label: '5. Superávit/Déficit' },
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => setActiveStatement(tab.id)}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${
                  activeStatement === tab.id
                    ? 'bg-sky-700 text-white shadow-sm'
                    : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* View Toggles & NE reset */}
          <div className="flex items-center gap-2">
            <div className="flex items-center bg-slate-100 rounded-lg p-0.5 border border-slate-200 text-xs">
              <button
                onClick={() => setViewMode('diagram')}
                className={`px-2.5 py-1 rounded font-medium transition-colors ${
                  viewMode === 'diagram'
                    ? 'bg-white text-sky-800 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Diagramado Oficial
              </button>
              <button
                onClick={() => setViewMode('decimals')}
                className={`px-2.5 py-1 rounded font-medium transition-colors ${
                  viewMode === 'decimals'
                    ? 'bg-white text-sky-800 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Com Decimais
              </button>
              <button
                onClick={() => setViewMode('cells')}
                className={`px-2.5 py-1 rounded font-medium transition-colors ${
                  viewMode === 'cells'
                    ? 'bg-white text-sky-800 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Células Excel
              </button>
            </div>

            <button
              onClick={handleResetAllNe}
              className="inline-flex items-center space-x-1.5 px-2.5 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 transition-colors border border-slate-300 cursor-pointer"
              title="Restaurar valores padrões das Notas Explicativas (1 a 16)"
            >
              <RotateCcw className="w-3.5 h-3.5 text-slate-600" />
              <span>Restaurar NEs</span>
            </button>

            {onGoToDashboard && (
              <button
                onClick={onGoToDashboard}
                className="inline-flex items-center space-x-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-sky-50 hover:bg-sky-100 text-sky-800 transition-colors border border-sky-200 cursor-pointer"
                title="Visualizar Dashboard Patrimonial Executivo"
              >
                <LayoutDashboard className="w-3.5 h-3.5 text-sky-600" />
                <span>Dashboard</span>
              </button>
            )}
          </div>
        </div>

        <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
          <div className="flex items-center gap-2">
            <Info className="w-4 h-4 text-sky-600" />
            <span>
              Diagramação fidedigna às imagens oficiais do COMAER. Fundo branco uniforme, recuos hierárquicos e sublinhados normativos preservados.
            </span>
          </div>
          <span className="font-mono text-slate-400">
            Cores: #285F98 (Principal) • #31549A (Complementares)
          </span>
        </div>
      </div>

      {/* Main Statements Render Area (Simulated Document Sheet) */}
      <div className="bg-slate-200/60 p-4 sm:p-8 rounded-xl border border-slate-300 flex flex-col items-center">
        <div className="w-full max-w-[960px] bg-white shadow-xl rounded-none p-6 sm:p-12 border border-slate-300 text-slate-950">
          {/* Header on Document (Centralizado, SEM subtítulo por especificação) */}
          <div className="text-center mb-6">
            <h2 className="text-xl sm:text-2xl font-bold font-['Times_New_Roman',_serif] tracking-normal text-slate-950">
              Balanço Patrimonial
            </h2>
          </div>

          {/* QUADRO 1: ATIVO (#285F98) */}
          {(activeStatement === 'all' || activeStatement === 'ativo') && (
            <div>
              {renderFourColumnTable('ATIVO', data.ativo, 'ativo', '#285F98', false)}
            </div>
          )}

          {/* QUADRO 2: PASSIVO E PATRIMÔNIO LÍQUIDO (#285F98) */}
          {(activeStatement === 'all' || activeStatement === 'passivo') && (
            <div>
              {activeStatement === 'all' && (
                <div className="my-8 border-b-2 border-dashed border-slate-300 relative text-center">
                  <span className="bg-white px-3 text-xs text-slate-400 uppercase tracking-widest font-sans font-semibold absolute -top-2.5 left-1/2 -translate-x-1/2">
                    Quebra de Página 1 → Página 2
                  </span>
                </div>
              )}
              {renderFourColumnTable(
                'PASSIVO E PATRIMÔNIO LÍQUIDO',
                data.passivoPL,
                'passivoPL',
                '#285F98',
                false
              )}
            </div>
          )}

          {/* QUADROS COMPLEMENTARES (#31549A) */}
          {(activeStatement === 'all' ||
            activeStatement === 'fp' ||
            activeStatement === 'comp' ||
            activeStatement === 'sd') && (
            <div>
              {activeStatement === 'all' && (
                <div className="my-8 border-b-2 border-dashed border-slate-300 relative text-center">
                  <span className="bg-white px-3 text-xs text-slate-400 uppercase tracking-widest font-sans font-semibold absolute -top-2.5 left-1/2 -translate-x-1/2">
                    Quebra de Página 2 → Página 3 (Quadros Complementares)
                  </span>
                </div>
              )}

              {/* Quadro 3: Ativos e Passivos Financeiros e Permanente */}
              {(activeStatement === 'all' || activeStatement === 'fp') && (
                <div className="my-6">
                  <h3 className="text-base font-bold text-center font-['Times_New_Roman',_serif] text-slate-950 mb-1">
                    Quadro de Ativos e Passivos Financeiros e Permanente
                  </h3>
                  {renderFourColumnTable(
                    '',
                    data.financeiroPermanente,
                    'financeiroPermanente',
                    '#31549A',
                    false
                  )}
                </div>
              )}

              {/* Quadro 4: Quadro de Compensações (com faixa azul antes de Fonte SIAFI) */}
              {(activeStatement === 'all' || activeStatement === 'comp') && (
                <div className="my-6">
                  <h3 className="text-base font-bold text-center font-['Times_New_Roman',_serif] text-slate-950 mb-1">
                    Quadro de Compensações
                  </h3>
                  {renderFourColumnTable(
                    '',
                    data.compensacoes,
                    'compensacoes',
                    '#31549A',
                    true
                  )}
                </div>
              )}

              {/* Quadro 5: Superávit/Déficit Financeiro */}
              {(activeStatement === 'all' || activeStatement === 'sd') && (
                <div className="my-6">
                  <h3 className="text-base font-bold text-center font-['Times_New_Roman',_serif] text-slate-950 mb-1">
                    Demonstrativo do Superávit / Déficit Financeiro Apurado no Balanço Patrimonial
                  </h3>
                  {renderSuperavitTable()}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
