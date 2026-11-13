import React from 'react';
import {
  FileSpreadsheet,
  FileText,
  CheckCircle2,
  AlertTriangle,
  ShieldCheck,
  LayoutDashboard,
} from 'lucide-react';
import { BalancoData } from '../types/balanco';

export type TabKey = 'import' | 'diagram' | 'dashboard' | 'validation' | 'export';

interface HeaderProps {
  data: BalancoData | null;
  activeTab: TabKey;
  setActiveTab: (tab: TabKey) => void;
  validationStatus: {
    canExport: boolean;
    conformeCount: number;
    avisoCount: number;
    erroCount: number;
  };
}

export const Header: React.FC<HeaderProps> = ({
  data,
  activeTab,
  setActiveTab,
  validationStatus,
}) => {
  return (
    <header className="bg-slate-900 border-b border-slate-800 text-white sticky top-0 z-30 shadow-md">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between py-3 gap-3">
          {/* Logo & Identity */}
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-lg bg-sky-600 flex items-center justify-center text-white shadow-inner font-bold text-lg border border-sky-400/40">
              <ShieldCheck className="w-6 h-6 text-sky-100" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-sky-400 bg-sky-950/80 border border-sky-800/80 px-2 py-0.5 rounded">
                  COMAER • SIAFI
                </span>
                <span className="text-xs text-slate-400">
                  {data ? `UG ${data.metadata.orgaoCode}` : 'Aguardando Carga'}
                </span>
              </div>
              <h1 className="text-lg font-bold tracking-tight text-white flex items-center gap-2">
                Balanço Patrimonial da Aeronáutica
                <span className="text-xs font-normal text-slate-300 hidden sm:inline">
                  (Excel → Word .docx)
                </span>
              </h1>
            </div>
          </div>

          {/* Metadata pill badge */}
          <div className="flex flex-wrap items-center gap-2 text-xs bg-slate-800/80 border border-slate-700 rounded-lg px-3 py-1.5 text-slate-300">
            {data ? (
              <>
                <div>
                  <span className="text-slate-400">Exercício:</span>{' '}
                  <span className="font-semibold text-white">{data.metadata.colCurrent}</span> vs{' '}
                  <span className="text-slate-400">{data.metadata.colPrevious}</span>
                </div>
                <div className="text-slate-500">•</div>
                <div>
                  <span className="text-slate-400">Período:</span>{' '}
                  <span className="font-semibold text-sky-300">{data.metadata.periodo}</span>
                </div>
                <div className="text-slate-500">•</div>
                <div>
                  <span className="text-slate-400">Unidade:</span>{' '}
                  <span className="font-semibold text-emerald-400">{data.metadata.unidadeNormalizada}</span>
                </div>
              </>
            ) : (
              <div className="text-slate-400 italic">
                Nenhum arquivo processado. Carregue o arquivo Excel para iniciar.
              </div>
            )}
          </div>
        </div>

        {/* Tab Navigation */}
        <div className="flex space-x-1 sm:space-x-2 border-t border-slate-800 pt-1 overflow-x-auto">
          <button
            onClick={() => setActiveTab('import')}
            className={`flex items-center space-x-2 px-4 py-2.5 text-sm font-medium border-b-2 transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'import'
                ? 'border-sky-500 text-sky-400 bg-slate-800/60 rounded-t-md'
                : 'border-transparent text-slate-400 hover:text-slate-200 hover:border-slate-700'
            }`}
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>1. Importação</span>
          </button>

          <button
            onClick={() => setActiveTab('diagram')}
            className={`flex items-center space-x-2 px-4 py-2.5 text-sm font-medium border-b-2 transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'diagram'
                ? 'border-sky-500 text-sky-400 bg-slate-800/60 rounded-t-md'
                : 'border-transparent text-slate-400 hover:text-slate-200 hover:border-slate-700'
            }`}
          >
            <FileText className="w-4 h-4" />
            <span>2. Balanço Diagramado</span>
            {!data && (
              <span className="text-[10px] bg-slate-800 text-slate-400 px-1.5 py-0.2 rounded">
                Vazio
              </span>
            )}
          </button>

          <button
            onClick={() => setActiveTab('dashboard')}
            className={`flex items-center space-x-2 px-4 py-2.5 text-sm font-medium border-b-2 transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'dashboard'
                ? 'border-sky-500 text-sky-400 bg-slate-800/60 rounded-t-md'
                : 'border-transparent text-slate-400 hover:text-slate-200 hover:border-slate-700'
            }`}
          >
            <LayoutDashboard className="w-4 h-4 text-sky-400" />
            <span className="font-semibold text-sky-300">3. Dashboard Patrimonial</span>
          </button>

          <button
            onClick={() => setActiveTab('validation')}
            className={`flex items-center space-x-2 px-4 py-2.5 text-sm font-medium border-b-2 transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'validation'
                ? 'border-sky-500 text-sky-400 bg-slate-800/60 rounded-t-md'
                : 'border-transparent text-slate-400 hover:text-slate-200 hover:border-slate-700'
            }`}
          >
            {validationStatus.erroCount > 0 ? (
              <AlertTriangle className="w-4 h-4 text-rose-400" />
            ) : (
              <CheckCircle2 className="w-4 h-4 text-emerald-400" />
            )}
            <span>4. Validações</span>
            {data && (
              validationStatus.erroCount > 0 ? (
                <span className="bg-rose-500/20 text-rose-300 text-xs px-1.5 py-0.5 rounded-full font-bold">
                  {validationStatus.erroCount}
                </span>
              ) : (
                <span className="bg-emerald-500/20 text-emerald-300 text-xs px-1.5 py-0.5 rounded-full font-semibold">
                  OK
                </span>
              )
            )}
          </button>

          <button
            onClick={() => setActiveTab('export')}
            className={`flex items-center space-x-2 px-4 py-2.5 text-sm font-medium border-b-2 transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'export'
                ? 'border-sky-500 text-sky-400 bg-slate-800/60 rounded-t-md'
                : 'border-transparent text-slate-400 hover:text-slate-200 hover:border-slate-700'
            }`}
          >
            <FileText className="w-4 h-4 text-sky-400" />
            <span>5. Exportação Word (.docx)</span>
          </button>
        </div>
      </div>
    </header>
  );
};
