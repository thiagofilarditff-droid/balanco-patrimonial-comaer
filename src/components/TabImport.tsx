import React, { useRef, useState } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  Building2,
  Calendar,
  Layers,
  Coins,
  ArrowRight,
  Play,
  FileCheck2,
  Trash2,
} from 'lucide-react';
import { BalancoData } from '../types/balanco';
import { parseBalancoExcel } from '../services/excelParser';

interface TabImportProps {
  data: BalancoData | null;
  setData: React.Dispatch<React.SetStateAction<BalancoData | null>>;
  onGoToDiagram: () => void;
}

interface LoadedFileInfo {
  name: string;
  size: number;
  buffer: ArrayBuffer;
}

export const TabImport: React.FC<TabImportProps> = ({
  data,
  setData,
  onGoToDiagram,
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [loadedFile, setLoadedFile] = useState<LoadedFileInfo | null>(null);
  const [statusMessage, setStatusMessage] = useState<{
    type: 'success' | 'error' | 'info';
    text: string;
    warnings?: string[];
  } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const handleFileSelection = (file: File) => {
    if (!file.name.endsWith('.xls') && !file.name.endsWith('.xlsx')) {
      setStatusMessage({
        type: 'error',
        text: 'Formato não suportado. Por favor, selecione um arquivo Excel (.xls ou .xlsx).',
      });
      return;
    }

    const reader = new FileReader();
    reader.onload = event => {
      const buffer = event.target?.result as ArrayBuffer;
      if (buffer) {
        setLoadedFile({
          name: file.name,
          size: file.size,
          buffer,
        });
        setStatusMessage({
          type: 'info',
          text: `Arquivo "${file.name}" (${formatFileSize(file.size)}) carregado. Clique no botão abaixo para gerar a diagramação dos saldos.`,
        });
      }
    };
    reader.onerror = () => {
      setStatusMessage({
        type: 'error',
        text: 'Falha ao ler o arquivo selecionado.',
      });
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFileSelection(file);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFileSelection(file);
    }
  };

  // Button Action: "Gerar Diagramação dos Saldos"
  const handleGenerateDiagram = async () => {
    if (!loadedFile) {
      setStatusMessage({
        type: 'error',
        text: 'Nenhum arquivo carregado. Por favor, selecione ou arraste um arquivo Excel primeiro.',
      });
      return;
    }

    setIsProcessing(true);
    setStatusMessage(null);

    try {
      const result = await parseBalancoExcel(loadedFile.buffer);
      setData(result.data);
      setStatusMessage({
        type: 'success',
        text: `Diagramação gerada com sucesso a partir de "${loadedFile.name}"! Aba identificada: "${result.data.metadata.sourceSheetName}".`,
        warnings: result.warnings,
      });

      // Navigate directly to the diagram tab as requested
      setTimeout(() => {
        onGoToDiagram();
      }, 500);
    } catch (err) {
      console.error(err);
      setStatusMessage({
        type: 'error',
        text: `Erro ao processar e diagramar arquivo: ${err instanceof Error ? err.message : String(err)}`,
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleClearFile = () => {
    setLoadedFile(null);
    setData(null);
    setStatusMessage(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleSheetChange = async (sheetName: string) => {
    if (!loadedFile) return;
    setIsProcessing(true);
    try {
      const result = await parseBalancoExcel(loadedFile.buffer, sheetName);
      setData(result.data);
      setStatusMessage({
        type: 'success',
        text: `Aba alterada para "${sheetName}" e saldos reprocessados.`,
        warnings: result.warnings,
      });
    } catch (err) {
      setStatusMessage({
        type: 'error',
        text: `Erro ao processar aba "${sheetName}": ${err instanceof Error ? err.message : String(err)}`,
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const toggleUnitConfirm = () => {
    if (!data) return;
    setData(prev => {
      if (!prev) return null;
      return {
        ...prev,
        metadata: {
          ...prev.metadata,
          unitConfirmed: !prev.metadata.unitConfirmed,
        },
      };
    });
  };

  return (
    <div className="space-y-6">
      {/* Upload Zone & Actions */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="pb-5 border-b border-slate-100 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
          <div>
            <h2 className="text-xl font-bold text-slate-900 flex items-center gap-2">
              <FileSpreadsheet className="w-6 h-6 text-sky-700" />
              Carga do Balanço Patrimonial (Excel)
            </h2>
            <p className="text-sm text-slate-600 mt-1">
              Faça a carga do arquivo original <strong>.xls</strong> ou <strong>.xlsx</strong> do SIAFI.
              Após a carga, aperte o botão para gerar a diagramação fiel dos saldos.
            </p>
          </div>

          {loadedFile && (
            <button
              onClick={handleClearFile}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 transition-colors w-fit"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Remover Arquivo</span>
            </button>
          )}
        </div>

        {/* Drag & Drop Container */}
        <div className="mt-6">
          <div
            onDragOver={e => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all ${
              isDragging
                ? 'border-sky-500 bg-sky-50/60 scale-[0.99]'
                : loadedFile
                ? 'border-emerald-400 bg-emerald-50/20 hover:bg-emerald-50/40'
                : 'border-slate-300 hover:border-sky-400 hover:bg-slate-50/60'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={handleFileChange}
              className="hidden"
            />
            <div className="flex flex-col items-center justify-center space-y-3">
              <div
                className={`w-14 h-14 rounded-full flex items-center justify-center shadow-sm ${
                  loadedFile
                    ? 'bg-emerald-100 text-emerald-700'
                    : 'bg-sky-100 text-sky-700'
                }`}
              >
                {loadedFile ? (
                  <FileCheck2 className="w-7 h-7" />
                ) : (
                  <UploadCloud className="w-7 h-7" />
                )}
              </div>

              <div>
                <p className="text-base font-semibold text-slate-800">
                  {loadedFile
                    ? 'Arquivo carregado no painel'
                    : 'Arraste e solte o arquivo Excel (.xls ou .xlsx) aqui'}
                </p>
                <p className="text-xs text-slate-500 mt-1">
                  {loadedFile
                    ? 'Clique para substituir por outro arquivo se desejar'
                    : 'Ou clique para navegar e selecionar do seu computador'}
                </p>
              </div>

              {loadedFile ? (
                <div className="inline-flex items-center gap-2 px-3.5 py-1.5 bg-emerald-100 text-emerald-900 rounded-lg text-xs font-semibold border border-emerald-300">
                  <span>📄 {loadedFile.name}</span>
                  <span className="text-emerald-700">({formatFileSize(loadedFile.size)})</span>
                </div>
              ) : (
                <div className="text-[11px] text-slate-400">
                  Suporta arquivos .xls antigos e .xlsx recentes do SIAFI
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Status Message */}
        {statusMessage && (
          <div
            className={`mt-4 p-4 rounded-lg flex items-start gap-3 text-sm ${
              statusMessage.type === 'success'
                ? 'bg-emerald-50 border border-emerald-200 text-emerald-900'
                : statusMessage.type === 'error'
                ? 'bg-rose-50 border border-rose-200 text-rose-900'
                : 'bg-sky-50 border border-sky-200 text-sky-900'
            }`}
          >
            {statusMessage.type === 'success' ? (
              <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
            ) : statusMessage.type === 'error' ? (
              <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
            ) : (
              <FileCheck2 className="w-5 h-5 text-sky-600 shrink-0 mt-0.5" />
            )}
            <div className="flex-1">
              <p className="font-medium">{statusMessage.text}</p>
              {statusMessage.warnings && statusMessage.warnings.length > 0 && (
                <ul className="mt-2 list-disc list-inside text-xs text-amber-800 space-y-0.5">
                  {statusMessage.warnings.map((w, i) => (
                    <li key={i}>{w}</li>
                  ))}
                </ul>
              )}
            </div>
          </div>
        )}

        {/* PRIMARY ACTION BUTTON: "Gerar Diagramação dos Saldos" */}
        <div className="mt-6 pt-5 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="text-xs text-slate-500">
            {loadedFile
              ? 'Arquivo pronto para análise e estruturação dos cinco quadros contábeis.'
              : 'Faça a carga de um arquivo Excel para habilitar o processamento.'}
          </div>

          <button
            onClick={handleGenerateDiagram}
            disabled={!loadedFile || isProcessing}
            className={`w-full sm:w-auto inline-flex items-center justify-center space-x-2.5 px-6 py-3.5 rounded-xl font-bold text-sm shadow-md transition-all ${
              loadedFile && !isProcessing
                ? 'bg-sky-700 hover:bg-sky-800 text-white cursor-pointer hover:shadow-lg hover:scale-[1.01]'
                : 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
            }`}
          >
            <Play className={`w-4 h-4 ${isProcessing ? 'animate-spin' : ''}`} />
            <span>
              {isProcessing
                ? 'Processando e Diagramando...'
                : 'Gerar Diagramação dos Saldos'}
            </span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Extracted Metadata Inspection Card (Shown after diagramming is generated) */}
      {data && (
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 animate-fadeIn">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Layers className="w-5 h-5 text-sky-700" />
              Identificação dos Metadados Extraídos do Arquivo
            </h3>
            <span className="text-xs font-semibold px-2.5 py-0.5 bg-emerald-100 text-emerald-800 rounded-full">
              Diagramação Ativa
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-sm">
            {/* Órgão */}
            <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 flex items-start gap-3">
              <Building2 className="w-5 h-5 text-slate-500 mt-0.5" />
              <div>
                <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  Órgão / UG
                </span>
                <p className="font-bold text-slate-900 mt-0.5">
                  {data.metadata.orgaoCode} — {data.metadata.orgaoName}
                </p>
              </div>
            </div>

            {/* Período e Exercício */}
            <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 flex items-start gap-3">
              <Calendar className="w-5 h-5 text-slate-500 mt-0.5" />
              <div>
                <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  Exercício & Período
                </span>
                <p className="font-bold text-slate-900 mt-0.5">
                  {data.metadata.exercicio} • {data.metadata.periodo}
                </p>
                <span className="text-xs text-slate-500">
                  Emissão: {data.metadata.dataEmissao}
                </span>
              </div>
            </div>

            {/* Colunas Comparativas */}
            <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 flex items-start gap-3">
              <Layers className="w-5 h-5 text-slate-500 mt-0.5" />
              <div>
                <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  Colunas Comparativas
                </span>
                <p className="font-bold text-slate-900 mt-0.5">
                  {data.metadata.colCurrent} <span className="text-slate-400 font-normal">e</span>{' '}
                  {data.metadata.colPrevious}
                </p>
                <span className="text-xs text-slate-500">
                  Preservadas diretamente da fonte
                </span>
              </div>
            </div>

            {/* Unidade Monetária */}
            <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 flex items-start gap-3">
              <Coins className="w-5 h-5 text-slate-500 mt-0.5" />
              <div className="flex-1">
                <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  Unidade Monetária
                </span>
                <div className="flex items-center justify-between mt-0.5">
                  <span className="font-bold text-slate-900">
                    {data.metadata.unidadeNormalizada}
                  </span>
                  <span
                    className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                      data.metadata.unitConfirmed
                        ? 'bg-emerald-100 text-emerald-800'
                        : 'bg-amber-100 text-amber-800'
                    }`}
                  >
                    {data.metadata.unitConfirmed ? 'Confirmada' : 'Aguardando Confirmação'}
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-1">
                  Original: "{data.metadata.unidadeOriginal}"
                </p>
                <button
                  type="button"
                  onClick={toggleUnitConfirm}
                  className="mt-2 text-xs font-medium text-sky-700 hover:text-sky-900 underline"
                >
                  {data.metadata.unitConfirmed ? 'Alterar / revalidar unidade' : 'Confirmar unidade agora'}
                </button>
              </div>
            </div>

            {/* Aba Selecionada */}
            <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 flex items-start gap-3">
              <FileSpreadsheet className="w-5 h-5 text-slate-500 mt-0.5" />
              <div className="flex-1">
                <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  Aba do Balanço
                </span>
                {data.metadata.availableSheets && data.metadata.availableSheets.length > 1 ? (
                  <select
                    value={data.metadata.sourceSheetName}
                    onChange={e => handleSheetChange(e.target.value)}
                    className="mt-1 block w-full text-xs font-semibold bg-white border border-slate-300 rounded px-2 py-1 text-slate-800"
                  >
                    {data.metadata.availableSheets.map(s => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                ) : (
                  <p className="font-bold text-slate-900 mt-0.5">
                    {data.metadata.sourceSheetName}
                  </p>
                )}
                <span className="text-xs text-slate-500 mt-1 block">
                  Identificação por conteúdo ou nome
                </span>
              </div>
            </div>

            {/* Rastreabilidade e Mapeamento */}
            <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-200 flex items-start gap-3">
              <CheckCircle2 className="w-5 h-5 text-emerald-600 mt-0.5" />
              <div>
                <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  Quadros Diagramados
                </span>
                <p className="font-bold text-emerald-700 mt-0.5">
                  5 Quadros Estruturados
                </p>
                <span className="text-xs text-slate-500">
                  Ativo, Passivo/PL, Financeiro, Compensações e Superávit
                </span>
              </div>
            </div>
          </div>

          <div className="mt-5 pt-4 border-t border-slate-100 flex justify-end">
            <button
              onClick={onGoToDiagram}
              className="inline-flex items-center space-x-2 px-5 py-2.5 rounded-lg bg-sky-700 hover:bg-sky-800 text-white font-semibold text-sm shadow-sm transition-all"
            >
              <span>Ver Balanço Diagramado</span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
