import React, { useState, useCallback } from 'react';
import { Card } from '@components/Card';
import { Button } from '@components/Button';
import {
  Upload,
  Download,
  Layers,
  FileSpreadsheet,
  ArrowRight,
  Trash2,
  Check,
  Settings,
} from 'lucide-react';
import { cn } from '@/lib/utils';
import { readWorkbookFromFile, writeExcel, type WorkbookData } from '@/services/excel-reader';

// ── Types ────────────────────────────────────────────────────────────────────

interface UploadedFile {
  id: string;
  name: string;
  sheets: string[];
  data: WorkbookData;
  selectedSheet: string;
}

interface ColumnMapping {
  source: string;
  target: string;
  enabled: boolean;
}

// ── Component ────────────────────────────────────────────────────────────────

export default function Organizador() {
  const [files, setFiles] = useState<UploadedFile[]>([]);
  const [uploading, setUploading] = useState(false);
  const [columnMappings, setColumnMappings] = useState<ColumnMapping[]>([]);
  const [targetColumns, setTargetColumns] = useState<string[]>([]);
  const [mergedData, setMergedData] = useState<any[][] | null>(null);
  const [step, setStep] = useState<'upload' | 'mapping' | 'result'>('upload');

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList) return;
    setUploading(true);
    try {
      const newFiles: UploadedFile[] = [];
      for (let i = 0; i < fileList.length; i++) {
        const file = fileList[i];
        const wb = await readWorkbookFromFile(file);
        newFiles.push({
          id: `${Date.now()}-${i}`,
          name: file.name,
          sheets: wb.sheetNames,
          data: wb,
          selectedSheet: wb.sheetNames[0],
        });
      }
      setFiles((prev) => [...prev, ...newFiles]);
    } catch (err) {
      console.error('Erro ao ler arquivos:', err);
    } finally {
      setUploading(false);
    }
  }, []);

  const removeFile = (id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  };

  const changeSheet = (id: string, sheet: string) => {
    setFiles((prev) =>
      prev.map((f) => (f.id === id ? { ...f, selectedSheet: sheet } : f))
    );
  };

  const handlePrepareMapping = useCallback(() => {
    // Collect all unique columns from all selected sheets
    const allCols = new Set<string>();
    files.forEach((f) => {
      const sheetData = f.data.sheets[f.selectedSheet];
      if (sheetData && sheetData[0]) {
        sheetData[0].forEach((col: any) => {
          if (col) allCols.add(String(col));
        });
      }
    });

    const cols = Array.from(allCols);
    setTargetColumns(cols);
    setColumnMappings(
      cols.map((c) => ({
        source: c,
        target: c,
        enabled: true,
      }))
    );
    setStep('mapping');
  }, [files]);

  const toggleMapping = (index: number) => {
    setColumnMappings((prev) =>
      prev.map((m, i) => (i === index ? { ...m, enabled: !m.enabled } : m))
    );
  };

  const updateMappingTarget = (index: number, target: string) => {
    setColumnMappings((prev) =>
      prev.map((m, i) => (i === index ? { ...m, target } : m))
    );
  };

  const handleMerge = useCallback(() => {
    const enabledMappings = columnMappings.filter((m) => m.enabled);
    const headers = enabledMappings.map((m) => m.target);
    const result: any[][] = [headers];

    files.forEach((f) => {
      const sheetData = f.data.sheets[f.selectedSheet];
      if (!sheetData || sheetData.length < 2) return;

      const sourceHeaders = sheetData[0].map(String);

      for (let rowIdx = 1; rowIdx < sheetData.length; rowIdx++) {
        const row: any[] = [];
        enabledMappings.forEach((mapping) => {
          const colIdx = sourceHeaders.indexOf(mapping.source);
          row.push(colIdx >= 0 ? sheetData[rowIdx][colIdx] : '');
        });
        result.push(row);
      }
    });

    setMergedData(result);
    setStep('result');
  }, [files, columnMappings]);

  const handleExport = useCallback(() => {
    if (!mergedData) return;
    writeExcel(mergedData, 'Dados_Organizados.xlsx', 'Organizado');
  }, [mergedData]);

  const handleReset = () => {
    setFiles([]);
    setColumnMappings([]);
    setTargetColumns([]);
    setMergedData(null);
    setStep('upload');
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-somus-text-primary">
            Organizador de Planilhas
          </h1>
          <p className="text-sm text-somus-text-secondary mt-1">
            Combine e organize multiplas planilhas Excel
          </p>
        </div>
        {step !== 'upload' && (
          <Button variant="ghost" onClick={handleReset}>
            Recomecar
          </Button>
        )}
      </div>

      {/* Steps indicator */}
      <div className="flex items-center gap-4">
        {['Upload', 'Mapeamento', 'Resultado'].map((label, i) => {
          const stepKeys = ['upload', 'mapping', 'result'];
          const isActive = stepKeys.indexOf(step) >= i;
          return (
            <React.Fragment key={label}>
              <div className="flex items-center gap-2">
                <div
                  className={cn(
                    'w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold',
                    isActive
                      ? 'bg-somus-green text-white'
                      : 'bg-somus-border text-somus-text-secondary'
                  )}
                >
                  {i + 1}
                </div>
                <span
                  className={cn(
                    'text-sm font-medium',
                    isActive ? 'text-somus-green' : 'text-somus-text-tertiary'
                  )}
                >
                  {label}
                </span>
              </div>
              {i < 2 && (
                <ArrowRight className="h-4 w-4 text-somus-text-tertiary" />
              )}
            </React.Fragment>
          );
        })}
      </div>

      {/* Step: Upload */}
      {step === 'upload' && (
        <>
          <Card title="Arquivos">
            <div className="mt-4 space-y-4">
              <label className="cursor-pointer block">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  onChange={handleUpload}
                  className="hidden"
                />
                <div className="border-2 border-dashed border-somus-border rounded-lg p-8 text-center hover:border-somus-green/50 transition-colors">
                  <Upload className="h-8 w-8 text-somus-text-tertiary mx-auto mb-3" />
                  <p className="text-sm text-somus-text-secondary">
                    {uploading ? 'Carregando...' : 'Clique para selecionar arquivos Excel'}
                  </p>
                  <p className="text-xs text-somus-text-tertiary mt-1">
                    Suporta multiplos arquivos .xlsx e .xls
                  </p>
                </div>
              </label>

              {files.length > 0 && (
                <div className="space-y-2">
                  {files.map((f) => (
                    <div
                      key={f.id}
                      className="flex items-center justify-between py-3 px-4 bg-somus-bg-hover rounded-lg"
                    >
                      <div className="flex items-center gap-3">
                        <FileSpreadsheet className="h-5 w-5 text-somus-green" />
                        <div>
                          <div className="text-sm font-medium text-somus-text-primary">{f.name}</div>
                          <div className="text-xs text-somus-text-tertiary">
                            {f.sheets.length} sheet(s)
                          </div>
                        </div>
                      </div>
                      <div className="flex items-center gap-3">
                        <select
                          value={f.selectedSheet}
                          onChange={(e) => changeSheet(f.id, e.target.value)}
                          className="text-xs border border-somus-border rounded px-2 py-1"
                        >
                          {f.sheets.map((s) => (
                            <option key={s} value={s}>{s}</option>
                          ))}
                        </select>
                        <button
                          onClick={() => removeFile(f.id)}
                          className="text-somus-text-tertiary hover:text-red-400 transition-colors"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {files.length > 0 && (
                <div className="flex justify-end">
                  <Button
                    variant="primary"
                    icon={<Settings className="h-4 w-4" />}
                    onClick={handlePrepareMapping}
                  >
                    Configurar Mapeamento
                  </Button>
                </div>
              )}
            </div>
          </Card>
        </>
      )}

      {/* Step: Mapping */}
      {step === 'mapping' && (
        <Card title="Mapeamento de Colunas">
          <div className="mt-4 space-y-3">
            {columnMappings.map((m, i) => (
              <div
                key={i}
                className={cn(
                  'flex items-center gap-4 py-2 px-4 rounded-lg',
                  m.enabled ? 'bg-somus-bg-hover' : 'bg-somus-border/30 opacity-60'
                )}
              >
                <input
                  type="checkbox"
                  checked={m.enabled}
                  onChange={() => toggleMapping(i)}
                  className="w-4 h-4 rounded border-somus-border text-somus-green focus:ring-somus-green/40"
                />
                <div className="flex-1">
                  <span className="text-sm text-somus-text-secondary">Origem:</span>
                  <span className="text-sm font-medium text-somus-text-primary ml-2">{m.source}</span>
                </div>
                <ArrowRight className="h-4 w-4 text-somus-text-tertiary" />
                <div className="flex-1">
                  <span className="text-sm text-somus-text-secondary">Destino:</span>
                  <input
                    type="text"
                    value={m.target}
                    onChange={(e) => updateMappingTarget(i, e.target.value)}
                    disabled={!m.enabled}
                    className="text-sm border border-somus-border rounded px-2 py-1 ml-2 w-40 focus:ring-2 focus:ring-somus-green/40 focus:outline-none disabled:opacity-50"
                  />
                </div>
              </div>
            ))}

            <div className="flex justify-end gap-3 pt-4">
              <Button variant="secondary" onClick={() => setStep('upload')}>
                Voltar
              </Button>
              <Button
                variant="primary"
                icon={<Layers className="h-4 w-4" />}
                onClick={handleMerge}
              >
                Organizar
              </Button>
            </div>
          </div>
        </Card>
      )}

      {/* Step: Result */}
      {step === 'result' && mergedData && (
        <Card
          title={`Resultado (${mergedData.length - 1} linhas)`}
          headerRight={
            <Button
              variant="primary"
              size="sm"
              icon={<Download className="h-4 w-4" />}
              onClick={handleExport}
            >
              Exportar Excel
            </Button>
          }
        >
          <div className="mt-4 overflow-x-auto max-h-[500px] overflow-y-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-somus-bg-secondary">
                <tr className="border-b border-somus-border">
                  {mergedData[0].map((header: string, i: number) => (
                    <th key={i} className="text-left py-3 px-3 font-semibold text-somus-text-secondary whitespace-nowrap">
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {mergedData.slice(1, 101).map((row, rowIdx) => (
                  <tr
                    key={rowIdx}
                    className="border-b border-somus-border/30 hover:bg-somus-bg-hover"
                  >
                    {row.map((cell: any, colIdx: number) => (
                      <td key={colIdx} className="py-2 px-3 text-somus-text-primary whitespace-nowrap">
                        {cell ?? ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
            {mergedData.length > 101 && (
              <div className="text-center py-4 text-sm text-somus-text-tertiary">
                Exibindo 100 de {mergedData.length - 1} linhas
              </div>
            )}
          </div>
        </Card>
      )}
    </div>
  );
}
