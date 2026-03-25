import React, { useState, useMemo } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import {
  Calculator,
  Landmark,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import { useAppStore } from '@/stores/appStore';
import { calcularCreditoLance, calcularCustoCombinado, NasaEngine } from '@engine/index';
import type { FinanciamentoResult } from '@engine/financiamento';
import type { CustoCombinado } from '@engine/comparativo';
import { loadData, saveData } from '@/services/storage';

// ── Format helpers ──────────────────────────────────────────────────────────

function fmtBRL(v: number): string {
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(v: number, d = 2): string {
  return `${v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d })}%`;
}

// ── Schema ──────────────────────────────────────────────────────────────────

const schema = z.object({
  valorEmprestimo: z.number().min(1),
  prazo: z.number().min(1).max(360),
  taxaMensal: z.number().min(0),
  metodo: z.enum(['price', 'sac']),
  carencia: z.number().min(0),
  tac: z.number().min(0),
  avaliacaoGarantia: z.number().min(0),
  comissao: z.number().min(0),
  pagAntecipadoMes: z.number().min(0),
  pagAntecipadoValor: z.number().min(0),
});

type FormData = z.infer<typeof schema>;

const defaults: FormData = {
  valorEmprestimo: 100000,
  prazo: 60,
  taxaMensal: 1.5,
  metodo: 'price',
  carencia: 0,
  tac: 0,
  avaliacaoGarantia: 0,
  comissao: 0,
  pagAntecipadoMes: 0,
  pagAntecipadoValor: 0,
};

const inputCls = 'w-full px-2.5 py-1.5 text-sm bg-somus-bg-input border border-somus-border rounded-md text-somus-text-primary focus:ring-1 focus:ring-somus-green/50 focus:border-somus-green outline-none';

function DInput({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-[10px] font-medium text-somus-text-secondary uppercase tracking-wider mb-1">{label}</label>
      {children}
    </div>
  );
}

function Section({ title, tag, children, defaultOpen = true }: {
  title: string; tag: string; children: React.ReactNode; defaultOpen?: boolean;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
      <button type="button" onClick={() => setOpen(!open)}
        className="w-full flex items-center justify-between px-4 py-3 hover:bg-somus-bg-hover transition-colors">
        <div className="flex items-center gap-2">
          <span className="text-[10px] font-bold bg-somus-green/20 text-somus-green px-1.5 py-0.5 rounded">{tag}</span>
          <span className="text-sm font-semibold text-somus-text-primary">{title}</span>
        </div>
        {open ? <ChevronUp size={16} className="text-somus-text-tertiary" /> : <ChevronDown size={16} className="text-somus-text-tertiary" />}
      </button>
      {open && <div className="px-4 pb-4 pt-2">{children}</div>}
    </div>
  );
}

// ── Main Component ──────────────────────────────────────────────────────────

export default function CreditoLance() {
  const [result, setResult] = useState<FinanciamentoResult | null>(null);
  const [combinado, setCombinado] = useState<CustoCombinado | null>(null);
  const [loading, setLoading] = useState(false);

  const { control, handleSubmit } = useForm<FormData>({
    resolver: zodResolver(schema),
    defaultValues: defaults,
  });

  function onCalculate(data: FormData) {
    setLoading(true);
    try {
      const pagAntecipado = data.pagAntecipadoMes > 0 && data.pagAntecipadoValor > 0
        ? { mes: data.pagAntecipadoMes, valor: data.pagAntecipadoValor }
        : undefined;

      const resultado = calcularCreditoLance({
        valor: data.valorEmprestimo,
        prazo: data.prazo,
        taxa: data.taxaMensal,
        metodo: data.metodo,
        carencia: data.carencia,
        tac: data.tac,
        avaliacaoGarantia: data.avaliacaoGarantia,
        comissao: data.comissao,
        pagamentoAntecipado: pagAntecipado,
      });

      setResult(resultado);

      // Save to localStorage for CustoCombinado page
      saveData('somus-credito-lance-result', resultado);

      // Try to compute combined cost if consortium data exists
      try {
        const consorcioData = loadData<any>('somus-simulador-fluxo', null);
        if (consorcioData && consorcioData.cashflow) {
          const comb = calcularCustoCombinado(consorcioData, resultado);
          setCombinado(comb);
        } else {
          setCombinado(null);
        }
      } catch {
        setCombinado(null);
      }
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="flex flex-col h-full w-full overflow-hidden">
      <header className="sticky top-0 z-20 bg-somus-bg-primary/90 backdrop-blur-md border-b border-somus-border px-6 py-3">
        <div className="flex items-center gap-3">
          <Landmark size={20} className="text-somus-purple" />
          <div>
            <h1 className="text-lg font-semibold text-somus-text-primary">Crédito para Lance</h1>
            <p className="text-xs text-somus-text-tertiary">Op. Crédito para Lance - espelha aba "Op. Crédito para Lance" da NASA HD</p>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto bg-somus-bg-primary p-5 space-y-5">
        <form onSubmit={handleSubmit(onCalculate)} className="space-y-4">
          {/* Main params */}
          <Section title="Parâmetros do Empréstimo" tag="A">
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <DInput label="Valor do Empréstimo (R$)"><Controller name="valorEmprestimo" control={control} render={({ field }) => (
                <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Prazo (meses)"><Controller name="prazo" control={control} render={({ field }) => (
                <input type="number" min={1} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Taxa Mensal (%)"><Controller name="taxaMensal" control={control} render={({ field }) => (
                <input type="number" step={0.01} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Método"><Controller name="metodo" control={control} render={({ field }) => (
                <select value={field.value} onChange={(e) => field.onChange(e.target.value)} className={inputCls}>
                  <option value="price">Price</option>
                  <option value="sac">SAC</option>
                </select>
              )} /></DInput>
            </div>
          </Section>

          {/* Costs */}
          <Section title="Custos Adicionais" tag="B" defaultOpen={false}>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <DInput label="Carência (meses)"><Controller name="carencia" control={control} render={({ field }) => (
                <input type="number" min={0} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="TAC (R$)"><Controller name="tac" control={control} render={({ field }) => (
                <input type="number" step={100} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Avaliação Garantia (R$)"><Controller name="avaliacaoGarantia" control={control} render={({ field }) => (
                <input type="number" step={100} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Comissão (R$)"><Controller name="comissao" control={control} render={({ field }) => (
                <input type="number" step={100} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
            </div>
          </Section>

          {/* Early payment */}
          <Section title="Pagamento Antecipado (Opcional)" tag="C" defaultOpen={false}>
            <div className="grid grid-cols-2 gap-3">
              <DInput label="Mês Antecipação"><Controller name="pagAntecipadoMes" control={control} render={({ field }) => (
                <input type="number" min={0} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
              <DInput label="Valor Antecipação (R$)"><Controller name="pagAntecipadoValor" control={control} render={({ field }) => (
                <input type="number" step={1000} value={field.value} onChange={(e) => field.onChange(Number(e.target.value))} className={inputCls} />
              )} /></DInput>
            </div>
          </Section>

          <button type="submit" disabled={loading}
            className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-semibold bg-somus-green text-white rounded-lg hover:bg-somus-green-light transition-colors disabled:opacity-50">
            <Calculator size={16} /> {loading ? 'Calculando...' : 'Calcular Crédito Lance'}
          </button>
        </form>

        {result && (
          <div className="space-y-5">
            {/* KPI Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Total Pago</span>
                <p className="text-sm font-bold text-somus-text-primary mt-1">{fmtBRL(result.total_pago)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">Total Juros</span>
                <p className="text-sm font-bold text-red-400 mt-1">{fmtBRL(result.total_juros)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">IOF</span>
                <p className="text-sm font-bold text-somus-text-primary mt-1">{fmtBRL(result.iof)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">CET Anual</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(result.cet_anual * 100, 2)}</p>
              </div>
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-3">
                <span className="text-[10px] text-somus-text-secondary uppercase">TIR Mensal</span>
                <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(result.tir_mensal * 100, 4)}</p>
              </div>
            </div>

            {/* Amortization Table */}
            <div className="bg-somus-bg-secondary border border-somus-border rounded-lg overflow-hidden">
              <div className="px-4 py-3 border-b border-somus-border flex items-center justify-between">
                <h3 className="text-sm font-semibold text-somus-text-primary">Tabela de Amortização</h3>
                <span className="text-[10px] text-somus-text-tertiary">{result.parcelas.length} parcelas</span>
              </div>
              <div className="overflow-x-auto max-h-[400px]">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-somus-bg-tertiary">
                    <tr>
                      <th className="px-3 py-2 text-left text-somus-text-secondary font-medium">Mês</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Parcela</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Juros</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Amortização</th>
                      <th className="px-3 py-2 text-right text-somus-text-secondary font-medium">Saldo</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.parcelas.map((r) => (
                      <tr key={r.mes} className="border-b border-somus-border/30">
                        <td className="px-3 py-1.5 text-somus-text-primary font-medium">{r.mes}</td>
                        <td className="px-3 py-1.5 text-right text-somus-text-primary">{fmtBRL(r.parcela)}</td>
                        <td className="px-3 py-1.5 text-right text-red-400">{fmtBRL(r.juros)}</td>
                        <td className="px-3 py-1.5 text-right text-somus-green">{fmtBRL(r.amortizacao)}</td>
                        <td className="px-3 py-1.5 text-right text-somus-text-secondary">{fmtBRL(r.saldo)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Combined Cost Section */}
            {combinado && (
              <div className="bg-somus-bg-secondary border border-somus-border rounded-lg p-5">
                <h3 className="text-sm font-semibold text-somus-text-primary mb-4">Custo Combinado (Consórcio + Lance)</h3>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  <div className="bg-somus-bg-tertiary rounded-md p-3">
                    <span className="text-[10px] text-somus-text-secondary uppercase">Total Combinado</span>
                    <p className="text-sm font-bold text-somus-text-primary mt-1">{fmtBRL(combinado.totalPago)}</p>
                  </div>
                  <div className="bg-somus-bg-tertiary rounded-md p-3">
                    <span className="text-[10px] text-somus-text-secondary uppercase">TIR Mensal Combinada</span>
                    <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(combinado.tirMensal * 100, 4)}</p>
                  </div>
                  <div className="bg-somus-bg-tertiary rounded-md p-3">
                    <span className="text-[10px] text-somus-text-secondary uppercase">TIR Anual Combinada</span>
                    <p className="text-sm font-bold text-somus-gold mt-1">{fmtPct(combinado.tirAnual * 100, 2)}</p>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
