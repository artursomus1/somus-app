/**
 * Brazilian formatting utilities for currency, percentages, dates, and numbers.
 */

const MESES_PT = [
  'Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
];

/**
 * Format a number as Brazilian Real currency.
 * fmtCurrency(1234567.89) => "R$ 1.234.567,89"
 */
export function fmtCurrency(value: number): string {
  return value.toLocaleString('pt-BR', {
    style: 'currency',
    currency: 'BRL',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/**
 * Format a number as abbreviated Brazilian Real currency.
 * fmtCurrencyShort(1234567) => "R$ 1,2M"
 * fmtCurrencyShort(500000)  => "R$ 500K"
 * fmtCurrencyShort(1500)    => "R$ 1.500,00"
 */
export function fmtCurrencyShort(value: number): string {
  const abs = Math.abs(value);
  const sign = value < 0 ? '-' : '';

  if (abs >= 1_000_000_000) {
    const v = abs / 1_000_000_000;
    return `${sign}R$ ${v.toLocaleString('pt-BR', { maximumFractionDigits: 1 })}B`;
  }
  if (abs >= 1_000_000) {
    const v = abs / 1_000_000;
    return `${sign}R$ ${v.toLocaleString('pt-BR', { maximumFractionDigits: 1 })}M`;
  }
  if (abs >= 10_000) {
    const v = abs / 1_000;
    return `${sign}R$ ${v.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}K`;
  }
  return fmtCurrency(value);
}

/**
 * Format a number as a percentage with Brazilian locale.
 * fmtPct(12.5)    => "12,50%"
 * fmtPct(12.5, 1) => "12,5%"
 */
export function fmtPct(value: number, decimals: number = 2): string {
  return `${value.toLocaleString('pt-BR', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  })}%`;
}

/**
 * Format a monthly rate.
 * fmtPctAm(1.5) => "1,50% a.m."
 */
export function fmtPctAm(value: number): string {
  return `${fmtPct(value)} a.m.`;
}

/**
 * Format an annual rate.
 * fmtPctAa(19.56) => "19,56% a.a."
 */
export function fmtPctAa(value: number): string {
  return `${fmtPct(value)} a.a.`;
}

/**
 * Format a month ordinal.
 * fmtMes(60) => "60o mes"
 */
export function fmtMes(month: number): string {
  return `${month}o mes`;
}

/**
 * Format a Date as "Mes Ano" in Portuguese.
 * fmtMesAno(new Date(2026, 2, 24)) => "Marco 2026"
 */
export function fmtMesAno(date: Date): string {
  return `${MESES_PT[date.getMonth()]} ${date.getFullYear()}`;
}

/**
 * Format a Date as "DD/MM/YYYY".
 * fmtDate(new Date(2026, 2, 24)) => "24/03/2026"
 */
export function fmtDate(date: Date): string {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

/**
 * Parse a Brazilian-formatted number string to a JavaScript number.
 * parseBRNumber("1.234,56") => 1234.56
 * parseBRNumber("1234")     => 1234
 */
export function parseBRNumber(str: string): number {
  if (!str || typeof str !== 'string') return 0;
  const cleaned = str
    .replace(/[^\d.,-]/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const result = parseFloat(cleaned);
  return isNaN(result) ? 0 : result;
}

/**
 * Format a month count.
 * fmtMeses(225) => "225 meses"
 * fmtMeses(1)   => "1 mes"
 */
export function fmtMeses(n: number): string {
  return n === 1 ? '1 mes' : `${n} meses`;
}

/**
 * Format a month count as years and months.
 * fmtHorizonte(225) => "18 anos e 9 meses"
 * fmtHorizonte(12)  => "1 ano"
 * fmtHorizonte(5)   => "5 meses"
 */
export function fmtHorizonte(meses: number): string {
  const anos = Math.floor(meses / 12);
  const rest = meses % 12;

  if (anos === 0) {
    return fmtMeses(rest);
  }
  if (rest === 0) {
    return anos === 1 ? '1 ano' : `${anos} anos`;
  }

  const anoStr = anos === 1 ? '1 ano' : `${anos} anos`;
  const mesStr = rest === 1 ? '1 mes' : `${rest} meses`;
  return `${anoStr} e ${mesStr}`;
}
