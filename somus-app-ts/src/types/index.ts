// ── User & Auth ────────────────────────────────────────────────────────────────

export type Role = 'admin' | 'corporate' | 'mesa_produtos' | 'seguros';
export type Equipe = 'SP' | 'LEBLON' | 'PRODUTOS' | 'CORPORATE' | 'BACKOFFICE';

// ── Module navigation ──────────────────────────────────────────────────────────

export type ModuleName = 'mesa_produtos' | 'corporate' | 'seguros';
export type PageKey = string;

// ── Financial types ────────────────────────────────────────────────────────────

export interface ConsorcioParams {
  valorCarta: number;
  prazoMeses: number;
  taxaAdm: number;
  fundoReserva: number;
  seguro: number;
  prazoContemp: number;
  parcelaReduzidaPct: number;
  lanceLivrePct: number;
  lanceEmbutidoPct: number;
  correcaoAnual: number;
  tipoCorrecao: 'Pre-fixado' | 'Pos-fixado';
  indiceCorrecao: 'INCC' | 'IPCA' | 'IGP-M' | 'Outro';
}

// ── NASA HD Types ──────────────────────────────────────────────────────────────

export interface NasaConfig {
  seguroBase: 'saldo_devedor' | 'valor_credito';
  momentoAntecipacaoTA: 'junto_1a_parcela' | 'diluido';
  momentoLanceEmbutido: 'na_contemplacao' | 'desde_inicio';
  baseCalculoLanceEmbutido: 'credito_original' | 'original_txadm' | 'atualizado' | 'saldo_devedor';
  baseCalculoLanceLivre: 'credito_original' | 'original_txadm' | 'atualizado' | 'saldo_devedor';
  atualizarValorCredito: boolean;
  baseCalculoFundoReserva: string;
  metodoTIR: 'fluxo_original' | 'fluxo_ajustado';
}

export interface PeriodoDistribuicao {
  start: number;
  end: number;
  fcPct: number;
  taPct: number;
  frPct: number;
}

export interface NasaParams {
  valorCredito: number;
  prazoMeses: number;
  taxaAdmPct: number;
  fundoReservaPct: number;
  periodos: PeriodoDistribuicao[];
  momentoContemplacao: number;
  lanceEmbutidoPct: number;
  lanceLivrePct: number;
  reajustePrePct: number;
  reajustePosPct: number;
  reajustePreFreq: 'Mensal' | 'Bimestral' | 'Trimestral' | 'Semestral' | 'Anual';
  reajustePosFreq: 'Mensal' | 'Bimestral' | 'Trimestral' | 'Semestral' | 'Anual';
  seguroVidaPct: number;
  seguroVidaInicio: number;
  antecipacaoTAPct: number;
  antecipacaoTAParcelas: number;
  taxaVPCredito: number;
  tma: number;
  almAnual: number;
  hurdleAnual: number;
  custosAcessorios: CustoAcessorio[];
}

export interface CustoAcessorio {
  descricao: string;
  valor: number;
  momento: number;
}

export interface FluxoMensal {
  mes: number;
  mesesRestantes: number;
  valorBaseFundoComum: number;
  lanceEmbutido: number;
  lanceLivre: number;
  valorBaseFinal: number;
  pctMensalFC: number;
  pctAcumuladoFC: number;
  amortizacao: number;
  saldoPrincipal: number;
  taxaAdministracao: number;
  fundoReserva: number;
  valorParcela: number;
  saldoDevedor: number;
  pctReajuste: number;
  pctReajusteAcum: number;
  parcelaReajustada: number;
  saldoDevedorReaj: number;
  seguroVida: number;
  parcelaComSeguro: number;
  outrosCustos: number;
  cartaCredito: number;
  cartaCreditoReaj: number;
  fluxoCaixa: number;
  fluxoCaixaTIR: number;
}

export interface FluxoResult {
  fluxoMensal: FluxoMensal[];
  cashflow: number[];
  totalPago: number;
  cartaLiquida: number;
  lanceLivreValor: number;
  lanceEmbutidoValor: number;
  parcelaF1Base: number;
  parcelaF2Base: number;
  tirMensal: number;
  tirAnual: number;
  cetAnual: number;
}

export interface VPLResult {
  b0: number;
  h0: number;
  d0: number;
  pvPosT: number;
  deltaVPL: number;
  criaValor: boolean;
  breakEvenLance: number;
  tirMensal: number;
  tirAnual: number;
  cetAnual: number;
  vplTotal: number;
}

// ── Financiamento ──────────────────────────────────────────────────────────────

export interface FinanciamentoParams {
  valor: number;
  prazoMeses: number;
  taxaMensalPct: number;
  metodo: 'price' | 'sac';
  carenciaMeses?: number;
  iof?: boolean;
  custosTAC?: number;
  custosAvaliacao?: number;
}

export interface ParcelaFinanc {
  mes: number;
  parcela: number;
  juros: number;
  amortizacao: number;
  saldo: number;
}

export interface FinanciamentoResult {
  parcelas: ParcelaFinanc[];
  cashflow: number[];
  totalPago: number;
  totalJuros: number;
  valor: number;
  tirMensal?: number;
  tirAnual?: number;
}

// ── Comparativo ────────────────────────────────────────────────────────────────

export interface ComparativoResult {
  consorcio: FluxoResult;
  financiamento: FinanciamentoResult;
  vplConsorcio: number;
  vplFinanciamento: number;
  economiaVPL: number;
  tirConsorcioMensal: number;
  tirConsorcioAnual: number;
  tirFinancMensal: number;
  tirFinancAnual: number;
}

// ── Venda de Operacao ──────────────────────────────────────────────────────────

export interface VendaOperacaoParams {
  momentoVenda: number;
  valorVenda: number;
  tma: number;
}

export interface VendaResult {
  vpl: number;
  ganhoPct: number;
  prazoMedio: number;
  ganhoMes: number;
  margemMes: number;
  custoComprador: number;
}

// ── Scenarios ──────────────────────────────────────────────────────────────────

export interface Scenario {
  id: number;
  nome: string;
  params: NasaParams;
  resultado: FluxoResult;
  vpl?: VPLResult;
  timestamp: string;
}

// ── PPTX Generation types ──────────────────────────────────────────────────────

export type PropostaSubtipo = 'CETHD' | 'CETHD_Lance_Fixo' | 'Lance_Fixo' | 'PF_Tradicional' | 'PJ_Tradicional';

export interface CenarioComparativa {
  contemplacaoMes: number;
  cetMensal: number;
  cetAnual: number;
  lanceLivrePct: number;
  lanceLivreValor: number;
  creditoDisponivel: number;
  alavancagem: number;
  valorCarta: number;
  lanceEmbutidoPct: number;
  lanceEmbutidoValor: number;
  taxaAdm: number;
  fundoReserva: number;
  prazoTotal: number;
  correcaoAnual: number;
  parcelas: number[];
}

export interface ComparativaData {
  clienteNome: string;
  dataMesAno: string;
  cenarios: CenarioComparativa[];
  contato1Nome: string;
  contato1Telefone: string;
  contato2Nome: string;
  contato2Telefone: string;
}

export interface PropostaData {
  clienteNome: string;
  dataMesAno: string;
  empresaNome: string;
  localizacao: string;
  descricaoOperacao: string;
  objetivoTitulo: string;
  horizonteTexto: string;
  objetivoDescricao: string;
  administradora: string;
  cenario: CenarioComparativa & {
    lanceProprioPct: number;
    lanceProprioValor: number;
    creditoEfetivo: number;
    estruturaDescricao: string;
    cronogramaResumo: string;
    captacaoMeses: number;
    totalLances: number;
  };
  contato1Nome: string;
  contato1Telefone: string;
  contato2Nome: string;
  contato2Telefone: string;
}

// ── Mesa Produtos types ────────────────────────────────────────────────────────

export interface Assessor {
  codigo: string;
  nome: string;
  equipe: Equipe;
  email: string;
  telefone?: string;
}

export interface ReceitaItem {
  assessor: string;
  produto: string;
  valor: number;
  data: string;
  equipe: Equipe;
}

export interface Operacao {
  id: string;
  clienteNome: string;
  assessor: string;
  tipo: string;
  valorCarta: number;
  status: 'ativa' | 'contemplada' | 'encerrada';
  dataInicio: string;
  pagamentos: PagamentoOperacao[];
}

export interface PagamentoOperacao {
  mes: number;
  valor: number;
  data: string;
  status: 'pago' | 'pendente' | 'atrasado';
}

// ── Seguros ────────────────────────────────────────────────────────────────────

export interface Renovacao {
  cliente: string;
  assessor: string;
  seguradora: string;
  produto: string;
  vencimento: string;
  premio: number;
}

// ── App State ──────────────────────────────────────────────────────────────────

export interface AppState {
  currentModule: ModuleName;
  currentPage: PageKey;
  role: Role;
  userName: string;
  sidebarCollapsed: boolean;
}
