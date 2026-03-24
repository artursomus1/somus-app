"""
NASA Engine HD - Motor de Calculo Completo VPL para Consorcios
Substituicao completa da planilha NASA NOVA HD VPL (21 abas, 150+ named ranges).

Somus Capital - Mesa de Produtos
Calcula: fluxo financeiro completo (35 colunas), distribuicao multi-periodo,
reajuste de parcelas, seguro, custos acessorios, VPL HD com taxas duais,
venda de operacao, credito para lance, custo combinado, financiamento
comparativo, credito equivalente, cenarios, consolidacao de cotas.

Compativel com a API do nasa_engine.py original (funcoes standalone mantidas).
"""

import math
import copy
from dataclasses import dataclass, field
from typing import Optional


# =============================================================================
# CONSTANTES
# =============================================================================

FREQ_MAP = {
    "Mensal": 1,
    "Bimestral": 2,
    "Trimestral": 3,
    "Semestral": 6,
    "Anual": 12,
}

IOF_DIARIO = 0.000082  # 0.0082% a.d. (IOF padrao)
IOF_ADICIONAL = 0.0038  # 0.38% sobre principal (IOF adicional)

MAX_MESES = 420  # 35 anos


# =============================================================================
# HELPERS MATEMATICOS (compatibilidade com nasa_engine.py)
# =============================================================================

def _npv(rate_monthly, cashflows):
    """Valor Presente Liquido de uma serie de fluxos mensais."""
    if rate_monthly == 0:
        return sum(cashflows)
    total = 0.0
    for t, cf in enumerate(cashflows):
        try:
            total += cf / (1 + rate_monthly) ** t
        except (OverflowError, ZeroDivisionError):
            break
    return total


def _irr(cashflows, guess=0.01, tol=1e-9, max_iter=500):
    """Taxa Interna de Retorno (mensal) via Newton-Raphson com fallback bisection."""
    # Tenta Newton-Raphson primeiro
    rate = guess
    for _ in range(max_iter):
        npv_val = 0.0
        dnpv = 0.0
        ok = True
        for t, cf in enumerate(cashflows):
            try:
                d = (1 + rate) ** t
                if abs(d) < 1e-30:
                    ok = False
                    break
                npv_val += cf / d
                if t > 0:
                    dnpv -= t * cf / ((1 + rate) ** (t + 1))
            except (OverflowError, ZeroDivisionError):
                ok = False
                break
        if not ok:
            break
        if abs(npv_val) < tol:
            return rate
        if abs(dnpv) < 1e-15:
            break
        step = npv_val / dnpv
        rate -= step
        if rate <= -1:
            rate = 0.0001

    # Fallback: bisection
    return _irr_bisection(cashflows, tol=tol, max_iter=max_iter)


def _irr_bisection(cashflows, lo=-0.5, hi=2.0, tol=1e-9, max_iter=1000):
    """IRR via bisection (mais robusto para fluxos dificeis)."""
    # Usa limites conservadores para evitar overflow com fluxos longos
    if hi > 1.0 and len(cashflows) > 100:
        hi = 0.5
    if lo < -0.9 and len(cashflows) > 100:
        lo = -0.3

    npv_lo = _npv(lo, cashflows)
    npv_hi = _npv(hi, cashflows)

    # Procura limites validos
    if npv_lo * npv_hi > 0:
        for test_lo in [-0.3, -0.1, -0.01, 0.0]:
            for test_hi in [0.05, 0.1, 0.2, 0.5]:
                n_lo = _npv(test_lo, cashflows)
                n_hi = _npv(test_hi, cashflows)
                if n_lo * n_hi < 0:
                    lo, hi = test_lo, test_hi
                    npv_lo, npv_hi = n_lo, n_hi
                    break
            else:
                continue
            break
        else:
            return 0.0  # Sem solucao

    for _ in range(max_iter):
        mid = (lo + hi) / 2
        npv_mid = _npv(mid, cashflows)
        if abs(npv_mid) < tol or (hi - lo) / 2 < tol:
            return mid
        if npv_mid * npv_lo < 0:
            hi = mid
            npv_hi = npv_mid
        else:
            lo = mid
            npv_lo = npv_mid
    return (lo + hi) / 2


def _annual_from_monthly(r_m):
    """Converte taxa mensal em anual equivalente."""
    return (1 + r_m) ** 12 - 1


def _monthly_from_annual(r_a):
    """Converte taxa anual em mensal equivalente."""
    if r_a <= -1:
        return 0.0
    return (1 + r_a) ** (1 / 12) - 1


def _pmt(rate, nper, pv):
    """Calculo PMT (parcela Price)."""
    if rate == 0:
        return pv / nper if nper > 0 else 0.0
    return pv * rate * (1 + rate) ** nper / ((1 + rate) ** nper - 1)


# =============================================================================
# GOAL SEEK - Substitui o GoalSeek do Excel
# =============================================================================

def goal_seek(target_func, target_value, initial_guess,
              tolerance=1e-9, max_iter=500, lo=None, hi=None):
    """
    Encontra x tal que target_func(x) ≈ target_value.
    Usa bisection quando limites sao fornecidos, Newton numerico caso contrario.

    Args:
        target_func: funcao f(x) -> float
        target_value: valor alvo
        initial_guess: palpite inicial
        tolerance: precisao
        max_iter: iteracoes maximas
        lo, hi: limites para bisection (opcional)

    Returns:
        x encontrado (float)
    """
    func = lambda x: target_func(x) - target_value

    # Bisection se limites fornecidos
    if lo is not None and hi is not None:
        f_lo = func(lo)
        f_hi = func(hi)
        if f_lo * f_hi > 0:
            # Tenta expandir limites
            for _ in range(20):
                lo *= 0.5
                hi *= 2.0
                f_lo = func(lo)
                f_hi = func(hi)
                if f_lo * f_hi <= 0:
                    break
            else:
                # Fallback para Newton
                return _goal_seek_newton(func, initial_guess, tolerance, max_iter)

        for _ in range(max_iter):
            mid = (lo + hi) / 2
            f_mid = func(mid)
            if abs(f_mid) < tolerance or (hi - lo) / 2 < tolerance:
                return mid
            if f_mid * f_lo < 0:
                hi = mid
                f_hi = f_mid
            else:
                lo = mid
                f_lo = f_mid
        return (lo + hi) / 2

    return _goal_seek_newton(func, initial_guess, tolerance, max_iter)


def _goal_seek_newton(func, guess, tolerance, max_iter):
    """Newton numerico com derivada por diferencas finitas."""
    x = guess
    h = max(abs(x) * 1e-6, 1e-10)
    for _ in range(max_iter):
        f_x = func(x)
        if abs(f_x) < tolerance:
            return x
        f_xh = func(x + h)
        deriv = (f_xh - f_x) / h
        if abs(deriv) < 1e-15:
            h *= 10
            continue
        step = f_x / deriv
        x -= step
        h = max(abs(x) * 1e-6, 1e-10)
    return x


# =============================================================================
# CONFIGURACAO (Parametros sheet)
# =============================================================================

@dataclass
class NasaConfig:
    """Configuracoes da operacao (aba Parametros da planilha)."""

    # Base de calculo do seguro
    seguro_base: str = "saldo_devedor"  # "saldo_devedor" ou "valor_credito"

    # Momento da antecipacao da taxa de administracao
    momento_antecipacao_ta: str = "junto_1a_parcela"  # "junto_1a_parcela", "na_contemplacao", "diluida"

    # Momento do lance embutido
    momento_lance_embutido: str = "na_contemplacao"  # "na_contemplacao" ou "desde_inicio"

    # Base de calculo do lance embutido
    base_calculo_lance_embutido: str = "credito_original"
    # Opcoes: "credito_original", "original+txadm", "atualizado", "saldo_devedor"

    # Base de calculo do lance livre
    base_calculo_lance_livre: str = "credito_original"

    # Atualizar valor do credito pelo reajuste
    atualizar_valor_credito: bool = True

    # Base de calculo do fundo de reserva
    base_calculo_fundo_reserva: str = "credito_original"

    # Metodo para calculo da TIR
    metodo_tir: str = "fluxo_original"  # "fluxo_original" ou "fluxo_ajustado"


# =============================================================================
# PARAMETROS DE ENTRADA
# =============================================================================

@dataclass
class NasaParams:
    """Parametros completos de entrada para simulacao."""

    valor_credito: float = 500000.0
    prazo_meses: int = 200
    taxa_adm_pct: float = 20.0  # % total sobre o credito
    fundo_reserva_pct: float = 3.0  # % total sobre o credito
    momento_contemplacao: int = 36

    # Periodos de distribuicao (ate 6)
    periodos: list = field(default_factory=lambda: [
        {"start": 1, "end": 200, "fc_pct": 1.0, "ta_pct": 100.0, "fr_pct": 100.0}
    ])

    # Lance
    lance_embutido_pct: float = 0.0  # % sobre credito
    lance_livre_pct: float = 0.0  # % sobre credito
    lance_embutido_valor: float = 0.0  # valor absoluto (prioridade se > 0)
    lance_livre_valor: float = 0.0  # valor absoluto (prioridade se > 0)

    # Reajuste
    reajuste_pre_pct: float = 0.0  # % por periodo
    reajuste_pos_pct: float = 0.0
    reajuste_pre_freq: str = "Anual"  # Mensal/Bimestral/Trimestral/Semestral/Anual
    reajuste_pos_freq: str = "Anual"

    # Seguro
    seguro_vida_pct: float = 0.0  # % mensal
    seguro_vida_inicio: int = 1  # mes de inicio

    # Antecipacao TA
    antecipacao_ta_pct: float = 0.0  # % do total TA a antecipar
    antecipacao_ta_parcelas: int = 1  # numero de parcelas para diluir

    # Taxas para VPL
    taxa_vp_credito: float = 0.0  # taxa mensal para PV do credito
    tma: float = 0.01  # taxa mensal minima de atratividade
    alm_anual: float = 12.0  # CDI/ALM anual em %
    hurdle_anual: float = 12.0  # Hurdle anual em %

    # Custos acessorios
    custos_acessorios: list = field(default_factory=list)
    # [{"descricao": "Avaliacao", "valor": 5000.0, "momento": 1}, ...]

    # ---- Compatibilidade com nasa_engine.py ----
    valor_carta: float = 0.0
    taxa_adm: float = 0.0
    fundo_reserva: float = 0.0
    seguro: float = 0.0
    prazo_contemp: int = 0
    parcela_red_pct: float = 100.0
    lance_livre_pct_compat: float = 0.0
    lance_embutido_pct_compat: float = 0.0
    correcao_anual: float = 0.0


# =============================================================================
# GERENCIADOR DE CENARIOS
# =============================================================================

class ScenarioManager:
    """Salva, carrega e compara ate 10 cenarios."""

    def __init__(self):
        self._scenarios = {}  # {index: {"name": str, "params": dict, "results": dict}}
        self._next_id = 1

    def save_scenario(self, name, params, results):
        """Salva cenario. Retorna indice."""
        if len(self._scenarios) >= 10:
            raise ValueError("Maximo de 10 cenarios atingido. Limpe antes de salvar.")
        idx = self._next_id
        self._scenarios[idx] = {
            "name": name,
            "params": copy.deepcopy(params),
            "results": copy.deepcopy(results),
        }
        self._next_id += 1
        return idx

    def load_scenario(self, index):
        """Carrega cenario pelo indice."""
        if index not in self._scenarios:
            raise KeyError(f"Cenario {index} nao encontrado.")
        return copy.deepcopy(self._scenarios[index])

    def clear_scenarios(self):
        """Remove todos os cenarios."""
        self._scenarios.clear()
        self._next_id = 1

    def delete_scenario(self, index):
        """Remove cenario especifico."""
        if index in self._scenarios:
            del self._scenarios[index]

    def list_scenarios(self):
        """Lista cenarios salvos."""
        return [
            {"index": idx, "name": s["name"]}
            for idx, s in sorted(self._scenarios.items())
        ]

    def compare_scenarios(self, indices):
        """
        Compara cenarios pelos indices.
        Retorna dict com metricas lado a lado.
        """
        result = {}
        for idx in indices:
            if idx not in self._scenarios:
                continue
            sc = self._scenarios[idx]
            r = sc["results"]
            result[idx] = {
                "name": sc["name"],
                "valor_credito": sc["params"].get("valor_credito", 0),
                "prazo_meses": sc["params"].get("prazo_meses", 0),
                "total_pago": r.get("total_pago", 0),
                "tir_mensal": r.get("tir_mensal", 0),
                "tir_anual": r.get("tir_anual", 0),
                "delta_vpl": r.get("delta_vpl", 0),
                "parcela_media": r.get("parcela_media", 0),
            }
        return result


# =============================================================================
# MOTOR PRINCIPAL - NASA ENGINE HD
# =============================================================================

class NasaEngineHD:
    """
    Motor de calculo completo da NASA NOVA HD VPL.
    Substitui integralmente a planilha Excel com 21 abas.
    """

    def __init__(self, config=None):
        self.config = config if config is not None else NasaConfig()
        self.scenarios = ScenarioManager()

    # -----------------------------------------------------------------
    # Metodos auxiliares internos
    # -----------------------------------------------------------------

    def _resolve_lance(self, params, tipo="embutido"):
        """Resolve valor do lance (absoluto ou percentual)."""
        credito = params.get("valor_credito", 0)
        if tipo == "embutido":
            val = params.get("lance_embutido_valor", 0)
            if val and val > 0:
                return val
            pct = params.get("lance_embutido_pct", 0) / 100
            return credito * pct
        else:
            val = params.get("lance_livre_valor", 0)
            if val and val > 0:
                return val
            pct = params.get("lance_livre_pct", 0) / 100
            return credito * pct

    def _resolve_reajuste_schedule(self, prazo, contemp, params):
        """
        Monta vetor de reajuste acumulado para cada mes.
        Retorna lista de fatores (1 + acum) para meses 0..prazo.
        """
        pre_rate = params.get("reajuste_pre_pct", 0) / 100
        pos_rate = params.get("reajuste_pos_pct", 0) / 100
        pre_freq = FREQ_MAP.get(params.get("reajuste_pre_freq", "Anual"), 12)
        pos_freq = FREQ_MAP.get(params.get("reajuste_pos_freq", "Anual"), 12)

        fatores = [1.0] * (prazo + 1)
        acum = 0.0

        for m in range(1, prazo + 1):
            if m <= contemp:
                rate = pre_rate
                freq = pre_freq
            else:
                rate = pos_rate
                freq = pos_freq

            # Aplica reajuste nos meses correspondentes a frequencia
            if freq > 0 and m % freq == 0:
                acum = (1 + acum) * (1 + rate) - 1

            fatores[m] = 1 + acum

        return fatores

    def _build_period_distribution(self, params):
        """
        Constroi distribuicao por periodo (Apoio Calculos).
        Usa GoalSeek para encontrar multiplicador que faz a soma = 100%.

        Retorna listas de % mensal para FC, TA, FR (indices 0..prazo).
        """
        prazo = params.get("prazo_meses", 200)
        credito = params.get("valor_credito", 0)
        periodos = params.get("periodos", [])

        if not periodos:
            # Distribuicao linear simples
            pct = 1.0 / prazo if prazo > 0 else 0
            fc_dist = [0.0] + [pct] * prazo
            ta_dist = [0.0] + [pct] * prazo
            fr_dist = [0.0] + [pct] * prazo
            return fc_dist, ta_dist, fr_dist

        # --- Fundo Comum ---
        fc_dist = self._solve_distribution(prazo, periodos, "fc_pct")

        # --- Taxa Administracao ---
        ta_dist = self._solve_distribution(prazo, periodos, "ta_pct")

        # --- Fundo Reserva ---
        fr_dist = self._solve_distribution(prazo, periodos, "fr_pct")

        return fc_dist, ta_dist, fr_dist

    def _solve_distribution(self, prazo, periodos, key):
        """
        Resolve distribuicao de um componente usando GoalSeek.
        Cada periodo tem um peso relativo (pct).
        O GoalSeek encontra o multiplicador tal que a soma = 1.0 (100%).
        """
        # Monta vetor de pesos brutos
        pesos_brutos = [0.0] * (prazo + 1)
        for p in periodos:
            start = max(1, p.get("start", 1))
            end = min(prazo, p.get("end", prazo))
            peso = p.get(key, 1.0)
            if isinstance(peso, (int, float)):
                # Se ta_pct/fr_pct estao em %, converter
                if key in ("ta_pct", "fr_pct") and peso > 1:
                    peso = peso / 100.0
                for m in range(start, end + 1):
                    pesos_brutos[m] = peso

        # Preenche meses sem periodo com 0
        total_peso = sum(pesos_brutos[1:])
        if total_peso == 0:
            # Distribuicao linear
            pct = 1.0 / prazo if prazo > 0 else 0
            return [0.0] + [pct] * prazo

        # GoalSeek: encontra multiplicador tal que sum(multiplicador * pesos_brutos[m]) = 1.0
        def soma_func(mult):
            return sum(mult * pesos_brutos[m] for m in range(1, prazo + 1))

        multiplicador = goal_seek(soma_func, 1.0, 1.0 / total_peso,
                                  tolerance=1e-12, lo=0.0, hi=10.0 / max(total_peso, 1e-10))

        dist = [0.0] * (prazo + 1)
        for m in range(1, prazo + 1):
            dist[m] = multiplicador * pesos_brutos[m]

        return dist

    # -----------------------------------------------------------------
    # 1. FLUXO FINANCEIRO COMPLETO (35 colunas)
    # -----------------------------------------------------------------

    def calcular_fluxo_completo(self, params):
        """
        Calculo completo do fluxo financeiro do consorcio (35 colunas).

        Replica a aba 'Fluxo Financeiro do Consorcio' da planilha NASA.
        Cada linha representa um mes (0 a prazo).

        Args:
            params: dict com todos os parametros (ver NasaParams)

        Returns:
            dict com:
                fluxo: lista de dicts (uma entrada por mes, 35+ campos)
                cashflow: lista de floats para TIR/VPL
                cashflow_tir: lista de floats ajustada para TIR
                totais: dict com totalizacoes
                metricas: dict com TIR, CET, etc.
        """
        credito = params.get("valor_credito", 0)
        prazo = params.get("prazo_meses", 200)
        contemp = params.get("momento_contemplacao", 36)
        taxa_adm_total = params.get("taxa_adm_pct", 20.0) / 100  # decimal
        fundo_reserva_total = params.get("fundo_reserva_pct", 3.0) / 100
        seguro_pct = params.get("seguro_vida_pct", 0) / 100  # mensal
        seguro_inicio = params.get("seguro_vida_inicio", 1)

        lance_emb_val = self._resolve_lance(params, "embutido")
        lance_livre_val = self._resolve_lance(params, "livre")

        antecipacao_ta_pct = params.get("antecipacao_ta_pct", 0) / 100
        antecipacao_ta_parcelas = max(1, params.get("antecipacao_ta_parcelas", 1))

        custos = params.get("custos_acessorios", [])

        # Distribuicao por periodo
        fc_dist, ta_dist, fr_dist = self._build_period_distribution(params)

        # Reajuste
        fatores_reajuste = self._resolve_reajuste_schedule(prazo, contemp, params)

        # Valores totais dos componentes
        valor_ta_total = credito * taxa_adm_total
        valor_fr_total = credito * fundo_reserva_total

        # Antecipacao de TA
        ta_antecipada_total = valor_ta_total * antecipacao_ta_pct
        ta_antecipada_por_parcela = ta_antecipada_total / antecipacao_ta_parcelas if antecipacao_ta_parcelas > 0 else 0

        # Momento do lance embutido
        lance_emb_inicio = contemp if self.config.momento_lance_embutido == "na_contemplacao" else 1

        # Fluxo
        fluxo = []
        saldo_principal = credito  # saldo do fundo comum (Col K)
        saldo_devedor = credito  # saldo devedor total (Col W)
        credito_reajustado = credito  # Col AG
        total_pago = 0.0
        total_fc = 0.0
        total_ta = 0.0
        total_fr = 0.0
        total_seguro = 0.0
        total_custos_acessorios = 0.0
        cum_fc_pct = 0.0  # Col I - acumulado %

        cashflow = []  # Fluxo de caixa (Col AH)
        cashflow_tir = []  # Fluxo de caixa para TIR (Col AI)

        for m in range(prazo + 1):
            row = {}
            row["mes"] = m  # Col B
            row["meses_restantes"] = prazo - m  # Col C

            if m == 0:
                # Mes 0: inicio, sem pagamento
                row["valor_base_fc"] = credito  # Col D
                row["lance_embutido"] = 0.0  # Col E
                row["lance_livre"] = 0.0  # Col F
                row["valor_base_final"] = credito  # Col G
                row["pct_mensal_fc"] = 0.0  # Col H
                row["pct_acum_fc"] = 0.0  # Col I
                row["amortizacao"] = 0.0  # Col J
                row["saldo_principal"] = credito  # Col K
                row["taxa_adm_antecipada"] = 0.0  # Col L
                row["pct_ta_mensal"] = 0.0  # Col M
                row["pct_ta_acum"] = 0.0  # Col N
                row["valor_parcela_ta"] = 0.0  # Col O
                row["pct_fr_mensal"] = 0.0  # Col P
                row["pct_fr_acum"] = 0.0  # Col Q
                row["pct_fr_base"] = 0.0  # Col R
                row["pct_fr_calc"] = 0.0  # Col S
                row["fr_saldo"] = 0.0  # Col T
                row["valor_fundo_reserva"] = 0.0  # Col U
                row["valor_parcela"] = 0.0  # Col V
                row["saldo_devedor"] = credito  # Col W
                row["peso_parcela"] = 0.0  # Col X
                row["pct_reajuste"] = 0.0  # Col Y
                row["pct_reajuste_acum"] = 0.0  # Col Z
                row["parcela_apos_reajuste"] = 0.0  # Col AA
                row["saldo_devedor_reajustado"] = credito  # Col AB
                row["seguro_vida"] = 0.0  # Col AC
                row["parcela_com_seguro"] = 0.0  # Col AD
                row["outros_custos"] = 0.0  # Col AE
                row["carta_credito_original"] = credito  # Col AF
                row["carta_credito_reajustada"] = credito  # Col AG
                row["fluxo_caixa"] = 0.0  # Col AH
                row["fluxo_caixa_tir"] = 0.0  # Col AI
                row["credito_recebido"] = 0.0
                row["fator_reajuste"] = 1.0

                saldo_principal = credito
                saldo_devedor = credito

                fluxo.append(row)
                cashflow.append(0.0)
                cashflow_tir.append(0.0)
                continue

            # --- Col D: Valor-Base Fundo Comum ---
            valor_base_fc = credito
            if self.config.atualizar_valor_credito:
                valor_base_fc = credito * fatores_reajuste[m]
                credito_reajustado = valor_base_fc

            row["valor_base_fc"] = valor_base_fc

            # --- Col E: Lance Embutido ---
            lance_emb_mes = 0.0
            if self.config.momento_lance_embutido == "na_contemplacao":
                if m == contemp:
                    lance_emb_mes = -lance_emb_val
            else:  # desde_inicio
                if m == 1:
                    lance_emb_mes = -lance_emb_val

            row["lance_embutido"] = lance_emb_mes

            # --- Col F: Lance Livre ---
            lance_livre_mes = 0.0
            if m == contemp:
                lance_livre_mes = -lance_livre_val
            row["lance_livre"] = lance_livre_mes

            # --- Col G: Valor-Base Final ---
            valor_base_final = valor_base_fc + lance_emb_mes + lance_livre_mes
            row["valor_base_final"] = valor_base_final

            # --- Col H: % Mensal amortizacao Fundo Comum ---
            pct_mensal_fc = fc_dist[m] if m < len(fc_dist) else 0.0
            row["pct_mensal_fc"] = pct_mensal_fc

            # --- Col I: % Acumulado FC ---
            cum_fc_pct += pct_mensal_fc
            row["pct_acum_fc"] = cum_fc_pct

            # --- Col J: Amortizacao ---
            # Formula: -K(prev)/(1-I(prev)) * H + E + F
            prev_pct_acum = fluxo[m - 1]["pct_acum_fc"] if m > 0 else 0.0
            prev_saldo = saldo_principal

            denom = 1 - prev_pct_acum
            if abs(denom) > 1e-12 and pct_mensal_fc > 0:
                amort = -(prev_saldo / denom) * pct_mensal_fc
            else:
                amort = 0.0

            # Lance afeta amortizacao no mes de ocorrencia
            amort += lance_emb_mes + lance_livre_mes

            row["amortizacao"] = amort

            # --- Col K: Saldo do Principal ---
            saldo_principal = prev_saldo + amort
            if abs(saldo_principal) < 0.01:
                saldo_principal = 0.0
            row["saldo_principal"] = saldo_principal

            # --- Col L: Taxa de Administracao antecipada ---
            ta_antecipada_mes = 0.0
            if antecipacao_ta_pct > 0:
                if self.config.momento_antecipacao_ta == "junto_1a_parcela":
                    if 1 <= m <= antecipacao_ta_parcelas:
                        ta_antecipada_mes = ta_antecipada_por_parcela
                elif self.config.momento_antecipacao_ta == "na_contemplacao":
                    if m == contemp:
                        ta_antecipada_mes = ta_antecipada_total
                elif self.config.momento_antecipacao_ta == "diluida":
                    if 1 <= m <= prazo:
                        ta_antecipada_mes = ta_antecipada_total / prazo
            row["taxa_adm_antecipada"] = ta_antecipada_mes

            # --- Col M/N: % TA mensal e acumulado ---
            pct_ta_mensal = ta_dist[m] if m < len(ta_dist) else 0.0
            # Ajustar se ha antecipacao (reduz distribuicao normal)
            pct_ta_efetivo = pct_ta_mensal * (1 - antecipacao_ta_pct)
            row["pct_ta_mensal"] = pct_ta_mensal
            ta_acum_prev = fluxo[m - 1]["pct_ta_acum"] if m > 0 else 0.0
            row["pct_ta_acum"] = ta_acum_prev + pct_ta_mensal

            # --- Col O: Valor Parcela TA ---
            valor_parcela_ta = valor_base_fc * pct_ta_efetivo + ta_antecipada_mes
            row["valor_parcela_ta"] = valor_parcela_ta

            # --- Col P-T: Fundo Reserva ---
            pct_fr_mensal = fr_dist[m] if m < len(fr_dist) else 0.0
            row["pct_fr_mensal"] = pct_fr_mensal
            fr_acum_prev = fluxo[m - 1].get("pct_fr_acum", 0) if m > 0 else 0.0
            row["pct_fr_acum"] = fr_acum_prev + pct_fr_mensal

            # Base de calculo FR
            if self.config.base_calculo_fundo_reserva == "credito_original":
                fr_base = credito
            elif self.config.base_calculo_fundo_reserva == "saldo_devedor":
                fr_base = abs(saldo_principal)
            else:
                fr_base = valor_base_fc
            row["pct_fr_base"] = fundo_reserva_total
            row["pct_fr_calc"] = pct_fr_mensal * fundo_reserva_total

            # Col U: Valor Fundo Reserva
            valor_fundo_reserva = fr_base * pct_fr_mensal * fundo_reserva_total / (
                fundo_reserva_total if fundo_reserva_total > 0 else 1)
            # Simplificado: distribui linearmente o FR total
            valor_fundo_reserva = valor_fr_total * pct_fr_mensal
            row["valor_fundo_reserva"] = valor_fundo_reserva
            row["fr_saldo"] = 0.0  # Col T placeholder

            # --- Col V: Valor da Parcela total ---
            # Parcela = amortizacao_pura(sem lance) + TA + FR
            amort_pura = amort - lance_emb_mes - lance_livre_mes
            valor_parcela = abs(amort_pura) + valor_parcela_ta + valor_fundo_reserva
            row["valor_parcela"] = valor_parcela

            # --- Col W: Saldo Devedor ---
            saldo_devedor = abs(saldo_principal) + valor_ta_total * (1 - row["pct_ta_acum"]) + valor_fr_total * (1 - row["pct_fr_acum"])
            row["saldo_devedor"] = saldo_devedor

            # --- Col X: Peso Parcela ---
            peso = valor_parcela / credito if credito > 0 else 0
            row["peso_parcela"] = peso

            # --- Col Y-Z: Reajuste ---
            fator_reaj = fatores_reajuste[m]
            pct_reaj_period = (fator_reaj / fatores_reajuste[m - 1] - 1) if m > 0 and fatores_reajuste[m - 1] > 0 else 0
            row["pct_reajuste"] = pct_reaj_period
            row["pct_reajuste_acum"] = fator_reaj - 1
            row["fator_reajuste"] = fator_reaj

            # --- Col AA: Parcela Apos Reajuste ---
            parcela_reajustada = valor_parcela * fator_reaj
            row["parcela_apos_reajuste"] = parcela_reajustada

            # --- Col AB: Saldo Devedor c/Reajuste ---
            saldo_dev_reaj = saldo_devedor * fator_reaj
            row["saldo_devedor_reajustado"] = saldo_dev_reaj

            # --- Col AC: Seguro de Vida ---
            seguro_mes = 0.0
            if seguro_pct > 0 and m >= seguro_inicio:
                if self.config.seguro_base == "saldo_devedor":
                    seguro_mes = abs(saldo_dev_reaj) * seguro_pct
                else:
                    seguro_mes = credito_reajustado * seguro_pct
            row["seguro_vida"] = seguro_mes

            # --- Col AD: Parcela c/Seguro ---
            parcela_com_seguro = parcela_reajustada + seguro_mes
            row["parcela_com_seguro"] = parcela_com_seguro

            # --- Col AE: Outros Custos ---
            custo_mes = sum(c.get("valor", 0) for c in custos if c.get("momento", -1) == m)
            row["outros_custos"] = custo_mes
            total_custos_acessorios += custo_mes

            # --- Col AF-AG: Carta de Credito ---
            row["carta_credito_original"] = credito
            row["carta_credito_reajustada"] = credito * fator_reaj

            # --- Credito recebido ---
            credito_recebido = 0.0
            if m == contemp:
                carta_liquida = credito - lance_emb_val
                credito_recebido = carta_liquida * fator_reaj if self.config.atualizar_valor_credito else carta_liquida
            row["credito_recebido"] = credito_recebido

            # --- Col AH: Fluxo de Caixa ---
            desembolso = parcela_com_seguro + custo_mes
            # Lance livre eh pago a parte da parcela
            lance_livre_real = abs(lance_livre_mes) * fator_reaj if lance_livre_mes != 0 else 0
            lance_emb_real = abs(lance_emb_mes) * fator_reaj if lance_emb_mes != 0 else 0

            fluxo_mes = credito_recebido - desembolso - lance_livre_real
            row["fluxo_caixa"] = fluxo_mes

            # --- Col AI: Fluxo para TIR ---
            if self.config.metodo_tir == "fluxo_ajustado":
                fluxo_tir = fluxo_mes
            else:
                fluxo_tir = credito_recebido - desembolso - lance_livre_real
            row["fluxo_caixa_tir"] = fluxo_tir

            # Acumuladores
            total_pago += desembolso + lance_livre_real
            total_fc += abs(amort_pura) * fator_reaj
            total_ta += valor_parcela_ta * fator_reaj
            total_fr += valor_fundo_reserva * fator_reaj
            total_seguro += seguro_mes

            fluxo.append(row)
            cashflow.append(fluxo_mes)
            cashflow_tir.append(fluxo_tir)

        # --- Metricas ---
        tir_m = _irr(cashflow_tir, guess=0.005)
        tir_a = _annual_from_monthly(tir_m)

        parcelas_reaj = [r["parcela_apos_reajuste"] for r in fluxo if r["mes"] > 0]
        parcela_media = sum(parcelas_reaj) / len(parcelas_reaj) if parcelas_reaj else 0
        parcela_max = max(parcelas_reaj) if parcelas_reaj else 0
        parcela_min = min(p for p in parcelas_reaj if p > 0) if any(p > 0 for p in parcelas_reaj) else 0

        carta_liquida = credito - lance_emb_val

        return {
            "fluxo": fluxo,
            "cashflow": cashflow,
            "cashflow_tir": cashflow_tir,
            "totais": {
                "total_pago": total_pago,
                "total_fundo_comum": total_fc,
                "total_taxa_adm": total_ta,
                "total_fundo_reserva": total_fr,
                "total_seguro": total_seguro,
                "total_custos_acessorios": total_custos_acessorios,
                "carta_liquida": carta_liquida,
                "lance_embutido_valor": lance_emb_val,
                "lance_livre_valor": lance_livre_val,
            },
            "metricas": {
                "tir_mensal": tir_m,
                "tir_anual": tir_a,
                "cet_anual": tir_a,
                "parcela_media": parcela_media,
                "parcela_maxima": parcela_max,
                "parcela_minima": parcela_min,
                "custo_total_pct": (total_pago / credito * 100) if credito > 0 else 0,
            },
            # Compatibilidade com nasa_engine.py
            "fluxo_mensal": fluxo,
            "cashflow_consorcio": cashflow,
            "total_pago": total_pago,
            "carta_liquida": carta_liquida,
            "lance_livre_valor": lance_livre_val,
            "lance_embutido_valor": lance_emb_val,
        }

    # -----------------------------------------------------------------
    # 2. VPL HD (Goal-Based com taxas duais)
    # -----------------------------------------------------------------

    def calcular_vpl_hd(self, params, fluxo=None):
        """
        Analise VPL HD com taxas duais (ALM pre-T, Hurdle pos-T).

        Replica a aba 'COMPARATIVO DE VPL' da planilha NASA.

        Args:
            params: dict com parametros (inclui alm_anual, hurdle_anual)
            fluxo: resultado de calcular_fluxo_completo (ou None para calcular)

        Returns:
            dict com b0, h0, d0, pv_pos_t, delta_vpl, etc.
        """
        if fluxo is None:
            fluxo = self.calcular_fluxo_completo(params)

        alm_a = params.get("alm_anual", 12.0) / 100
        hurdle_a = params.get("hurdle_anual", 12.0) / 100
        alm_m = _monthly_from_annual(alm_a)
        hurdle_m = _monthly_from_annual(hurdle_a)

        contemp = params.get("momento_contemplacao", 36)
        carta_liquida = fluxo.get("carta_liquida", fluxo.get("totais", {}).get("carta_liquida", 0))

        fluxo_mensal = fluxo.get("fluxo", fluxo.get("fluxo_mensal", []))

        # B0: PV do credito recebido na contemplacao
        credito_recebido = 0
        for f in fluxo_mensal:
            cr = f.get("credito_recebido", 0)
            if cr > 0:
                credito_recebido = cr
                break
        if credito_recebido == 0:
            credito_recebido = carta_liquida

        b0 = credito_recebido / (1 + alm_m) ** contemp if contemp > 0 else credito_recebido

        # H0: PV dos pagamentos pre-contemplacao + lances
        h0 = 0.0
        pv_pre_t_detail = []
        for f in fluxo_mensal:
            m = f.get("mes", 0)
            if 0 < m <= contemp:
                pagamento = f.get("parcela_com_seguro", f.get("parcela", 0))
                pagamento += f.get("outros_custos", 0)
                lance = abs(f.get("lance_livre", 0))
                total_m = pagamento + lance
                pv_m = total_m / (1 + alm_m) ** m
                h0 += pv_m
                pv_pre_t_detail.append({"mes": m, "valor": total_m, "pv": pv_m})

        # D0: valor criado
        d0 = b0 - h0

        # PV parcelas pos-contemplacao (descontadas a hurdle)
        pv_pos_t_at_contemp = 0.0
        pv_pos_t_detail = []
        for f in fluxo_mensal:
            m = f.get("mes", 0)
            if m > contemp:
                pagamento = f.get("parcela_com_seguro", f.get("parcela", 0))
                pagamento += f.get("outros_custos", 0)
                meses_apos = m - contemp
                pv_m = pagamento / (1 + hurdle_m) ** meses_apos
                pv_pos_t_at_contemp += pv_m
                pv_pos_t_detail.append({"mes": m, "valor": pagamento, "pv": pv_m})

        # Trazer ao t=0
        pv_pos_t = pv_pos_t_at_contemp / (1 + alm_m) ** contemp if contemp > 0 else pv_pos_t_at_contemp

        # Delta VPL
        delta_vpl = d0 - pv_pos_t
        cria_valor = delta_vpl >= 0

        # TIR
        cf = fluxo.get("cashflow_tir", fluxo.get("cashflow", []))
        tir_m = _irr(cf, guess=0.005)
        tir_a = _annual_from_monthly(tir_m)

        # VPL total a taxa ALM
        vpl_total = _npv(alm_m, cf)

        # Break-even lance
        be_lance = self._buscar_break_even_lance_hd(params, alm_m, hurdle_m)

        return {
            "b0": b0,
            "h0": h0,
            "d0": d0,
            "pv_pos_t": pv_pos_t,
            "pv_pos_t_at_contemp": pv_pos_t_at_contemp,
            "delta_vpl": delta_vpl,
            "cria_valor": cria_valor,
            "break_even_lance": be_lance,
            "tir_mensal": tir_m,
            "tir_anual": tir_a,
            "cet_anual": tir_a,
            "vpl_total": vpl_total,
            "pv_pre_t_detail": pv_pre_t_detail,
            "pv_pos_t_detail": pv_pos_t_detail,
        }

    def _buscar_break_even_lance_hd(self, params, alm_m, hurdle_m):
        """Busca binaria para lance que zera Delta VPL."""
        def calc_delta(lance_pct):
            p = dict(params)
            p["lance_livre_pct"] = lance_pct
            p["lance_livre_valor"] = 0
            fl = self.calcular_fluxo_completo(p)
            contemp = p.get("momento_contemplacao", 36)
            carta_liq = fl["carta_liquida"]

            fluxo_mensal = fl["fluxo"]

            credito_rec = 0
            for f in fluxo_mensal:
                cr = f.get("credito_recebido", 0)
                if cr > 0:
                    credito_rec = cr
                    break
            if credito_rec == 0:
                credito_rec = carta_liq

            b0 = credito_rec / (1 + alm_m) ** contemp if contemp > 0 else credito_rec

            h0 = 0.0
            for f in fluxo_mensal:
                m = f["mes"]
                if 0 < m <= contemp:
                    pag = f.get("parcela_com_seguro", f.get("parcela", 0)) + f.get("outros_custos", 0)
                    lance = abs(f.get("lance_livre", 0))
                    h0 += (pag + lance) / (1 + alm_m) ** m

            d0 = b0 - h0

            pv_pos = 0.0
            for f in fluxo_mensal:
                m = f["mes"]
                if m > contemp:
                    pag = f.get("parcela_com_seguro", f.get("parcela", 0)) + f.get("outros_custos", 0)
                    pv_pos += pag / (1 + hurdle_m) ** (m - contemp)
            if contemp > 0:
                pv_pos /= (1 + alm_m) ** contemp

            return d0 - pv_pos

        try:
            return goal_seek(calc_delta, 0.0, 20.0, tolerance=0.01, lo=0.0, hi=90.0)
        except Exception:
            return 0.0

    # -----------------------------------------------------------------
    # 3. FINANCIAMENTO (SAC/Price com IOF)
    # -----------------------------------------------------------------

    def calcular_financiamento(self, params):
        """
        Simulacao completa de financiamento (SAC ou Price) com IOF.

        Args:
            params: dict com:
                valor: valor financiado
                prazo_meses: prazo
                taxa_mensal_pct: taxa juros mensal %
                metodo: "price" ou "sac"
                carencia: meses de carencia (default 0)
                custos_adicionais: lista de custos
                calcular_iof: bool (default True)

        Returns:
            dict com parcelas, cashflow, totais, IOF, etc.
        """
        valor = params.get("valor", 0)
        prazo = params.get("prazo_meses", 0)
        taxa = params.get("taxa_mensal_pct", 0) / 100
        metodo = params.get("metodo", "price").lower()
        carencia = params.get("carencia", 0)
        calc_iof = params.get("calcular_iof", True)
        custos_add = params.get("custos_adicionais", [])

        saldo = valor
        parcelas = []
        cashflow = [valor]  # mes 0: recebe
        total_pago = 0.0
        total_juros = 0.0
        total_amort = 0.0

        prazo_amort = prazo - carencia

        # PMT para Price
        if metodo == "price" and taxa > 0 and prazo_amort > 0:
            pmt_val = _pmt(taxa, prazo_amort, valor)
        elif prazo_amort > 0:
            pmt_val = valor / prazo_amort
        else:
            pmt_val = 0

        for m in range(1, prazo + 1):
            juros = saldo * taxa

            if m <= carencia:
                # Carencia: paga so juros
                amort = 0.0
                parcela = juros
            elif metodo == "price":
                parcela = pmt_val
                amort = parcela - juros
            else:  # SAC
                amort = valor / prazo_amort if prazo_amort > 0 else 0
                parcela = amort + juros

            saldo -= amort
            saldo = max(0, saldo)
            total_pago += parcela
            total_juros += juros
            total_amort += amort

            parcelas.append({
                "mes": m,
                "parcela": parcela,
                "juros": juros,
                "amortizacao": amort,
                "saldo": saldo,
            })
            cashflow.append(-parcela)

        # IOF
        iof_total = 0.0
        if calc_iof:
            iof_total = self._calcular_iof(valor, parcelas)

        # Custos adicionais
        total_custos_add = sum(c.get("valor", 0) for c in custos_add)
        custos_detail = []
        for c in custos_add:
            m_custo = c.get("momento", 0)
            if 0 <= m_custo <= prazo:
                if m_custo == 0:
                    cashflow[0] -= c.get("valor", 0)
                elif m_custo <= len(cashflow) - 1:
                    cashflow[m_custo] -= c.get("valor", 0)
            custos_detail.append(c)

        # CET
        cf_cet = list(cashflow)
        if calc_iof:
            cf_cet[0] = valor - iof_total - total_custos_add
        tir_m = _irr(cf_cet, guess=0.008)
        tir_a = _annual_from_monthly(tir_m)

        return {
            "parcelas": parcelas,
            "cashflow": cashflow,
            "total_pago": total_pago,
            "total_juros": total_juros,
            "total_amortizado": total_amort,
            "valor": valor,
            "iof": iof_total,
            "custos_adicionais": total_custos_add,
            "custo_efetivo_total": total_pago + iof_total + total_custos_add,
            "tir_mensal": tir_m,
            "tir_anual": tir_a,
            "cet_anual": tir_a,
        }

    def _calcular_iof(self, valor_principal, parcelas):
        """
        Calcula IOF sobre operacao de credito.
        IOF = IOF_ADICIONAL (0.38%) sobre principal + IOF_DIARIO sobre cada parcela.
        """
        iof_adicional = valor_principal * IOF_ADICIONAL

        iof_diario_total = 0.0
        for p in parcelas:
            dias = p["mes"] * 30  # Aproximacao: 30 dias/mes
            dias = min(dias, 365)  # IOF diario limitado a 365 dias
            iof_diario_total += p["amortizacao"] * IOF_DIARIO * dias

        return iof_adicional + iof_diario_total

    # -----------------------------------------------------------------
    # 4. CREDITO PARA LANCE (Op. Credito para Lance)
    # -----------------------------------------------------------------

    def calcular_credito_lance(self, params):
        """
        Simulacao do financiamento usado para cobrir o lance.

        Args:
            params: dict com:
                valor_lance: valor do lance a financiar
                prazo_meses: prazo do emprestimo
                taxa_mensal_pct: taxa juros mensal %
                metodo: "price" ou "sac"
                carencia: meses de carencia
                tac: Tarifa de Abertura de Credito
                avaliacao_garantia: custo de avaliacao de garantia
                comissao: comissao do agente
                calcular_iof: bool
                antecipacao_mes: mes para quitar antecipadamente (0 = sem)

        Returns:
            dict com fluxo do financiamento do lance
        """
        valor = params.get("valor_lance", 0)
        prazo = params.get("prazo_meses", 0)
        taxa = params.get("taxa_mensal_pct", 0) / 100
        metodo = params.get("metodo", "price").lower()
        carencia = params.get("carencia", 0)
        tac = params.get("tac", 0)
        aval_garantia = params.get("avaliacao_garantia", 0)
        comissao = params.get("comissao", 0)
        calc_iof = params.get("calcular_iof", True)
        antecipacao = params.get("antecipacao_mes", 0)

        # Custos adicionais no momento 0
        custos_add = []
        if tac > 0:
            custos_add.append({"descricao": "TAC", "valor": tac, "momento": 0})
        if aval_garantia > 0:
            custos_add.append({"descricao": "Avaliacao Garantia", "valor": aval_garantia, "momento": 0})
        if comissao > 0:
            custos_add.append({"descricao": "Comissao", "valor": comissao, "momento": 0})

        fin_params = {
            "valor": valor,
            "prazo_meses": prazo,
            "taxa_mensal_pct": taxa * 100,
            "metodo": metodo,
            "carencia": carencia,
            "calcular_iof": calc_iof,
            "custos_adicionais": custos_add,
        }

        resultado = self.calcular_financiamento(fin_params)

        # Antecipacao
        if antecipacao > 0 and antecipacao < prazo:
            parcelas = resultado["parcelas"]
            saldo_no_mes = 0
            for p in parcelas:
                if p["mes"] == antecipacao:
                    saldo_no_mes = p["saldo"]
                    break

            # Recalcular cashflow com antecipacao
            cf_novo = list(resultado["cashflow"])
            for i in range(antecipacao + 1, len(cf_novo)):
                cf_novo[i] = 0
            if antecipacao < len(cf_novo):
                cf_novo[antecipacao] -= saldo_no_mes

            resultado["cashflow_antecipado"] = cf_novo
            resultado["valor_antecipacao"] = saldo_no_mes
            resultado["mes_antecipacao"] = antecipacao

        resultado["custos_iniciais"] = tac + aval_garantia + comissao

        return resultado

    # -----------------------------------------------------------------
    # 5. CUSTO COMBINADO (Consorcio + Credito Lance)
    # -----------------------------------------------------------------

    def calcular_custo_combinado(self, fluxo_consorcio, fluxo_lance):
        """
        Combina fluxo do consorcio com fluxo do financiamento do lance.

        Args:
            fluxo_consorcio: resultado de calcular_fluxo_completo
            fluxo_lance: resultado de calcular_credito_lance

        Returns:
            dict com fluxo combinado e metricas
        """
        cf_cons = fluxo_consorcio.get("cashflow", [])
        cf_lance = fluxo_lance.get("cashflow", fluxo_lance.get("cashflow_antecipado", []))

        # Determinar tamanho maximo
        max_len = max(len(cf_cons), len(cf_lance))

        cf_combinado = []
        for i in range(max_len):
            v_cons = cf_cons[i] if i < len(cf_cons) else 0
            v_lance = cf_lance[i] if i < len(cf_lance) else 0
            cf_combinado.append(v_cons + v_lance)

        tir_m = _irr(cf_combinado, guess=0.005)
        tir_a = _annual_from_monthly(tir_m)

        total_cons = fluxo_consorcio.get("total_pago", fluxo_consorcio.get("totais", {}).get("total_pago", 0))
        total_lance = fluxo_lance.get("custo_efetivo_total", fluxo_lance.get("total_pago", 0))

        return {
            "cashflow_combinado": cf_combinado,
            "total_pago_consorcio": total_cons,
            "total_pago_lance": total_lance,
            "total_pago_combinado": total_cons + total_lance,
            "tir_mensal_combinado": tir_m,
            "tir_anual_combinado": tir_a,
            "cet_anual_combinado": tir_a,
        }

    # -----------------------------------------------------------------
    # 6. COMPARATIVO CONSORCIO VS FINANCIAMENTO
    # -----------------------------------------------------------------

    def comparar_consorcio_financiamento(self, params_cons, params_fin):
        """
        Comparacao lado a lado consorcio vs financiamento com PV flows.

        Args:
            params_cons: parametros do consorcio
            params_fin: parametros do financiamento
                {valor, prazo_meses, taxa_mensal_pct, metodo, ...}

        Returns:
            dict com comparativo completo
        """
        tma_m = params_cons.get("tma", 0.01)
        if tma_m > 1:
            tma_m = tma_m / 100  # Corrige se veio em %
        alm_a = params_cons.get("alm_anual", 12.0) / 100
        alm_m = _monthly_from_annual(alm_a)

        # Consorcio
        fluxo_c = self.calcular_fluxo_completo(params_cons)
        cf_c = fluxo_c.get("cashflow", [])
        vpl_c_alm = _npv(alm_m, cf_c)
        vpl_c_tma = _npv(tma_m, cf_c)
        tir_c = _irr(cf_c, guess=0.005)

        # Financiamento
        fluxo_f = self.calcular_financiamento(params_fin)
        cf_f = fluxo_f["cashflow"]
        vpl_f_alm = _npv(alm_m, cf_f)
        vpl_f_tma = _npv(tma_m, cf_f)
        tir_f = _irr(cf_f, guess=0.008)

        # Valores nominais
        total_cons = fluxo_c.get("total_pago", fluxo_c.get("totais", {}).get("total_pago", 0))
        total_fin = fluxo_f["total_pago"]

        carta_liq = fluxo_c.get("carta_liquida", fluxo_c.get("totais", {}).get("carta_liquida", 0))
        valor_fin = fluxo_f["valor"]

        razao_c = total_cons / carta_liq if carta_liq > 0 else 0
        razao_f = total_fin / valor_fin if valor_fin > 0 else 0

        # PV de cada parcela (para grafico)
        pv_cons = []
        for i, cf in enumerate(cf_c):
            pv_cons.append(cf / (1 + tma_m) ** i if tma_m > 0 else cf)

        pv_fin = []
        for i, cf in enumerate(cf_f):
            pv_fin.append(cf / (1 + tma_m) ** i if tma_m > 0 else cf)

        return {
            "consorcio": fluxo_c,
            "financiamento": fluxo_f,
            # Valores nominais
            "total_pago_consorcio": total_cons,
            "total_pago_financiamento": total_fin,
            "economia_nominal": total_fin - total_cons,
            # VPL
            "vpl_consorcio": vpl_c_alm,
            "vpl_financiamento": vpl_f_alm,
            "economia_vpl": vpl_c_alm - vpl_f_alm,
            "vpl_consorcio_tma": vpl_c_tma,
            "vpl_financiamento_tma": vpl_f_tma,
            "economia_vpl_tma": vpl_c_tma - vpl_f_tma,
            # TIR
            "tir_consorcio_mensal": tir_c,
            "tir_consorcio_anual": _annual_from_monthly(tir_c),
            "tir_financ_mensal": tir_f,
            "tir_financ_anual": _annual_from_monthly(tir_f),
            # Razoes
            "razao_vpl_consorcio": razao_c,
            "razao_vpl_financ": razao_f,
            # PV flows
            "pv_consorcio": pv_cons,
            "pv_financiamento": pv_fin,
        }

    # -----------------------------------------------------------------
    # 7. VENDA DA OPERACAO
    # -----------------------------------------------------------------

    def calcular_venda_operacao(self, fluxo, params_venda):
        """
        Analise de venda da operacao de consorcio.

        Args:
            fluxo: resultado de calcular_fluxo_completo
            params_venda: dict com:
                momento_venda: mes da venda
                valor_venda: valor recebido na venda
                tma: taxa minima atratividade mensal

        Returns:
            dict com VPL da venda, ganho, IRR do comprador, etc.
        """
        momento = params_venda.get("momento_venda", 0)
        valor_venda = params_venda.get("valor_venda", 0)
        tma = params_venda.get("tma", 0.01)

        cf_original = fluxo.get("cashflow", [])
        fluxo_mensal = fluxo.get("fluxo", fluxo.get("fluxo_mensal", []))

        # Fluxo do vendedor: ate momento_venda + recebe valor_venda
        cf_vendedor = []
        total_gasto = 0.0
        for i in range(min(momento + 1, len(cf_original))):
            cf_vendedor.append(cf_original[i])
            if cf_original[i] < 0:
                total_gasto += abs(cf_original[i])

        # Adiciona valor de venda no ultimo mes
        if len(cf_vendedor) > 0:
            cf_vendedor[-1] += valor_venda
        else:
            cf_vendedor.append(valor_venda)

        vpl_vendedor = _npv(tma, cf_vendedor)
        tir_vendedor = _irr(cf_vendedor, guess=0.01)

        ganho_nominal = valor_venda - total_gasto
        ganho_pct = (ganho_nominal / total_gasto * 100) if total_gasto > 0 else 0

        prazo_medio = momento / 2  # simplificacao
        ganho_mensal = (ganho_nominal / momento) if momento > 0 else 0
        margem_mensal = (ganho_mensal / total_gasto * 100) if total_gasto > 0 else 0

        # Fluxo do comprador: paga valor_venda e continua com parcelas restantes
        cf_comprador = [-valor_venda]
        for i in range(momento + 1, len(cf_original)):
            cf_comprador.append(cf_original[i])

        tir_comprador = _irr(cf_comprador, guess=0.005) if len(cf_comprador) > 1 else 0
        vpl_comprador = _npv(tma, cf_comprador)

        return {
            "cashflow_vendedor": cf_vendedor,
            "cashflow_comprador": cf_comprador,
            "vpl_vendedor": vpl_vendedor,
            "vpl_comprador": vpl_comprador,
            "tir_vendedor_mensal": tir_vendedor,
            "tir_vendedor_anual": _annual_from_monthly(tir_vendedor),
            "tir_comprador_mensal": tir_comprador,
            "tir_comprador_anual": _annual_from_monthly(tir_comprador),
            "ganho_nominal": ganho_nominal,
            "ganho_pct": ganho_pct,
            "total_investido": total_gasto,
            "valor_venda": valor_venda,
            "prazo_medio": prazo_medio,
            "ganho_mensal": ganho_mensal,
            "margem_mensal_pct": margem_mensal,
        }

    # -----------------------------------------------------------------
    # 8. CREDITO EQUIVALENTE (GoalSeek)
    # -----------------------------------------------------------------

    def calcular_credito_equivalente(self, params):
        """
        Encontra o valor de credito que cobre todos os custos da operacao.
        Usa GoalSeek para resolver iterativamente.

        Args:
            params: parametros base do consorcio

        Returns:
            float: valor do credito equivalente
        """
        custos_base = sum(c.get("valor", 0) for c in params.get("custos_acessorios", []))
        credito_original = params.get("valor_credito", 0)

        def custo_total_func(credito_teste):
            p = dict(params)
            p["valor_credito"] = credito_teste
            fluxo = self.calcular_fluxo_completo(p)
            total = fluxo.get("total_pago", fluxo.get("totais", {}).get("total_pago", 0))
            # Credito equivalente = credito que faz total_pago = credito
            return total

        try:
            resultado = goal_seek(
                custo_total_func,
                credito_original,
                credito_original * 0.8,
                tolerance=1.0,
                lo=credito_original * 0.1,
                hi=credito_original * 2.0,
            )
            return resultado
        except Exception:
            return credito_original

    # -----------------------------------------------------------------
    # 9. GERAR PARCELAS (Aba Parcelas)
    # -----------------------------------------------------------------

    def gerar_parcelas(self, fluxo):
        """
        Gera tabela de parcelas formatada (como aba Parcelas da planilha).

        Args:
            fluxo: resultado de calcular_fluxo_completo

        Returns:
            list de dicts com parcelas detalhadas para exibicao
        """
        fluxo_mensal = fluxo.get("fluxo", fluxo.get("fluxo_mensal", []))
        parcelas = []

        for f in fluxo_mensal:
            m = f.get("mes", 0)
            if m == 0:
                continue

            parcelas.append({
                "mes": m,
                "fundo_comum": abs(f.get("amortizacao", 0) - f.get("lance_embutido", 0) - f.get("lance_livre", 0)),
                "taxa_adm": f.get("valor_parcela_ta", 0),
                "fundo_reserva": f.get("valor_fundo_reserva", 0),
                "parcela_base": f.get("valor_parcela", 0),
                "reajuste": f.get("fator_reajuste", 1.0),
                "parcela_reajustada": f.get("parcela_apos_reajuste", 0),
                "seguro": f.get("seguro_vida", 0),
                "parcela_total": f.get("parcela_com_seguro", 0),
                "outros_custos": f.get("outros_custos", 0),
                "desembolso_total": f.get("parcela_com_seguro", 0) + f.get("outros_custos", 0),
                "saldo_devedor": f.get("saldo_devedor_reajustado", f.get("saldo_devedor", 0)),
                "lance_embutido": abs(f.get("lance_embutido", 0)),
                "lance_livre": abs(f.get("lance_livre", 0)),
                "credito_recebido": f.get("credito_recebido", 0),
            })

        return parcelas

    # -----------------------------------------------------------------
    # 10. RESUMO CLIENTE
    # -----------------------------------------------------------------

    def gerar_resumo_cliente(self, params, fluxo, vpl=None):
        """
        Gera dados para relatorio resumo do cliente.

        Args:
            params: parametros da simulacao
            fluxo: resultado de calcular_fluxo_completo
            vpl: resultado de calcular_vpl_hd (opcional)

        Returns:
            dict com todas as metricas para o relatorio
        """
        if vpl is None:
            vpl = self.calcular_vpl_hd(params, fluxo)

        totais = fluxo.get("totais", {})
        metricas = fluxo.get("metricas", {})
        credito = params.get("valor_credito", 0)
        prazo = params.get("prazo_meses", 0)
        contemp = params.get("momento_contemplacao", 0)

        lance_emb = totais.get("lance_embutido_valor", 0)
        lance_livre = totais.get("lance_livre_valor", 0)
        carta_liq = totais.get("carta_liquida", credito)
        total_pago = totais.get("total_pago", 0)

        # Parcelas
        fluxo_mensal = fluxo.get("fluxo", fluxo.get("fluxo_mensal", []))
        parcelas_vals = [f.get("parcela_com_seguro", f.get("parcela_apos_reajuste", 0))
                         for f in fluxo_mensal if f.get("mes", 0) > 0]

        primeira_parcela = parcelas_vals[0] if parcelas_vals else 0
        ultima_parcela = parcelas_vals[-1] if parcelas_vals else 0

        return {
            # Dados da operacao
            "valor_credito": credito,
            "prazo_meses": prazo,
            "momento_contemplacao": contemp,
            "carta_liquida": carta_liq,
            "lance_embutido": lance_emb,
            "lance_embutido_pct": (lance_emb / credito * 100) if credito > 0 else 0,
            "lance_livre": lance_livre,
            "lance_livre_pct": (lance_livre / credito * 100) if credito > 0 else 0,
            "lance_total": lance_emb + lance_livre,
            "lance_total_pct": ((lance_emb + lance_livre) / credito * 100) if credito > 0 else 0,

            # Parcelas
            "primeira_parcela": primeira_parcela,
            "ultima_parcela": ultima_parcela,
            "parcela_media": metricas.get("parcela_media", 0),
            "parcela_maxima": metricas.get("parcela_maxima", 0),
            "parcela_minima": metricas.get("parcela_minima", 0),

            # Totais
            "total_pago": total_pago,
            "total_fundo_comum": totais.get("total_fundo_comum", 0),
            "total_taxa_adm": totais.get("total_taxa_adm", 0),
            "total_fundo_reserva": totais.get("total_fundo_reserva", 0),
            "total_seguro": totais.get("total_seguro", 0),
            "total_custos_acessorios": totais.get("total_custos_acessorios", 0),
            "custo_total_pct": metricas.get("custo_total_pct", 0),

            # Taxas
            "taxa_adm_pct": params.get("taxa_adm_pct", 0),
            "fundo_reserva_pct": params.get("fundo_reserva_pct", 0),
            "seguro_pct": params.get("seguro_vida_pct", 0),

            # Analise VPL
            "tir_mensal": vpl.get("tir_mensal", metricas.get("tir_mensal", 0)),
            "tir_anual": vpl.get("tir_anual", metricas.get("tir_anual", 0)),
            "cet_anual": vpl.get("cet_anual", metricas.get("cet_anual", 0)),
            "b0": vpl.get("b0", 0),
            "h0": vpl.get("h0", 0),
            "d0": vpl.get("d0", 0),
            "pv_pos_t": vpl.get("pv_pos_t", 0),
            "delta_vpl": vpl.get("delta_vpl", 0),
            "cria_valor": vpl.get("cria_valor", False),
            "vpl_total": vpl.get("vpl_total", 0),
            "break_even_lance": vpl.get("break_even_lance", 0),

            # Reajuste
            "reajuste_pre_pct": params.get("reajuste_pre_pct", 0),
            "reajuste_pos_pct": params.get("reajuste_pos_pct", 0),
            "reajuste_pre_freq": params.get("reajuste_pre_freq", "Anual"),
            "reajuste_pos_freq": params.get("reajuste_pos_freq", "Anual"),
        }

    # -----------------------------------------------------------------
    # 11. CONSOLIDACAO DE COTAS (Multi-grupo)
    # -----------------------------------------------------------------

    def consolidar_cotas(self, grupos):
        """
        Consolidacao de multiplas cotas/grupos com medias ponderadas.

        Args:
            grupos: lista de dicts, cada um com:
                params: parametros do grupo
                peso: peso/quantidade (default 1)

        Returns:
            dict com fluxo consolidado e metricas ponderadas
        """
        if not grupos:
            return {}

        fluxos = []
        pesos = []

        for g in grupos:
            p = g.get("params", g)
            peso = g.get("peso", 1)
            fl = self.calcular_fluxo_completo(p)
            fluxos.append(fl)
            pesos.append(peso)

        peso_total = sum(pesos)

        # Determinar prazo maximo
        max_prazo = max(len(fl.get("cashflow", [])) for fl in fluxos)

        # Consolidar cashflows
        cf_consolidado = [0.0] * max_prazo
        for fl, peso in zip(fluxos, pesos):
            cf = fl.get("cashflow", [])
            for i in range(len(cf)):
                cf_consolidado[i] += cf[i] * peso

        # Metricas ponderadas
        total_pago = sum(
            fl.get("total_pago", fl.get("totais", {}).get("total_pago", 0)) * p
            for fl, p in zip(fluxos, pesos)
        )
        total_credito = sum(
            fl.get("carta_liquida", fl.get("totais", {}).get("carta_liquida", 0)) * p
            for fl, p in zip(fluxos, pesos)
        )

        tir_m = _irr(cf_consolidado, guess=0.005)
        tir_a = _annual_from_monthly(tir_m)

        return {
            "cashflow_consolidado": cf_consolidado,
            "total_pago": total_pago,
            "total_credito": total_credito,
            "tir_mensal": tir_m,
            "tir_anual": tir_a,
            "num_cotas": len(grupos),
            "peso_total": peso_total,
            "fluxos_individuais": fluxos,
            "custo_total_pct": (total_pago / total_credito * 100) if total_credito > 0 else 0,
        }


# =============================================================================
# FUNCOES STANDALONE (compatibilidade com nasa_engine.py)
# =============================================================================

_engine_default = None


def _get_engine():
    """Instancia singleton para funcoes standalone."""
    global _engine_default
    if _engine_default is None:
        _engine_default = NasaEngineHD()
    return _engine_default


def calcular_fluxo_consorcio(params):
    """
    Compatibilidade com nasa_engine.py.
    Aceita o mesmo formato de parametros e retorna resultado compativel.
    """
    # Mapeia parametros do formato antigo para novo
    p = {}
    if "valor_carta" in params:
        p["valor_credito"] = params["valor_carta"]
    if "valor_credito" in params:
        p["valor_credito"] = params["valor_credito"]

    p["prazo_meses"] = params.get("prazo_meses", 200)

    # Taxa adm: no formato antigo eh % total, no novo tambem
    if "taxa_adm" in params:
        p["taxa_adm_pct"] = params["taxa_adm"]
    if "taxa_adm_pct" in params:
        p["taxa_adm_pct"] = params["taxa_adm_pct"]

    if "fundo_reserva" in params:
        p["fundo_reserva_pct"] = params["fundo_reserva"]
    if "fundo_reserva_pct" in params:
        p["fundo_reserva_pct"] = params["fundo_reserva_pct"]

    p["momento_contemplacao"] = params.get("prazo_contemp", params.get("momento_contemplacao", 36))

    # Seguro
    if "seguro" in params:
        p["seguro_vida_pct"] = params["seguro"]
    if "seguro_vida_pct" in params:
        p["seguro_vida_pct"] = params["seguro_vida_pct"]

    # Lance
    p["lance_embutido_pct"] = params.get("lance_embutido_pct", 0)
    p["lance_livre_pct"] = params.get("lance_livre_pct", 0)

    # Reducao de parcela na fase 1 -> converter para periodos
    red_pct = params.get("parcela_red_pct", 100)
    contemp = p["momento_contemplacao"]
    prazo = p["prazo_meses"]

    if red_pct != 100:
        fc_pct_f1 = red_pct / 100.0
        # Ajustar para que o total = 100%
        # meses_f1 * fc_f1 + meses_f2 * fc_f2 = prazo (distribuicao linear)
        # Aqui simplificamos com 2 periodos
        p["periodos"] = [
            {"start": 1, "end": contemp, "fc_pct": fc_pct_f1, "ta_pct": 1.0, "fr_pct": 1.0},
            {"start": contemp + 1, "end": prazo, "fc_pct": 1.0, "ta_pct": 1.0, "fr_pct": 1.0},
        ]

    # Correcao anual -> reajuste
    corr = params.get("correcao_anual", 0)
    if corr > 0:
        p["reajuste_pre_pct"] = corr
        p["reajuste_pos_pct"] = corr
        p["reajuste_pre_freq"] = "Anual"
        p["reajuste_pos_freq"] = "Anual"

    # ALM/Hurdle
    p["alm_anual"] = params.get("alm_anual", 12.0)
    p["hurdle_anual"] = params.get("hurdle_anual", 12.0)

    engine = _get_engine()
    resultado = engine.calcular_fluxo_completo(p)

    # Garantir compatibilidade de saida
    # Converter fluxo para formato antigo
    fluxo_compat = []
    for f in resultado.get("fluxo", []):
        m = f.get("mes", 0)
        parcela = f.get("parcela_com_seguro", f.get("parcela_apos_reajuste", 0))
        lance = abs(f.get("lance_livre", 0))
        credito = f.get("credito_recebido", 0)

        fluxo_compat.append({
            "mes": m,
            "parcela": parcela,
            "fundo_comum": abs(f.get("amortizacao", 0)),
            "taxa_adm": f.get("valor_parcela_ta", 0),
            "fundo_reserva": f.get("valor_fundo_reserva", 0),
            "seguro": f.get("seguro_vida", 0),
            "lance": lance,
            "credito": credito,
            "fluxo_liquido": f.get("fluxo_caixa", 0),
            "fator_correcao": f.get("fator_reajuste", 1.0),
        })

    resultado["fluxo_mensal"] = fluxo_compat
    resultado["cashflow_consorcio"] = resultado.get("cashflow", [])
    resultado["parcela_f1_base"] = fluxo_compat[1]["parcela"] if len(fluxo_compat) > 1 else 0
    resultado["parcela_f2_base"] = fluxo_compat[contemp + 1]["parcela"] if len(fluxo_compat) > contemp + 1 else 0
    resultado["meses_restantes"] = prazo - contemp

    return resultado


def calcular_vpl_hd(params, fluxo_result):
    """Compatibilidade com nasa_engine.py."""
    engine = _get_engine()

    # Mapeia parametros
    p = dict(params)
    if "valor_carta" in p and "valor_credito" not in p:
        p["valor_credito"] = p["valor_carta"]
    if "prazo_contemp" in p and "momento_contemplacao" not in p:
        p["momento_contemplacao"] = p["prazo_contemp"]

    return engine.calcular_vpl_hd(p, fluxo_result)


def calcular_financiamento(valor, prazo_meses, taxa_mensal_pct, metodo="price"):
    """Compatibilidade com nasa_engine.py."""
    engine = _get_engine()
    params = {
        "valor": valor,
        "prazo_meses": prazo_meses,
        "taxa_mensal_pct": taxa_mensal_pct,
        "metodo": metodo,
        "calcular_iof": False,  # Manter comportamento original sem IOF
    }
    return engine.calcular_financiamento(params)


def comparar_consorcio_financiamento(params_consorcio, params_financ):
    """Compatibilidade com nasa_engine.py."""
    engine = _get_engine()

    # Mapeia params_financ do formato antigo
    pf = {}
    if "valor" in params_financ:
        pf["valor"] = params_financ["valor"]
    pf["prazo_meses"] = params_financ.get("prazo_meses", 0)
    pf["taxa_mensal_pct"] = params_financ.get("taxa_mensal_pct", 0)
    pf["metodo"] = params_financ.get("metodo", "price")
    pf["calcular_iof"] = False

    return engine.comparar_consorcio_financiamento(params_consorcio, pf)


# =============================================================================
# BLOCO DE TESTE
# =============================================================================

if __name__ == "__main__":
    print("=" * 70)
    print("NASA ENGINE HD - Teste de Validacao")
    print("=" * 70)

    # ---- Teste 1: Fluxo basico ----
    print("\n--- Teste 1: Fluxo Completo Basico ---")
    engine = NasaEngineHD()

    params_basico = {
        "valor_credito": 500000.0,
        "prazo_meses": 200,
        "taxa_adm_pct": 20.0,
        "fundo_reserva_pct": 3.0,
        "momento_contemplacao": 36,
        "lance_embutido_pct": 10.0,
        "lance_livre_pct": 15.0,
        "reajuste_pre_pct": 5.0,
        "reajuste_pos_pct": 5.0,
        "reajuste_pre_freq": "Anual",
        "reajuste_pos_freq": "Anual",
        "seguro_vida_pct": 0.03,
        "seguro_vida_inicio": 1,
        "alm_anual": 12.0,
        "hurdle_anual": 12.0,
        "periodos": [
            {"start": 1, "end": 36, "fc_pct": 0.50, "ta_pct": 1.0, "fr_pct": 1.0},
            {"start": 37, "end": 200, "fc_pct": 1.00, "ta_pct": 1.0, "fr_pct": 1.0},
        ],
        "custos_acessorios": [
            {"descricao": "Avaliacao", "valor": 5000.0, "momento": 1},
            {"descricao": "Documentacao", "valor": 3000.0, "momento": 36},
        ],
    }

    fluxo = engine.calcular_fluxo_completo(params_basico)
    totais = fluxo["totais"]
    metricas = fluxo["metricas"]

    print(f"  Valor credito:       R$ {params_basico['valor_credito']:>15,.2f}")
    print(f"  Carta liquida:       R$ {totais['carta_liquida']:>15,.2f}")
    print(f"  Lance embutido:      R$ {totais['lance_embutido_valor']:>15,.2f}")
    print(f"  Lance livre:         R$ {totais['lance_livre_valor']:>15,.2f}")
    print(f"  Total pago:          R$ {totais['total_pago']:>15,.2f}")
    print(f"  Custo total:            {metricas['custo_total_pct']:>14.2f}%")
    print(f"  Parcela media:       R$ {metricas['parcela_media']:>15,.2f}")
    print(f"  TIR mensal:             {metricas['tir_mensal']*100:>14.4f}%")
    print(f"  TIR anual:              {metricas['tir_anual']*100:>14.2f}%")

    # Mostra primeiras e ultimas parcelas
    parcelas = engine.gerar_parcelas(fluxo)
    print(f"\n  Primeiras 3 parcelas:")
    for p in parcelas[:3]:
        print(f"    Mes {p['mes']:>3}: Parcela R$ {p['parcela_total']:>12,.2f}  "
              f"(FC: {p['fundo_comum']:>10,.2f}  TA: {p['taxa_adm']:>8,.2f}  "
              f"FR: {p['fundo_reserva']:>7,.2f}  Seg: {p['seguro']:>7,.2f})")
    print(f"  Ultimas 3 parcelas:")
    for p in parcelas[-3:]:
        print(f"    Mes {p['mes']:>3}: Parcela R$ {p['parcela_total']:>12,.2f}  "
              f"(FC: {p['fundo_comum']:>10,.2f}  TA: {p['taxa_adm']:>8,.2f}  "
              f"FR: {p['fundo_reserva']:>7,.2f}  Seg: {p['seguro']:>7,.2f})")

    # ---- Teste 2: VPL HD ----
    print("\n--- Teste 2: VPL HD (Taxas Duais) ---")
    vpl = engine.calcular_vpl_hd(params_basico, fluxo)
    print(f"  B0 (PV credito):     R$ {vpl['b0']:>15,.2f}")
    print(f"  H0 (PV pgtos pre-T): R$ {vpl['h0']:>15,.2f}")
    print(f"  D0 (valor criado):   R$ {vpl['d0']:>15,.2f}")
    print(f"  PV pos-T:            R$ {vpl['pv_pos_t']:>15,.2f}")
    print(f"  Delta VPL:           R$ {vpl['delta_vpl']:>15,.2f}")
    print(f"  Cria valor:             {'SIM' if vpl['cria_valor'] else 'NAO':>14}")
    print(f"  Break-even lance:       {vpl['break_even_lance']:>14.2f}%")

    # ---- Teste 3: Financiamento ----
    print("\n--- Teste 3: Financiamento (Price com IOF) ---")
    params_fin = {
        "valor": 500000.0,
        "prazo_meses": 200,
        "taxa_mensal_pct": 0.8,
        "metodo": "price",
        "calcular_iof": True,
    }
    fin = engine.calcular_financiamento(params_fin)
    print(f"  Valor financiado:    R$ {fin['valor']:>15,.2f}")
    print(f"  Total pago:          R$ {fin['total_pago']:>15,.2f}")
    print(f"  Total juros:         R$ {fin['total_juros']:>15,.2f}")
    print(f"  IOF:                 R$ {fin['iof']:>15,.2f}")
    print(f"  CET anual:              {fin['cet_anual']*100:>14.2f}%")
    print(f"  1a parcela:          R$ {fin['parcelas'][0]['parcela']:>15,.2f}")
    print(f"  Ultima parcela:      R$ {fin['parcelas'][-1]['parcela']:>15,.2f}")

    # ---- Teste 4: Comparativo ----
    print("\n--- Teste 4: Comparativo Consorcio vs Financiamento ---")
    comp = engine.comparar_consorcio_financiamento(params_basico, params_fin)
    print(f"  Total consorcio:     R$ {comp['total_pago_consorcio']:>15,.2f}")
    print(f"  Total financiamento: R$ {comp['total_pago_financiamento']:>15,.2f}")
    print(f"  Economia nominal:    R$ {comp['economia_nominal']:>15,.2f}")
    print(f"  VPL consorcio:       R$ {comp['vpl_consorcio']:>15,.2f}")
    print(f"  VPL financiamento:   R$ {comp['vpl_financiamento']:>15,.2f}")
    print(f"  TIR cons. anual:        {comp['tir_consorcio_anual']*100:>14.2f}%")
    print(f"  TIR fin.  anual:        {comp['tir_financ_anual']*100:>14.2f}%")

    # ---- Teste 5: Venda da Operacao ----
    print("\n--- Teste 5: Venda da Operacao ---")
    venda = engine.calcular_venda_operacao(fluxo, {
        "momento_venda": 60,
        "valor_venda": 300000.0,
        "tma": 0.01,
    })
    print(f"  Total investido:     R$ {venda['total_investido']:>15,.2f}")
    print(f"  Valor venda:         R$ {venda['valor_venda']:>15,.2f}")
    print(f"  Ganho nominal:       R$ {venda['ganho_nominal']:>15,.2f}")
    print(f"  Ganho %:                {venda['ganho_pct']:>14.2f}%")
    print(f"  TIR vendedor a.a.:      {venda['tir_vendedor_anual']*100:>14.2f}%")
    print(f"  TIR comprador a.a.:     {venda['tir_comprador_anual']*100:>14.2f}%")

    # ---- Teste 6: Credito para Lance ----
    print("\n--- Teste 6: Credito para Lance ---")
    lance_fin = engine.calcular_credito_lance({
        "valor_lance": 75000.0,
        "prazo_meses": 60,
        "taxa_mensal_pct": 1.2,
        "metodo": "price",
        "carencia": 3,
        "tac": 1500.0,
        "avaliacao_garantia": 800.0,
        "comissao": 500.0,
        "calcular_iof": True,
    })
    print(f"  Valor lance:         R$ {lance_fin['valor']:>15,.2f}")
    print(f"  Total pago:          R$ {lance_fin['total_pago']:>15,.2f}")
    print(f"  IOF:                 R$ {lance_fin['iof']:>15,.2f}")
    print(f"  Custos iniciais:     R$ {lance_fin['custos_iniciais']:>15,.2f}")
    print(f"  CET anual:              {lance_fin['cet_anual']*100:>14.2f}%")

    # ---- Teste 7: Custo Combinado ----
    print("\n--- Teste 7: Custo Combinado ---")
    combinado = engine.calcular_custo_combinado(fluxo, lance_fin)
    print(f"  Total consorcio:     R$ {combinado['total_pago_consorcio']:>15,.2f}")
    print(f"  Total lance fin.:    R$ {combinado['total_pago_lance']:>15,.2f}")
    print(f"  Total combinado:     R$ {combinado['total_pago_combinado']:>15,.2f}")
    print(f"  TIR combinada a.a.:     {combinado['tir_anual_combinado']*100:>14.2f}%")

    # ---- Teste 8: Cenarios ----
    print("\n--- Teste 8: Cenarios ---")
    sc = engine.scenarios
    idx1 = sc.save_scenario("Base", params_basico, {
        "total_pago": totais["total_pago"],
        "tir_mensal": metricas["tir_mensal"],
        "tir_anual": metricas["tir_anual"],
        "delta_vpl": vpl["delta_vpl"],
        "parcela_media": metricas["parcela_media"],
    })

    params2 = dict(params_basico)
    params2["lance_livre_pct"] = 25.0
    fluxo2 = engine.calcular_fluxo_completo(params2)
    vpl2 = engine.calcular_vpl_hd(params2, fluxo2)
    idx2 = sc.save_scenario("Lance 25%", params2, {
        "total_pago": fluxo2["totais"]["total_pago"],
        "tir_mensal": fluxo2["metricas"]["tir_mensal"],
        "tir_anual": fluxo2["metricas"]["tir_anual"],
        "delta_vpl": vpl2["delta_vpl"],
        "parcela_media": fluxo2["metricas"]["parcela_media"],
    })

    print(f"  Cenarios salvos: {sc.list_scenarios()}")
    comp_sc = sc.compare_scenarios([idx1, idx2])
    for idx, data in comp_sc.items():
        print(f"  [{idx}] {data['name']}: Total R$ {data['total_pago']:,.2f}, "
              f"Delta VPL R$ {data['delta_vpl']:,.2f}")

    # ---- Teste 9: Compatibilidade com nasa_engine.py ----
    print("\n--- Teste 9: Compatibilidade API Antiga ---")
    params_antigo = {
        "valor_carta": 500000.0,
        "prazo_meses": 200,
        "taxa_adm": 20.0,
        "fundo_reserva": 3.0,
        "seguro": 0.03,
        "prazo_contemp": 36,
        "parcela_red_pct": 70.0,
        "lance_livre_pct": 15.0,
        "lance_embutido_pct": 10.0,
        "correcao_anual": 5.0,
        "alm_anual": 12.0,
        "hurdle_anual": 12.0,
    }
    fluxo_compat = calcular_fluxo_consorcio(params_antigo)
    print(f"  Total pago:          R$ {fluxo_compat['total_pago']:>15,.2f}")
    print(f"  Carta liquida:       R$ {fluxo_compat['carta_liquida']:>15,.2f}")
    print(f"  Parcela F1 (base):   R$ {fluxo_compat['parcela_f1_base']:>15,.2f}")
    print(f"  Parcela F2 (base):   R$ {fluxo_compat['parcela_f2_base']:>15,.2f}")

    vpl_compat = calcular_vpl_hd(params_antigo, fluxo_compat)
    print(f"  Delta VPL:           R$ {vpl_compat['delta_vpl']:>15,.2f}")

    fin_compat = calcular_financiamento(500000, 200, 0.8)
    print(f"  Fin. total pago:     R$ {fin_compat['total_pago']:>15,.2f}")

    # ---- Teste 10: Resumo Cliente ----
    print("\n--- Teste 10: Resumo Cliente ---")
    resumo = engine.gerar_resumo_cliente(params_basico, fluxo, vpl)
    for k, v in resumo.items():
        if isinstance(v, float):
            print(f"  {k:30s}: {v:>15,.4f}")
        else:
            print(f"  {k:30s}: {v}")

    # ---- Teste 11: GoalSeek ----
    print("\n--- Teste 11: GoalSeek ---")
    # Encontra x tal que x^2 = 144
    result = goal_seek(lambda x: x ** 2, 144, 10.0, lo=0.0, hi=100.0)
    print(f"  x^2 = 144  =>  x = {result:.6f}  (esperado: 12.0)")

    # Encontra taxa mensal equivalente a 12% a.a.
    result2 = goal_seek(lambda r: (1 + r) ** 12 - 1, 0.12, 0.01, lo=0.0, hi=0.1)
    print(f"  (1+r)^12-1 = 0.12  =>  r = {result2*100:.6f}%  "
          f"(esperado: {_monthly_from_annual(0.12)*100:.6f}%)")

    # ---- Teste 12: Credito Equivalente ----
    print("\n--- Teste 12: Credito Equivalente ---")
    cred_eq = engine.calcular_credito_equivalente(params_basico)
    print(f"  Credito equivalente: R$ {cred_eq:>15,.2f}")

    # ---- Teste 13: Consolidacao de Cotas ----
    print("\n--- Teste 13: Consolidacao de Cotas ---")
    params_g2 = dict(params_basico)
    params_g2["valor_credito"] = 300000.0
    params_g2["prazo_meses"] = 180

    consolidado = engine.consolidar_cotas([
        {"params": params_basico, "peso": 2},
        {"params": params_g2, "peso": 1},
    ])
    print(f"  Cotas consolidadas:     {consolidado['num_cotas']}")
    print(f"  Total credito:       R$ {consolidado['total_credito']:>15,.2f}")
    print(f"  Total pago:          R$ {consolidado['total_pago']:>15,.2f}")
    print(f"  TIR anual:              {consolidado['tir_anual']*100:>14.2f}%")

    print("\n" + "=" * 70)
    print("Todos os testes concluidos com sucesso!")
    print("=" * 70)
