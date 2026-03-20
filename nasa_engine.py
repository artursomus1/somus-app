"""
NASA Engine - Motor de Calculo VPL para Consorcios
Baseado na planilha NASA NOVA HD VPL da Somus Capital.

Calcula: fluxo de caixa mensal, VPL, TIR/IRR, CET, Delta VPL,
break-even lance, e comparativo consorcio vs financiamento.
"""

import math


# ─── helpers ────────────────────────────────────────────────────────────────

def _npv(rate_monthly, cashflows):
    """Valor Presente Liquido de uma serie de fluxos mensais."""
    if rate_monthly == 0:
        return sum(cashflows)
    total = 0.0
    for t, cf in enumerate(cashflows):
        total += cf / (1 + rate_monthly) ** t
    return total


def _irr(cashflows, guess=0.01, tol=1e-9, max_iter=500):
    """Taxa Interna de Retorno (mensal) via Newton-Raphson."""
    rate = guess
    for _ in range(max_iter):
        npv_val = 0.0
        dnpv = 0.0
        for t, cf in enumerate(cashflows):
            d = (1 + rate) ** t
            npv_val += cf / d
            if t > 0:
                dnpv -= t * cf / ((1 + rate) ** (t + 1))
        if abs(npv_val) < tol:
            return rate
        if abs(dnpv) < 1e-15:
            break
        rate -= npv_val / dnpv
        if rate <= -1:
            rate = 0.0001
    return rate


def _annual_from_monthly(r_m):
    """Converte taxa mensal em anual equivalente."""
    return (1 + r_m) ** 12 - 1


def _monthly_from_annual(r_a):
    """Converte taxa anual em mensal equivalente."""
    if r_a <= -1:
        return 0.0
    return (1 + r_a) ** (1 / 12) - 1


# ─── Fluxo de Caixa do Consorcio ────────────────────────────────────────────

def calcular_fluxo_consorcio(params):
    """
    Constroi o fluxo de caixa mensal do consorcio.

    params (dict):
        valor_carta      : valor da carta de credito
        prazo_meses      : prazo total em meses
        taxa_adm         : % taxa de administracao
        fundo_reserva    : % fundo de reserva
        seguro           : % seguro prestamista
        prazo_contemp    : mes da contemplacao
        parcela_red_pct  : % do fundo comum pago na fase 1 (ex: 70 = reducao 30%)
        lance_livre_pct  : % lance livre sobre a carta
        lance_embutido_pct: % lance embutido sobre a carta
        correcao_anual   : % correcao anual das parcelas

    Retorna dict com:
        fluxo_mensal     : lista de dicts com detalhes de cada mes
        cashflow_consorcio: lista de floats (fluxo para calculo de TIR/VPL)
        total_pago       : total desembolsado
        carta_liquida    : credito recebido apos lance embutido
    """
    vc = params["valor_carta"]
    prazo = params["prazo_meses"]
    taxa_adm = params["taxa_adm"] / 100
    fundo_res = params["fundo_reserva"] / 100
    seguro = params["seguro"] / 100
    contemp = params["prazo_contemp"]
    red_pct = params.get("parcela_red_pct", 100) / 100
    lance_livre = params["lance_livre_pct"] / 100
    lance_emb = params["lance_embutido_pct"] / 100
    corr = params.get("correcao_anual", 0) / 100

    # Valores base mensais
    fc_integral = vc / prazo
    ta_mensal = (vc * taxa_adm) / prazo
    fr_mensal = (vc * fundo_res) / prazo
    sg_mensal = (vc * seguro) / prazo
    taxas = ta_mensal + fr_mensal + sg_mensal

    # Fase 1
    fc_reduzido = fc_integral * red_pct
    parcela_f1 = fc_reduzido + taxas
    fundo_pago_f1 = fc_reduzido * contemp

    # Lances
    lance_livre_val = vc * lance_livre
    lance_emb_val = vc * lance_emb
    lance_total = lance_livre_val + lance_emb_val
    carta_liq = vc - lance_emb_val

    # Fase 2
    meses_rest = prazo - contemp
    fundo_rest = max(0, vc - fundo_pago_f1 - lance_total)
    fc_f2 = fundo_rest / meses_rest if meses_rest > 0 else 0
    parcela_f2 = fc_f2 + taxas

    # Construir fluxo mes a mes
    fluxo = []
    cashflow = []

    # Mes 0: recebimento do credito (na contemplacao, mas para VPL trazemos a t=0)
    # Na verdade, o credito e recebido no mes da contemplacao
    # Fluxo: mes 0 = 0, ..., mes contemp = +carta_liq, parcelas sao negativas

    total_pago = 0.0

    for mes in range(prazo + 1):
        fator_corr = (1 + corr) ** ((mes - 1) // 12) if mes > 0 else 1.0

        if mes == 0:
            # Mes 0: sem pagamento, sem credito
            fluxo.append({
                "mes": 0, "parcela": 0, "fundo_comum": 0,
                "taxa_adm": 0, "fundo_reserva": 0, "seguro": 0,
                "lance": 0, "credito": 0, "fluxo_liquido": 0,
                "fator_correcao": 1.0,
            })
            cashflow.append(0.0)
        elif mes <= contemp:
            # Fase 1: parcelas pre-contemplacao
            p = parcela_f1 * fator_corr
            lance_mes = lance_livre_val if mes == contemp else 0
            credito_mes = carta_liq if mes == contemp else 0

            total_pago += p + lance_mes
            fluxo.append({
                "mes": mes, "parcela": p,
                "fundo_comum": fc_reduzido * fator_corr,
                "taxa_adm": ta_mensal * fator_corr,
                "fundo_reserva": fr_mensal * fator_corr,
                "seguro": sg_mensal * fator_corr,
                "lance": lance_mes, "credito": credito_mes,
                "fluxo_liquido": credito_mes - p - lance_mes,
                "fator_correcao": fator_corr,
            })
            cashflow.append(credito_mes - p - lance_mes)
        else:
            # Fase 2: parcelas pos-contemplacao
            p = parcela_f2 * fator_corr
            total_pago += p
            fluxo.append({
                "mes": mes, "parcela": p,
                "fundo_comum": fc_f2 * fator_corr,
                "taxa_adm": ta_mensal * fator_corr,
                "fundo_reserva": fr_mensal * fator_corr,
                "seguro": sg_mensal * fator_corr,
                "lance": 0, "credito": 0,
                "fluxo_liquido": -p,
                "fator_correcao": fator_corr,
            })
            cashflow.append(-p)

    return {
        "fluxo_mensal": fluxo,
        "cashflow_consorcio": cashflow,
        "total_pago": total_pago,
        "carta_liquida": carta_liq,
        "lance_livre_valor": lance_livre_val,
        "lance_embutido_valor": lance_emb_val,
        "parcela_f1_base": parcela_f1,
        "parcela_f2_base": parcela_f2,
        "meses_restantes": meses_rest,
    }


# ─── Analise VPL HD (estilo NASA) ───────────────────────────────────────────

def calcular_vpl_hd(params, fluxo_result):
    """
    Analise VPL HD (Somus Rota A / Goal-Based).

    Desconta parcelas pre-contemplacao a taxa ALM (custo de oportunidade),
    parcelas pos-contemplacao a taxa Hurdle.

    params deve conter adicionalmente:
        alm_anual    : CDI liquido a.a. (ex: 12.0 para 12%)
        hurdle_anual : taxa hurdle a.a. (ex: 12.0 para 12%)

    Retorna dict com:
        b0             : PV do credito a taxa ALM
        h0             : PV dos pagamentos pre-contemplacao + lance
        d0             : b0 - h0 (valor criado antes das parcelas pos-T)
        pv_pos_t       : PV das parcelas pos-contemplacao a taxa hurdle
        delta_vpl      : d0 - pv_pos_t (resultado final)
        cria_valor     : True se delta_vpl >= 0
        break_even_lance: % lance livre para VPL = 0 (aproximado)
        tir_mensal     : TIR mensal do fluxo
        tir_anual      : TIR anual
        cet_anual      : CET anual efetivo
        vpl_total      : VPL total do fluxo a taxa ALM
    """
    alm_m = _monthly_from_annual(params.get("alm_anual", 12.0) / 100)
    hurdle_m = _monthly_from_annual(params.get("hurdle_anual", 12.0) / 100)
    contemp = params["prazo_contemp"]
    carta_liq = fluxo_result["carta_liquida"]
    fluxo = fluxo_result["fluxo_mensal"]

    # B0: PV do credito recebido na contemplacao, descontado a ALM ate t=0
    b0 = carta_liq / (1 + alm_m) ** contemp

    # H0: PV dos pagamentos pre-contemplacao (incluindo lance livre)
    h0 = 0.0
    for f in fluxo:
        if 0 < f["mes"] <= contemp:
            pagamento = f["parcela"] + f["lance"]
            h0 += pagamento / (1 + alm_m) ** f["mes"]

    # D0: valor criado antes das parcelas pos-T
    d0 = b0 - h0

    # PV parcelas pos-contemplacao descontadas a taxa hurdle
    # (trazidas ao mes da contemplacao, depois ao t=0)
    pv_pos_t_at_contemp = 0.0
    for f in fluxo:
        if f["mes"] > contemp:
            meses_apos = f["mes"] - contemp
            pv_pos_t_at_contemp += f["parcela"] / (1 + hurdle_m) ** meses_apos

    # Trazer ao t=0
    pv_pos_t = pv_pos_t_at_contemp / (1 + alm_m) ** contemp

    # Delta VPL
    delta_vpl = d0 - pv_pos_t
    cria_valor = delta_vpl >= 0

    # TIR do fluxo completo
    cf = fluxo_result["cashflow_consorcio"]
    tir_m = _irr(cf, guess=0.008)
    tir_a = _annual_from_monthly(tir_m)

    # CET = TIR do fluxo (custo efetivo total)
    cet_anual = tir_a

    # VPL total a taxa ALM
    vpl_total = _npv(alm_m, cf)

    # Break-even lance (busca binaria)
    be_lance = _buscar_break_even_lance(params, alm_m, hurdle_m)

    return {
        "b0": b0,
        "h0": h0,
        "d0": d0,
        "pv_pos_t": pv_pos_t,
        "delta_vpl": delta_vpl,
        "cria_valor": cria_valor,
        "break_even_lance": be_lance,
        "tir_mensal": tir_m,
        "tir_anual": tir_a,
        "cet_anual": cet_anual,
        "vpl_total": vpl_total,
    }


def _buscar_break_even_lance(params, alm_m, hurdle_m):
    """Busca binaria para encontrar o lance livre que zera o Delta VPL."""
    lo, hi = 0.0, 90.0
    best = 0.0

    for _ in range(60):
        mid = (lo + hi) / 2
        test_params = dict(params)
        test_params["lance_livre_pct"] = mid

        fluxo_r = calcular_fluxo_consorcio(test_params)
        contemp = test_params["prazo_contemp"]
        carta_liq = fluxo_r["carta_liquida"]
        fluxo = fluxo_r["fluxo_mensal"]

        b0 = carta_liq / (1 + alm_m) ** contemp

        h0 = 0.0
        for f in fluxo:
            if 0 < f["mes"] <= contemp:
                h0 += (f["parcela"] + f["lance"]) / (1 + alm_m) ** f["mes"]

        d0 = b0 - h0

        pv_pos = 0.0
        for f in fluxo:
            if f["mes"] > contemp:
                pv_pos += f["parcela"] / (1 + hurdle_m) ** (f["mes"] - contemp)
        pv_pos /= (1 + alm_m) ** contemp

        delta = d0 - pv_pos

        if abs(delta) < 1:
            return mid

        if delta > 0:
            # Valor positivo = lance pode ser maior
            lo = mid
        else:
            hi = mid
        best = mid

    return best


# ─── Financiamento (Price / SAC) ─────────────────────────────────────────────

def calcular_financiamento(valor, prazo_meses, taxa_mensal_pct, metodo="price"):
    """
    Simula uma operacao de credito (financiamento).

    valor          : valor financiado
    prazo_meses    : prazo em meses
    taxa_mensal_pct: taxa de juros mensal em %
    metodo         : 'price' ou 'sac'

    Retorna dict com:
        parcelas       : lista de dicts (mes, parcela, juros, amortizacao, saldo)
        cashflow       : lista de floats para calculo de VPL
        total_pago     : total desembolsado
        total_juros    : total de juros pagos
    """
    r = taxa_mensal_pct / 100
    saldo = valor
    parcelas = []
    cashflow = [valor]  # mes 0: recebe o valor
    total_pago = 0.0

    if metodo.lower() == "price":
        if r > 0:
            pmt = valor * r * (1 + r) ** prazo_meses / ((1 + r) ** prazo_meses - 1)
        else:
            pmt = valor / prazo_meses

        for mes in range(1, prazo_meses + 1):
            juros = saldo * r
            amort = pmt - juros
            saldo -= amort
            saldo = max(0, saldo)
            total_pago += pmt
            parcelas.append({
                "mes": mes, "parcela": pmt, "juros": juros,
                "amortizacao": amort, "saldo": saldo,
            })
            cashflow.append(-pmt)
    else:  # SAC
        amort_fixa = valor / prazo_meses
        for mes in range(1, prazo_meses + 1):
            juros = saldo * r
            pmt = amort_fixa + juros
            saldo -= amort_fixa
            saldo = max(0, saldo)
            total_pago += pmt
            parcelas.append({
                "mes": mes, "parcela": pmt, "juros": juros,
                "amortizacao": amort_fixa, "saldo": saldo,
            })
            cashflow.append(-pmt)

    total_juros = total_pago - valor

    return {
        "parcelas": parcelas,
        "cashflow": cashflow,
        "total_pago": total_pago,
        "total_juros": total_juros,
        "valor": valor,
    }


# ─── Comparativo Consorcio vs Financiamento ──────────────────────────────────

def comparar_consorcio_financiamento(params_consorcio, params_financ):
    """
    Compara consorcio vs financiamento lado a lado.

    params_consorcio : dict com parametros do consorcio (mesmo formato de calcular_fluxo_consorcio)
    params_financ    : dict com {valor, prazo_meses, taxa_mensal_pct, metodo}

    Retorna dict com:
        consorcio      : resultado do fluxo consorcio
        financiamento  : resultado do financiamento
        vpl_consorcio  : VPL do consorcio
        vpl_financ     : VPL do financiamento
        economia       : diferenca VPL (positivo = consorcio mais barato)
        tir_consorcio  : TIR do consorcio
        tir_financ     : TIR do financiamento
        razao_vpl_cons : total_pago / credito (consorcio)
        razao_vpl_fin  : total_pago / credito (financiamento)
    """
    alm_m = _monthly_from_annual(params_consorcio.get("alm_anual", 12.0) / 100)

    # Consorcio
    fluxo_c = calcular_fluxo_consorcio(params_consorcio)
    cf_c = fluxo_c["cashflow_consorcio"]
    vpl_c = _npv(alm_m, cf_c)
    tir_c = _irr(cf_c, guess=0.008)

    # Financiamento
    fin = calcular_financiamento(
        params_financ["valor"],
        params_financ["prazo_meses"],
        params_financ["taxa_mensal_pct"],
        params_financ.get("metodo", "price"),
    )
    cf_f = fin["cashflow"]
    vpl_f = _npv(alm_m, cf_f)
    tir_f = _irr(cf_f, guess=0.008)

    carta_liq = fluxo_c["carta_liquida"]
    razao_c = fluxo_c["total_pago"] / carta_liq if carta_liq > 0 else 0
    razao_f = fin["total_pago"] / fin["valor"] if fin["valor"] > 0 else 0

    return {
        "consorcio": fluxo_c,
        "financiamento": fin,
        "vpl_consorcio": vpl_c,
        "vpl_financiamento": vpl_f,
        "economia_vpl": vpl_c - vpl_f,
        "tir_consorcio_mensal": tir_c,
        "tir_consorcio_anual": _annual_from_monthly(tir_c),
        "tir_financ_mensal": tir_f,
        "tir_financ_anual": _annual_from_monthly(tir_f),
        "razao_vpl_consorcio": razao_c,
        "razao_vpl_financ": razao_f,
    }
