"""
Gerador de PDF - Simulação de Consórcio
Gera documento PDF com simulação detalhada de consórcio em duas fases.
Suporta: parcela reduzida, lance livre, lance embutido.
Padrao Somus Capital.
"""

import os
from datetime import datetime
from fpdf import FPDF

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo_somus.png")

# Cores Somus Capital
GREEN = (0, 77, 51)
GREEN_LIGHT = (0, 92, 61)
BLUE = (24, 99, 220)
WHITE = (255, 255, 255)
LIGHT_GRAY = (246, 247, 249)
DARK_GRAY = (45, 45, 45)
MID_GRAY = (180, 180, 180)
TEXT_GRAY = (80, 80, 80)
ORANGE = (230, 131, 42)
PURPLE = (140, 40, 140)
RED = (180, 40, 40)
TEAL = (0, 130, 90)


def sanitize_text(text):
    if text is None:
        return ""
    text = str(text)
    replacements = {
        '\u2013': '-', '\u2014': '-', '\u2018': "'", '\u2019': "'",
        '\u201c': '"', '\u201d': '"', '\u2026': '...', '\u2022': '-',
        '\u00ba': 'o', '\u00aa': 'a', '\u00a0': ' ',
        '\ufeff': '', '\u200b': '', '\u200e': '', '\u200f': '',
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    try:
        text.encode('latin-1')
    except UnicodeEncodeError:
        cleaned = []
        for ch in text:
            try:
                ch.encode('latin-1')
                cleaned.append(ch)
            except UnicodeEncodeError:
                cleaned.append('?')
        text = ''.join(cleaned)
    return text


def fmt_currency(value):
    if value is None or value == 0:
        return "R$ 0,00"
    try:
        v = float(value)
    except (ValueError, TypeError):
        return "R$ 0,00"
    neg = v < 0
    v = abs(v)
    inteiro = int(v)
    centavos = round((v - inteiro) * 100)
    if centavos >= 100:
        inteiro += 1
        centavos = 0
    s_int = f"{inteiro:,}".replace(",", ".")
    result = f"R$ {s_int},{centavos:02d}"
    return f"-{result}" if neg else result


def fmt_pct(value):
    return f"{value:.2f}%".replace(".", ",")


def _calc_custo_efetivo(carta_liq, total_pago, prazo):
    """Calcula a taxa de custo efetivo do consórcio (composta).

    Encontra a taxa mensal composta equivalente ao custo total da operação:
    taxa_mensal = (total_pago / carta_liq)^(1/prazo) - 1
    Retorna (taxa_mensal, taxa_anual, custo_total_pct).
    """
    if carta_liq <= 0 or prazo <= 0 or total_pago <= carta_liq:
        return 0.0, 0.0, 0.0

    custo_pct = (total_pago - carta_liq) / carta_liq * 100
    taxa_mensal = (total_pago / carta_liq) ** (1 / prazo) - 1
    taxa_anual = (1 + taxa_mensal) ** 12 - 1
    return taxa_mensal, taxa_anual, custo_pct


def _calc(d):
    """Calcula todas as variáveis a partir do dict de dados."""
    vc = d.get("valor_carta", 0)
    pz = d.get("prazo_meses", 1)
    ta_pct = d.get("taxa_adm", 0) / 100
    fr_pct = d.get("fundo_reserva", 0) / 100
    sg_pct = d.get("seguro", 0) / 100
    pc = d.get("prazo_contemplação", pz)
    pr = d.get("parcela_reduzida_pct", 100) / 100
    ll_pct = d.get("lance_livre_pct", 0) / 100
    le_pct = d.get("lance_embutido_pct", 0) / 100

    fc_integral = vc / pz
    ta = (vc * ta_pct) / pz
    fr = (vc * fr_pct) / pz
    sg = (vc * sg_pct) / pz
    taxas = ta + fr + sg

    fc_red = fc_integral * pr
    p1 = fc_red + taxas
    fundo_pago_f1 = fc_red * pc

    ll_val = vc * ll_pct
    le_val = vc * le_pct
    lance_total = ll_val + le_val
    carta_liq = vc - le_val

    mr = pz - pc
    fundo_rest = max(0, vc - fundo_pago_f1 - lance_total)
    fc_f2 = fundo_rest / mr if mr > 0 else 0
    p2 = fc_f2 + taxas

    desemb1 = p1 * pc
    desemb2 = p2 * mr if mr > 0 else 0
    total_des = desemb1 + ll_val + desemb2
    total_tax = taxas * pz

    # Correção anual
    corr_anual = d.get("correção_anual", 0) / 100
    tipo_corr = d.get("tipo_correção", "Pós-fixado")
    índice_corr = d.get("índice_correção", "INCC")
    tem_correção = corr_anual > 0

    desemb1_corr = 0.0
    desemb2_corr = 0.0
    p1_final = p1
    p2_final = p2

    for mes in range(1, pz + 1):
        fator = (1 + corr_anual) ** ((mes - 1) // 12)
        if mes <= pc:
            p_corr = p1 * fator
            desemb1_corr += p_corr
            p1_final = p_corr
        else:
            p_corr = p2 * fator
            desemb2_corr += p_corr
            p2_final = p_corr

    total_des_corr = desemb1_corr + ll_val + desemb2_corr

    # === CUSTO EFETIVO ===
    ce_mensal, ce_anual, ce_pct = _calc_custo_efetivo(carta_liq, total_des_corr, pz)

    # === RESUMO DA OPERAÇÃO ===
    custo_total = total_des_corr - carta_liq if tem_correção else total_des - carta_liq
    total_ref = total_des_corr if tem_correção else total_des
    relação_custo = (total_ref / carta_liq) if carta_liq > 0 else 0
    parcela_media = total_ref / pz if pz > 0 else 0

    return {
        "vc": vc, "pz": pz, "pc": pc, "mr": mr, "pr": pr,
        "fc_integral": fc_integral, "fc_red": fc_red, "fc_f2": fc_f2,
        "ta": ta, "fr": fr, "sg": sg, "taxas": taxas,
        "p1": p1, "p2": p2,
        "ll_val": ll_val, "le_val": le_val, "lance_total": lance_total,
        "carta_liq": carta_liq,
        "desemb1": desemb1, "desemb2": desemb2,
        "total_des": total_des, "total_tax": total_tax,
        "corr_anual": corr_anual, "tipo_corr": tipo_corr, "índice_corr": índice_corr,
        "tem_correção": tem_correção,
        "desemb1_corr": desemb1_corr, "desemb2_corr": desemb2_corr,
        "total_des_corr": total_des_corr,
        "p1_final": p1_final, "p2_final": p2_final,
        "ce_mensal": ce_mensal, "ce_anual": ce_anual, "ce_pct": ce_pct,
        "custo_total": custo_total, "relação_custo": relação_custo,
        "parcela_media": parcela_media,
    }


class ConsorcioPDF(FPDF):

    def __init__(self, dados):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.dados = dados
        self.c = _calc(dados)
        self.set_auto_page_break(auto=True, margin=20)
        self.logo_path = LOGO_PATH if os.path.exists(LOGO_PATH) else None

    def header(self):
        self.set_fill_color(*GREEN)
        self.rect(0, 0, 210, 18, "F")

        if self.logo_path:
            try:
                self.image(self.logo_path, x=8, y=3, h=12)
            except Exception:
                self.set_font("Helvetica", "B", 12)
                self.set_text_color(*WHITE)
                self.set_xy(8, 4)
                self.cell(50, 10, "SOMUS CAPITAL")

        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*WHITE)
        self.set_xy(0, 4)
        self.cell(210, 10, "SIMULAÇÃO DE CONSÓRCIO", align="C")

        self.set_fill_color(*BLUE)
        self.rect(0, 18, 210, 1.2, "F")
        self.set_y(23)

    def footer(self):
        self.set_y(-14)
        self.set_draw_color(*MID_GRAY)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(1.5)
        self.set_font("Helvetica", "", 6)
        self.set_text_color(140, 140, 140)
        self.cell(0, 5, "Somus Capital  |  Simulação de Consórcio", align="L")
        self.cell(0, 5, f"Página {self.page_no()}/{{nb}}", align="R")

    def _section_title(self, title):
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*GREEN)
        self.cell(0, 8, sanitize_text(title), new_x="LMARGIN", new_y="NEXT")

    def _param_row(self, label, value, idx):
        self.set_fill_color(*(LIGHT_GRAY if idx % 2 == 0 else WHITE))
        self.set_x(10)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*TEXT_GRAY)
        self.cell(80, 7, f"  {label}", fill=True)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*DARK_GRAY)
        self.cell(110, 7, f"  {value}", fill=True)
        self.ln(7)

    def _color_bar_row(self, label, value, color):
        y = self.get_y()
        self.set_fill_color(*color)
        self.rect(10, y, 2, 6.5, "F")
        self.set_x(15)
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*TEXT_GRAY)
        self.cell(95, 6.5, label)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*DARK_GRAY)
        self.cell(85, 6.5, value, align="R")
        self.ln(7)

    # ================================================================
    #  CAPA
    # ================================================================
    def add_cover(self):
        d = self.dados
        c = self.c
        self.ln(4)

        self.set_font("Helvetica", "B", 22)
        self.set_text_color(*GREEN)
        self.cell(0, 12, "Simulação de Consórcio", new_x="LMARGIN", new_y="NEXT")

        self.set_font("Helvetica", "", 10)
        self.set_text_color(*TEXT_GRAY)
        self.cell(0, 6, "Proposta simulada  -  Somus Capital", new_x="LMARGIN", new_y="NEXT")
        self.ln(5)

        # Card do cliente
        y = self.get_y()
        self.set_fill_color(*LIGHT_GRAY)
        self.rect(10, y, 190, 22, "F")
        self.set_fill_color(*GREEN)
        self.rect(10, y, 2.5, 22, "F")
        self.set_xy(17, y + 3)
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(*GREEN)
        self.cell(0, 7, sanitize_text(d.get("cliente_nome", "")))
        self.set_xy(17, y + 11)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*DARK_GRAY)
        self.cell(0, 6, f"Assessor: {sanitize_text(d.get('assessor', ''))}        Data: {datetime.now().strftime('%d/%m/%Y')}")
        self.set_y(y + 27)

        # Parâmetros
        self._section_title("Parâmetros da Simulação")
        self.ln(1)

        pr_pct = d.get("parcela_reduzida_pct", 100)
        pr_label = "Integral (100%)" if pr_pct == 100 else f"Reduzida ({pr_pct}% do fundo comum)"

        params = [
            ("Tipo do Bem", sanitize_text(d.get("tipo_bem", ""))),
            ("Administradora", sanitize_text(d.get("administradora", "-"))),
            ("Valor da Carta de Crédito", fmt_currency(d.get("valor_carta", 0))),
            ("Prazo do Grupo", f"{d.get('prazo_meses', 0)} meses"),
            ("Taxa de Administração (total)", fmt_pct(d.get("taxa_adm", 0))),
            ("Fundo de Reserva (total)", fmt_pct(d.get("fundo_reserva", 0))),
            ("Seguro Prestamista (total)", fmt_pct(d.get("seguro", 0))),
            ("Prazo de Contemplação", f"Mês {d.get('prazo_contemplação', '-')}"),
            ("Parcela até Contemplação", pr_label),
            ("Lance Livre Ofertado", f"{fmt_pct(d.get('lance_livre_pct', 0))}  ({fmt_currency(c['ll_val'])})"),
            ("Lance Embutido Ofertado", f"{fmt_pct(d.get('lance_embutido_pct', 0))}  ({fmt_currency(c['le_val'])})"),
        ]
        if c["tem_correção"]:
            corr_lbl = f"{c['tipo_corr']} {fmt_pct(c['corr_anual'] * 100)} a.a."
            if c["tipo_corr"] == "Pós-fixado":
                corr_lbl += f" ({c['índice_corr']})"
            params.append(("Correção Anual", corr_lbl))
        for i, (label, value) in enumerate(params):
            self._param_row(label, value, i)

        self.ln(4)

        # ======== RESULTADO FASE 1 ========
        self._section_title("Fase 1  -  Antes da Contemplação")
        self.ln(1)

        y = self.get_y()
        self.set_fill_color(*GREEN)
        self.rect(10, y, 190, 20, "F")
        self.set_xy(14, y + 2)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*WHITE)
        red_note = f"  (parcela reduzida {pr_pct}%)" if pr_pct < 100 else ""
        self.cell(0, 5, f"PARCELA FASE 1{red_note}  -  {c['pc']} meses")
        self.set_xy(14, y + 8)
        self.set_font("Helvetica", "B", 16)
        self.cell(0, 10, fmt_currency(c["p1"]))
        if c["tem_correção"]:
            self.set_xy(120, y + 2)
            self.set_font("Helvetica", "", 7)
            self.cell(0, 5, f"Parcela final estimada (corrigida):")
            self.set_xy(120, y + 8)
            self.set_font("Helvetica", "B", 11)
            self.cell(0, 10, fmt_currency(c["p1_final"]))
        self.set_y(y + 23)

        fc_lbl = "Fundo Comum" + (f" ({pr_pct}%)" if pr_pct < 100 else "")
        self._color_bar_row(fc_lbl, fmt_currency(c["fc_red"]), BLUE)
        self._color_bar_row("Taxa de Administração", fmt_currency(c["ta"]), ORANGE)
        self._color_bar_row("Fundo de Reserva", fmt_currency(c["fr"]), TEAL)
        self._color_bar_row("Seguro Prestamista", fmt_currency(c["sg"]), PURPLE)

        # ======== CONTEMPLAÇÃO / LANCES ========
        if c["ll_val"] > 0 or c["le_val"] > 0:
            self.ln(2)
            self._section_title(f"Contemplação  -  Mês {c['pc']}")
            self.ln(1)

            y = self.get_y()
            self.set_fill_color(*BLUE)
            self.rect(10, y, 190, 8, "F")
            self.set_xy(14, y + 1)
            self.set_font("Helvetica", "B", 8)
            self.set_text_color(*WHITE)
            self.cell(0, 6, "LANCES OFERTADOS")
            self.set_y(y + 9)

            i = 0
            if c["ll_val"] > 0:
                self._param_row("Lance Livre (recursos próprios)", fmt_currency(c["ll_val"]), i)
                i += 1
            if c["le_val"] > 0:
                self._param_row("Lance Embutido (da carta)", fmt_currency(c["le_val"]), i)
                i += 1
            self._param_row("Carta de Crédito Líquida", fmt_currency(c["carta_liq"]), i)

        # ======== RESULTADO FASE 2 ========
        if c["mr"] > 0:
            self.ln(2)
            self._section_title("Fase 2  -  Após Contemplação")
            self.ln(1)

            y = self.get_y()
            self.set_fill_color(*GREEN_LIGHT)
            self.rect(10, y, 190, 20, "F")
            self.set_xy(14, y + 2)
            self.set_font("Helvetica", "", 8)
            self.set_text_color(*WHITE)
            self.cell(0, 5, f"PARCELA FASE 2 (reajustada)  -  {c['mr']} meses")
            self.set_xy(14, y + 8)
            self.set_font("Helvetica", "B", 16)
            self.cell(0, 10, fmt_currency(c["p2"]))
            if c["tem_correção"]:
                self.set_xy(120, y + 2)
                self.set_font("Helvetica", "", 7)
                self.cell(0, 5, f"Parcela final estimada (corrigida):")
                self.set_xy(120, y + 8)
                self.set_font("Helvetica", "B", 11)
                self.cell(0, 10, fmt_currency(c["p2_final"]))
            self.set_y(y + 23)

            self._color_bar_row("Fundo Comum (reajustado)", fmt_currency(c["fc_f2"]), BLUE)
            self._color_bar_row("Taxa de Administração", fmt_currency(c["ta"]), ORANGE)
            self._color_bar_row("Fundo de Reserva", fmt_currency(c["fr"]), TEAL)
            self._color_bar_row("Seguro Prestamista", fmt_currency(c["sg"]), PURPLE)

        # ======== RESUMO FINANCEIRO ========
        self.ln(3)
        self._section_title("Resumo Financeiro")
        self.ln(1)

        i = 0
        self._param_row("Valor da Carta de Crédito", fmt_currency(c["vc"]), i); i += 1
        if c["le_val"] > 0:
            self._param_row("Carta Líquida (apos lance embutido)", fmt_currency(c["carta_liq"]), i); i += 1
        self._param_row(f"Desembolso Fase 1 ({c['pc']} meses)", fmt_currency(c["desemb1"]), i); i += 1
        if c["ll_val"] > 0:
            self._param_row("Lance Livre (desembolso próprio)", fmt_currency(c["ll_val"]), i); i += 1
        if c["mr"] > 0:
            self._param_row(f"Desembolso Fase 2 ({c['mr']} meses)", fmt_currency(c["desemb2"]), i); i += 1

        # Total destaque (sem correção)
        self.set_fill_color(*GREEN)
        self.set_x(10)
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*WHITE)
        self.cell(110, 8, "  TOTAL DESEMBOLSADO (SEM CORREÇÃO)", fill=True)
        self.set_font("Helvetica", "B", 10)
        self.cell(80, 8, f"{fmt_currency(c['total_des'])}  ", fill=True, align="R")
        self.ln(8)

        # Total com correção
        if c["tem_correção"]:
            corr_lbl = f"{c['tipo_corr']} {fmt_pct(c['corr_anual'] * 100)} a.a."
            if c["tipo_corr"] == "Pós-fixado":
                corr_lbl += f" ({c['índice_corr']})"
            self.set_fill_color(*RED)
            self.set_x(10)
            self.set_font("Helvetica", "B", 9)
            self.set_text_color(*WHITE)
            self.cell(110, 8, f"  TOTAL CORRIGIDO ({corr_lbl})", fill=True)
            self.set_font("Helvetica", "B", 10)
            self.cell(80, 8, f"{fmt_currency(c['total_des_corr'])}  ", fill=True, align="R")
            self.ln(8)

        self._param_row("Total em Taxas e Encargos", fmt_currency(c["total_tax"]), 0)
        custo_pct = (c["total_tax"] / c["vc"] * 100) if c["vc"] > 0 else 0
        self._param_row("Custo Efetivo sobre a Carta", fmt_pct(custo_pct), 1)

    # ================================================================
    #  RESUMO DA OPERAÇÃO
    # ================================================================
    def add_resumo_operacao(self):
        c = self.c
        d = self.dados

        self.add_page()
        self.ln(2)
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*GREEN)
        self.cell(0, 10, "Resumo da Operação", new_x="LMARGIN", new_y="NEXT")
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*TEXT_GRAY)
        self.cell(0, 5, "Visão consolidada dos custos e indicadores da operação de consórcio.",
                  new_x="LMARGIN", new_y="NEXT")
        self.ln(4)

        total_ref = c["total_des_corr"] if c["tem_correção"] else c["total_des"]

        # === Custo Efetivo destaque ===
        y = self.get_y()
        self.set_fill_color(*GREEN)
        self.rect(10, y, 190, 38, "F")
        self.set_fill_color(*BLUE)
        self.rect(10, y, 3, 38, "F")

        self.set_xy(18, y + 3)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(200, 230, 210)
        self.cell(0, 5, "CUSTO EFETIVO DO CONSÓRCIO")

        self.set_xy(18, y + 10)
        self.set_font("Helvetica", "B", 22)
        self.set_text_color(*WHITE)
        self.cell(0, 10, f"{fmt_pct(c['ce_pct'])} sobre o crédito")

        self.set_xy(18, y + 22)
        self.set_font("Helvetica", "B", 13)
        self.cell(0, 6, f"{fmt_pct(c['ce_anual'] * 100)} a.a.   |   {fmt_pct(c['ce_mensal'] * 100)} a.m.")

        self.set_xy(18, y + 30)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(200, 230, 210)
        self.cell(0, 5, "Taxa composta que representa o custo total dos encargos do consórcio ao longo do prazo.")

        self.set_y(y + 42)

        # === Indicadores ===
        self._section_title("Indicadores da Operação")
        self.ln(1)

        rows = [
            ("Crédito Recebido (carta líquida)", fmt_currency(c["carta_liq"])),
            ("Total Desembolsado pelo Cliente", fmt_currency(total_ref)),
            ("Custo Total (encargos)", fmt_currency(c["custo_total"])),
            ("Custo sobre o Crédito", fmt_pct(c["ce_pct"])),
            ("Taxa de Custo Anual (composta)", fmt_pct(c["ce_anual"] * 100)),
            ("Taxa de Custo Mensal (composta)", fmt_pct(c["ce_mensal"] * 100)),
            ("Relação Total Pago / Crédito", f"{c['relação_custo']:.2f}x".replace(".", ",")),
            ("Parcela Media", fmt_currency(c["parcela_media"])),
            ("Prazo Total", f"{c['pz']} meses ({c['pz'] / 12:.1f} anos)".replace(".", ",")),
            ("Contemplação no Mês", f"{c['pc']}"),
        ]

        for i, (label, value) in enumerate(rows):
            self._param_row(label, value, i)

        self.ln(3)

        # === Composição do custo ===
        self._section_title("Composição do Custo")
        self.ln(1)

        taxa_adm_total = c["ta"] * c["pz"]
        fundo_res_total = c["fr"] * c["pz"]
        seguro_total = c["sg"] * c["pz"]
        correção_custo = (total_ref - c["total_des"]) if c["tem_correção"] else 0

        self._color_bar_row("Taxa de Administração", fmt_currency(taxa_adm_total), ORANGE)
        self._color_bar_row("Fundo de Reserva", fmt_currency(fundo_res_total), TEAL)
        self._color_bar_row("Seguro Prestamista", fmt_currency(seguro_total), PURPLE)
        if c["tem_correção"]:
            self._color_bar_row("Custo da Correção Monetária", fmt_currency(correção_custo), RED)
        if c["ll_val"] > 0:
            self._color_bar_row("Lance Livre (recurso próprio)", fmt_currency(c["ll_val"]), BLUE)

        self.ln(3)

        # === Quadro síntese ===
        self._section_title("Síntese")
        self.ln(1)

        y = self.get_y()
        w3 = 60

        # Box 1 - Você recebe
        self.set_fill_color(*LIGHT_GRAY)
        self.rect(10, y, w3, 28, "F")
        self.set_fill_color(*GREEN)
        self.rect(10, y, w3, 8, "F")
        self.set_xy(12, y + 1)
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(*WHITE)
        self.cell(0, 6, "VOCÊ RECEBE")
        self.set_xy(14, y + 11)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*GREEN)
        self.cell(0, 5, fmt_currency(c["carta_liq"]))
        self.set_xy(14, y + 18)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*TEXT_GRAY)
        self.cell(0, 4, "Carta de crédito líquida")

        # Box 2 - Você paga
        self.set_fill_color(*LIGHT_GRAY)
        self.rect(75, y, w3, 28, "F")
        self.set_fill_color(*BLUE)
        self.rect(75, y, w3, 8, "F")
        self.set_xy(77, y + 1)
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(*WHITE)
        self.cell(0, 6, "VOCÊ PAGA")
        self.set_xy(79, y + 11)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*BLUE)
        self.cell(0, 5, fmt_currency(total_ref))
        self.set_xy(79, y + 18)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*TEXT_GRAY)
        self.cell(0, 4, f"Total em {c['pz']} meses")

        # Box 3 - Custo
        self.set_fill_color(*LIGHT_GRAY)
        self.rect(140, y, w3, 28, "F")
        self.set_fill_color(*ORANGE)
        self.rect(140, y, w3, 8, "F")
        self.set_xy(142, y + 1)
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(*WHITE)
        self.cell(0, 6, "CUSTO DA OPERAÇÃO")
        self.set_xy(144, y + 11)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*ORANGE)
        self.cell(0, 5, fmt_currency(c["custo_total"]))
        self.set_xy(144, y + 18)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*TEXT_GRAY)
        self.cell(0, 4, f"{fmt_pct(c['ce_pct'])} sobre o crédito")

        self.set_y(y + 33)

    # ================================================================
    #  CRONOGRAMA (duas fases)
    # ================================================================
    def add_payment_schedule(self):
        self.add_page()
        c = self.c

        self._section_title("Cronograma de Pagamentos")
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(*TEXT_GRAY)
        self.cell(
            0, 5,
            "Projeção mensal. Valores estimados, sujeitos a reajuste pela administradora.",
            new_x="LMARGIN", new_y="NEXT",
        )
        self.ln(2)

        has_corr = c["tem_correção"]
        # Larguras das colunas (total 190mm = margem 10 a 200)
        if has_corr:
            W = {"mês": 11, "fase": 9, "fc": 25, "ta": 25, "fr": 22, "sg": 20,
                 "parc": 27, "fator": 16, "acum": 35}
        else:
            W = {"mês": 12, "fase": 10, "fc": 28, "ta": 28, "fr": 24, "sg": 22,
                 "parc": 30, "fator": 0, "acum": 36}

        self._draw_schedule_header(has_corr, W)
        acumulado = 0.0
        acumulado_corr = 0.0

        for mes in range(1, c["pz"] + 1):
            if self.get_y() > 265:
                self.add_page()
                self._section_title("Cronograma de Pagamentos (continuação)")
                self.ln(2)
                self._draw_schedule_header(has_corr, W)

            # Determinar fase
            is_fase1 = mes <= c["pc"]
            fator = (1 + c["corr_anual"]) ** ((mes - 1) // 12) if has_corr else 1.0

            if is_fase1:
                fc_mes = c["fc_red"]
                parcela_mes = c["p1"]
                fase_label = "F1"
            else:
                fc_mes = c["fc_f2"]
                parcela_mes = c["p2"]
                fase_label = "F2"

            parcela_corr = parcela_mes * fator

            # Linha de lance (mes da contemplação)
            if mes == c["pc"] + 1 and c["lance_total"] > 0:
                self.set_fill_color(*BLUE)
                self.set_text_color(*WHITE)
                self.set_font("Helvetica", "B", 6.5)
                self.set_x(10)

                lance_txt = ""
                if c["ll_val"] > 0:
                    lance_txt += f"Lance Livre: {fmt_currency(c['ll_val'])}  "
                if c["le_val"] > 0:
                    lance_txt += f"Lance Embutido: {fmt_currency(c['le_val'])}  "
                lance_txt += f"Carta Líquida: {fmt_currency(c['carta_liq'])}"

                self.cell(190, 6, f"  CONTEMPLAÇÃO MÊS {c['pc']}  -  {lance_txt}", fill=True)
                self.ln(6)

                if self.get_y() > 265:
                    self.add_page()
                    self._section_title("Cronograma de Pagamentos (continuação)")
                    self.ln(2)
                    self._draw_schedule_header(has_corr, W)

            acumulado += parcela_mes
            acumulado_corr += parcela_corr

            # Cor alternada com tom diferente por fase
            if is_fase1:
                bg = LIGHT_GRAY if mes % 2 == 0 else WHITE
            else:
                bg = (240, 245, 255) if mes % 2 == 0 else WHITE
            self.set_fill_color(*bg)

            self.set_font("Helvetica", "", 6.5)
            self.set_text_color(*DARK_GRAY)
            self.set_x(10)
            self.cell(W["mês"], 5.5, f"{mes}", fill=True, align="C")
            self.cell(W["fase"], 5.5, fase_label, fill=True, align="C")
            self.cell(W["fc"], 5.5, f"{fmt_currency(fc_mes)} ", fill=True, align="R")
            self.cell(W["ta"], 5.5, f"{fmt_currency(c['ta'])} ", fill=True, align="R")
            self.cell(W["fr"], 5.5, f"{fmt_currency(c['fr'])} ", fill=True, align="R")
            self.cell(W["sg"], 5.5, f"{fmt_currency(c['sg'])} ", fill=True, align="R")
            self.cell(W["parc"], 5.5, f"{fmt_currency(parcela_mes)} ", fill=True, align="R")
            if has_corr:
                self.cell(W["fator"], 5.5, f"x{fator:.3f}", fill=True, align="C")
            self.cell(W["acum"], 5.5, f"{fmt_currency(acumulado_corr if has_corr else acumulado)} ", fill=True, align="R")
            self.ln(5.5)

        # Total
        self.set_fill_color(*GREEN)
        self.set_text_color(*WHITE)
        self.set_font("Helvetica", "B", 6.5)
        self.set_x(10)
        total_fc = c["fc_red"] * c["pc"] + c["fc_f2"] * c["mr"]
        total_parcelas = c["desemb1"] + c["desemb2"]
        self.cell(W["mês"], 6, "TOTAL", fill=True, align="C")
        self.cell(W["fase"], 6, "", fill=True)
        self.cell(W["fc"], 6, f"{fmt_currency(total_fc)} ", fill=True, align="R")
        self.cell(W["ta"], 6, f"{fmt_currency(c['ta'] * c['pz'])} ", fill=True, align="R")
        self.cell(W["fr"], 6, f"{fmt_currency(c['fr'] * c['pz'])} ", fill=True, align="R")
        self.cell(W["sg"], 6, f"{fmt_currency(c['sg'] * c['pz'])} ", fill=True, align="R")
        self.cell(W["parc"], 6, f"{fmt_currency(total_parcelas)} ", fill=True, align="R")
        if has_corr:
            self.cell(W["fator"], 6, "", fill=True)
        self.cell(W["acum"], 6, f"{fmt_currency(acumulado_corr if has_corr else acumulado)} ", fill=True, align="R")
        self.ln(6)

    def _draw_schedule_header(self, has_corr=False, W=None):
        if W is None:
            W = {"mês": 12, "fase": 10, "fc": 28, "ta": 28, "fr": 24, "sg": 22,
                 "parc": 30, "fator": 0, "acum": 36}
        self.set_fill_color(*BLUE)
        self.set_font("Helvetica", "B", 6)
        self.set_text_color(*WHITE)
        self.set_x(10)
        self.cell(W["mês"], 6, "MES", fill=True, align="C")
        self.cell(W["fase"], 6, "FASE", fill=True, align="C")
        self.cell(W["fc"], 6, "FDO. COMUM ", fill=True, align="R")
        self.cell(W["ta"], 6, "TAXA ADM. ", fill=True, align="R")
        self.cell(W["fr"], 6, "FDO. RES. ", fill=True, align="R")
        self.cell(W["sg"], 6, "SEGURO ", fill=True, align="R")
        self.cell(W["parc"], 6, "PARCELA ", fill=True, align="R")
        if has_corr:
            self.cell(W["fator"], 6, "FATOR ", fill=True, align="C")
        self.cell(W["acum"], 6, "ACUMULADO ", fill=True, align="R")
        self.ln(6)

    # ================================================================
    #  DISCLAIMER
    # ================================================================
    def add_disclaimer(self):
        self.ln(6)
        self.set_draw_color(*MID_GRAY)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(3)
        self.set_font("Helvetica", "I", 6)
        self.set_text_color(150, 150, 150)
        self.multi_cell(
            190, 3.5,
            "Esta simulação e meramente ilustrativa e não constitui oferta ou proposta formal. "
            "Os valores apresentados são estimativas baseadas nos parâmetros informados e podem "
            "sofrer alterações conforme regras da administradora do consórcio. "
            "As parcelas podem ser reajustadas anualmente conforme índice previsto em contrato. "
            "O lance embutido reduz o valor líquido da carta de crédito recebida. "
            "A taxa de administração incide sobre o valor original da carta, mesmo com lance embutido. "
            "Para informações oficiais, consulte a administradora ou seu assessor. "
            f"Simulação gerada em {datetime.now().strftime('%d/%m/%Y as %H:%M')}.",
            align="C",
        )


def generate_consorcio_pdf(dados, output_dir):
    """Gera PDF de simulação de consórcio com duas fases.

    dados dict keys:
        cliente_nome, assessor, tipo_bem, administradora,
        valor_carta, prazo_meses, taxa_adm, fundo_reserva, seguro,
        prazo_contemplação, parcela_reduzida_pct,
        lance_livre_pct, lance_embutido_pct,
        lance_livre_valor, lance_embutido_valor, carta_líquida,
        parcela_fase1, parcela_fase2,
        fc_integral, fc_reduzido, fc_fase2,
        ta_mensal, fr_mensal, sg_mensal,
        desemb_fase1, desemb_fase2,
        total_desembolsado, total_taxas
    """
    os.makedirs(output_dir, exist_ok=True)

    pdf = ConsorcioPDF(dados)
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.add_cover()
    pdf.add_resumo_operacao()
    pdf.add_payment_schedule()
    pdf.add_disclaimer()

    nome = sanitize_text(dados.get("cliente_nome", "Cliente"))
    nome_limpo = nome.replace(" ", "_").replace("/", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Consórcio_{nome_limpo}_{timestamp}.pdf"
    filepath = os.path.join(output_dir, filename)
    pdf.output(filepath)
    return filepath


# ==================================================================
#  RELATÓRIO EXPLICATIVO
# ==================================================================

class RelatorioConsorcioPDF(FPDF):
    """PDF com explicação detalhada (matematica + linguagem comum) da simulação."""

    def __init__(self, dados):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.dados = dados
        self.c = _calc(dados)
        self.set_auto_page_break(auto=True, margin=20)
        self.logo_path = LOGO_PATH if os.path.exists(LOGO_PATH) else None

    def header(self):
        self.set_fill_color(*GREEN)
        self.rect(0, 0, 210, 18, "F")
        if self.logo_path:
            try:
                self.image(self.logo_path, x=8, y=3, h=12)
            except Exception:
                pass
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*WHITE)
        self.set_xy(0, 4)
        self.cell(210, 10, "RELATÓRIO EXPLICATIVO  -  SIMULAÇÃO DE CONSÓRCIO", align="C")
        self.set_fill_color(*BLUE)
        self.rect(0, 18, 210, 1.2, "F")
        self.set_y(23)

    def footer(self):
        self.set_y(-14)
        self.set_draw_color(*MID_GRAY)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(1.5)
        self.set_font("Helvetica", "", 6)
        self.set_text_color(140, 140, 140)
        self.cell(0, 5, "Somus Capital  |  Relatório Explicativo", align="L")
        self.cell(0, 5, f"Página {self.page_no()}/{{nb}}", align="R")

    def _title(self, text):
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(*GREEN)
        self.cell(0, 9, sanitize_text(text), new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def _subtitle(self, text):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*BLUE)
        self.cell(0, 7, sanitize_text(text), new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def _text(self, text):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*DARK_GRAY)
        self.multi_cell(190, 5, sanitize_text(text))
        self.ln(1)

    def _formula(self, text):
        y = self.get_y()
        self.set_fill_color(240, 242, 248)
        self.set_x(10)
        self.set_font("Helvetica", "B", 8.5)
        self.set_text_color(60, 60, 120)
        h = 7
        self.cell(190, h, f"    {sanitize_text(text)}", fill=True)
        self.ln(h + 2)

    def _result_box(self, label, value):
        y = self.get_y()
        self.set_fill_color(*LIGHT_GRAY)
        self.rect(10, y, 190, 7, "F")
        self.set_fill_color(*GREEN)
        self.rect(10, y, 2, 7, "F")
        self.set_xy(16, y)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*TEXT_GRAY)
        self.cell(100, 7, sanitize_text(label))
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*DARK_GRAY)
        self.cell(84, 7, sanitize_text(value), align="R")
        self.ln(9)

    def build(self):
        d = self.dados
        c = self.c

        vc = c["vc"]
        pz = c["pz"]
        pc = c["pc"]
        mr = c["mr"]
        pr = c["pr"]
        ta_pct = d.get("taxa_adm", 0)
        fr_pct = d.get("fundo_reserva", 0)
        sg_pct = d.get("seguro", 0)
        ll_pct = d.get("lance_livre_pct", 0)
        le_pct = d.get("lance_embutido_pct", 0)
        corr = d.get("correção_anual", 0)
        pr_pct = d.get("parcela_reduzida_pct", 100)

        # ============ PAGINA 1: INTRODUÇÃO ============
        self.ln(3)
        self.set_font("Helvetica", "B", 20)
        self.set_text_color(*GREEN)
        self.cell(0, 12, "Relatório Explicativo", new_x="LMARGIN", new_y="NEXT")
        self.set_font("Helvetica", "", 10)
        self.set_text_color(*TEXT_GRAY)
        self.cell(0, 6, f"Cliente: {sanitize_text(d.get('cliente_nome', ''))}  |  "
                        f"Data: {datetime.now().strftime('%d/%m/%Y')}", new_x="LMARGIN", new_y="NEXT")
        self.ln(4)

        self._text(
            "Este relatório detalha, passo a passo, como cada valor da simulação de consórcio "
            "foi calculado. O objetivo e dar total transparência ao cliente, apresentando tanto "
            "as fórmulas matemáticas utilizadas quanto a explicação em linguagem acessível."
        )

        self.ln(2)
        self._title("1. O que e um Consórcio?")
        self._text(
            "O consórcio e uma modalidade de compra planejada em grupo. Um conjunto de pessoas "
            "contribui mensalmente para um fundo comum, e a cada mes um ou mais participantes "
            "são contemplados (por sorteio ou lance) e recebem uma carta de crédito para adquirir "
            "o bem desejado. Diferente de um financiamento, não ha cobranca de juros - porem ha "
            "taxas administrativas e encargos que compõem o custo total."
        )

        # ============ CÁLCULO DA PARCELA BASE ============
        self._title("2. Cálculo da Parcela Mensal Base")
        self._text(
            "A parcela mensal do consórcio e composta por quatro elementos: o Fundo Comum "
            "(que forma o crédito), a Taxa de Administração, o Fundo de Reserva e o Seguro. "
            "Cada componente e calculado dividindo o valor total pelo número de meses do grupo."
        )

        self._subtitle("2.1 Fundo Comum (a parte que vira crédito)")
        self._text(
            f"O fundo comum é a parcela que efetivamente forma a sua carta de crédito. "
            f"O valor da carta ({fmt_currency(vc)}) é dividido igualmente pelo prazo do grupo ({pz} meses):"
        )
        self._formula(f"Fundo Comum = Carta / Prazo = {fmt_currency(vc)} / {pz} = {fmt_currency(c['fc_integral'])} por mês")

        self._subtitle("2.2 Taxa de Administração")
        self._text(
            f"A taxa de administração e a remuneração da administradora por gerenciar o grupo. "
            f"É cobrada como percentual total sobre o valor da carta ({fmt_pct(ta_pct)}), diluída ao longo dos {pz} meses:"
        )
        self._formula(f"Taxa Adm. mensal = (Carta x Taxa%) / Prazo = ({fmt_currency(vc)} x {fmt_pct(ta_pct)}) / {pz} = {fmt_currency(c['ta'])}")

        self._subtitle("2.3 Fundo de Reserva")
        self._text(
            f"O fundo de reserva protege o grupo contra inadimplência e imprevistos. "
            f"Funciona como o taxa de administração: {fmt_pct(fr_pct)} sobre a carta, diluído em {pz} meses:"
        )
        self._formula(f"Fdo. Reserva mensal = ({fmt_currency(vc)} x {fmt_pct(fr_pct)}) / {pz} = {fmt_currency(c['fr'])}")

        if sg_pct > 0:
            self._subtitle("2.4 Seguro Prestamista")
            self._text(
                f"O seguro garante a quitação das parcelas em caso de sinistro. "
                f"Taxa de {fmt_pct(sg_pct)} sobre a carta:"
            )
            self._formula(f"Seguro mensal = ({fmt_currency(vc)} x {fmt_pct(sg_pct)}) / {pz} = {fmt_currency(c['sg'])}")

        self._subtitle("2.5 Parcela Integral")
        self._text("Somando todos os componentes, temos a parcela mensal completa:")
        componentes = f"{fmt_currency(c['fc_integral'])} + {fmt_currency(c['ta'])} + {fmt_currency(c['fr'])}"
        if sg_pct > 0:
            componentes += f" + {fmt_currency(c['sg'])}"
        parcela_integral = c["fc_integral"] + c["taxas"]
        self._formula(f"Parcela = Fdo.Comum + Taxa Adm. + Fdo.Reserva + Seguro = {componentes} = {fmt_currency(parcela_integral)}")
        self._result_box("Parcela Integral Mensal", fmt_currency(parcela_integral))

        # ============ MODELO DUAS FASES ============
        self._title("3. Modelo de Duas Fases")
        self._text(
            f"Esta simulação divide o consórcio em duas fases, usando o mes {pc} como ponto "
            f"de contemplação (momento em que você recebe a carta de crédito)."
        )

        self._subtitle("3.1 Fase 1 - Antes da Contemplação (meses 1 a {})".format(pc))

        if pr_pct < 100:
            self._text(
                f"Na fase 1, foi escolhida a opção de parcela reduzida a {pr_pct}% do fundo comum. "
                f"Isso significa que você paga apenas {pr_pct}% da parcela de fundo comum antes da "
                f"contemplação, reduzindo o desembolso inicial. A diferenca será compensada na fase 2."
            )
            self._formula(f"Fdo.Comum reduzido = {fmt_currency(c['fc_integral'])} x {pr_pct}% = {fmt_currency(c['fc_red'])}")
        else:
            self._text(
                "Na fase 1, a parcela é paga integralmente (100% do fundo comum). "
                "Você contribui o valor cheio desde o inicio."
            )

        self._formula(f"Parcela Fase 1 = {fmt_currency(c['fc_red'])} + {fmt_currency(c['taxas'])} (taxas) = {fmt_currency(c['p1'])} por mês")
        self._formula(f"Total Fase 1 = {fmt_currency(c['p1'])} x {pc} meses = {fmt_currency(c['desemb1'])}")
        self._result_box(f"Parcela Fase 1 ({pc} meses)", fmt_currency(c["p1"]))

        fundo_f1 = c["fc_red"] * pc
        self._text(
            f"Ao final da fase 1, você terá contribuido {fmt_currency(fundo_f1)} para o fundo comum "
            f"(de um total de {fmt_currency(vc)} necessários)."
        )

        # ============ LANCES ============
        if c["ll_val"] > 0 or c["le_val"] > 0:
            self._title("4. Lances na Contemplação")
            self._text(
                f"No mes {pc}, ao ser contemplado, você ofertou lances para antecipar a "
                f"contemplação ou reduzir o saldo devedor. Os lances abatem o saldo de fundo "
                f"comum restante."
            )

            if c["ll_val"] > 0:
                self._subtitle("4.1 Lance Livre")
                self._text(
                    f"O lance livre é um valor em dinheiro próprio que você oferece para aumentar "
                    f"suas chances de contemplação. Este valor abate diretamente o saldo devedor "
                    f"de fundo comum, mas não reduz a carta de crédito."
                )
                self._formula(f"Lance Livre = Carta x {fmt_pct(ll_pct)} = {fmt_currency(vc)} x {fmt_pct(ll_pct)} = {fmt_currency(c['ll_val'])}")

            if c["le_val"] > 0:
                self._subtitle("4.2 Lance Embutido")
                self._text(
                    f"O lance embutido é descontado da própria carta de crédito. Ou seja, você "
                    f"recebe uma carta menor, mas abate o saldo devedor. A taxa de administração "
                    f"continua incidindo sobre o valor ORIGINAL da carta."
                )
                self._formula(f"Lance Embutido = Carta x {fmt_pct(le_pct)} = {fmt_currency(vc)} x {fmt_pct(le_pct)} = {fmt_currency(c['le_val'])}")
                self._formula(f"Carta Líquida = {fmt_currency(vc)} - {fmt_currency(c['le_val'])} = {fmt_currency(c['carta_liq'])}")
                self._result_box("Carta de Crédito Líquida Recebida", fmt_currency(c["carta_liq"]))

            lance_total_str = fmt_currency(c["lance_total"])
            self._text(
                f"O total de lances ({lance_total_str}) é abatido do saldo de fundo comum "
                f"que ainda faltava pagar, reduzindo as parcelas da fase 2."
            )

        # ============ FASE 2 ============
        lance_section = "4" if (c["ll_val"] > 0 or c["le_val"] > 0) else "3"
        fase2_section = str(int(lance_section) + 1)

        if mr > 0:
            self._title(f"{fase2_section}. Fase 2 - Após Contemplação (meses {pc + 1} a {pz})")
            self._text(
                f"Apos a contemplação, o fundo comum restante é recalculado. Do total de "
                f"{fmt_currency(vc)}, subtraimos o que ja foi pago na fase 1 e os lances:"
            )
            self._formula(
                f"Fdo. Restante = {fmt_currency(vc)} - {fmt_currency(fundo_f1)} (pago F1) "
                f"- {fmt_currency(c['lance_total'])} (lances) = {fmt_currency(c['fc_f2'] * mr)}"
            )
            self._text(
                f"Este saldo restante é dividido pelos {mr} meses que faltam:"
            )
            self._formula(f"Fdo.Comum F2 = {fmt_currency(c['fc_f2'] * mr)} / {mr} = {fmt_currency(c['fc_f2'])} por mês")
            self._formula(f"Parcela Fase 2 = {fmt_currency(c['fc_f2'])} + {fmt_currency(c['taxas'])} (taxas) = {fmt_currency(c['p2'])} por mês")
            self._formula(f"Total Fase 2 = {fmt_currency(c['p2'])} x {mr} meses = {fmt_currency(c['desemb2'])}")
            self._result_box(f"Parcela Fase 2 ({mr} meses)", fmt_currency(c["p2"]))

            if pr_pct < 100:
                self._text(
                    f"Note que a parcela da fase 2 ({fmt_currency(c['p2'])}) é maior que a da "
                    f"fase 1 ({fmt_currency(c['p1'])}). Isso ocorre porque na fase 1 foi pago "
                    f"apenas {pr_pct}% do fundo comum, é o restante precisa ser compensado agora."
                )

        # ============ CORREÇÃO ANUAL ============
        corr_section = str(int(fase2_section) + 1)
        if corr > 0:
            self._title(f"{corr_section}. Correção Anual das Parcelas")
            tipo = d.get("tipo_correção", "Pós-fixado")
            índice = d.get("índice_correção", "INCC")

            if tipo == "Pós-fixado":
                self._text(
                    f"As parcelas são corrigidas anualmente pelo índice {índice}, com taxa estimada "
                    f"de {fmt_pct(corr)} ao ano (pós-fixado). Na prática, o índice pode variar; "
                    f"aqui usamos uma projeção fixa para fins de simulação."
                )
            else:
                self._text(
                    f"As parcelas são corrigidas anualmente por uma taxa pre-fixada de "
                    f"{fmt_pct(corr)} ao ano, definida em contrato."
                )

            self._text(
                "A correção e aplicada a cada 12 meses de forma composta. "
                "A fórmula do fator de correção para cada mes e:"
            )
            self._formula(f"Fator(mes) = (1 + {fmt_pct(corr)}/100) ^ (parte inteira de (mes - 1) / 12)")
            self._text("Por exemplo:")
            self._formula(f"Mês 1: fator = 1,000  |  Mês 13: fator = {(1 + corr/100)**1:.4f}  |  "
                          f"Mês 25: fator = {(1 + corr/100)**2:.4f}  |  Mês {pz}: fator = {(1 + corr/100)**((pz-1)//12):.4f}")

            self._text(
                f"A parcela corrigida de cada mes e: Parcela Base x Fator. "
                f"Assim, a primeira parcela e {fmt_currency(c['p1'])}, mas a última parcela da "
                f"fase 1 (mes {pc}) será {fmt_currency(c['p1_final'])}. "
                f"A última parcela da fase 2 (mes {pz}) será {fmt_currency(c['p2_final'])}."
            )
            self._result_box(f"Parcela Fase 1 final (mes {pc})", fmt_currency(c["p1_final"]))
            self._result_box(f"Parcela Fase 2 final (mes {pz})", fmt_currency(c["p2_final"]))

        # ============ TOTAIS ============
        totais_section = str(int(corr_section if corr > 0 else fase2_section) + 1)
        self._title(f"{totais_section}. Totalização")

        self._text("Somando todos os desembolsos do cliente ao longo do consórcio:")
        if corr > 0:
            self._subtitle("Sem correção (valores nominais)")
        self._formula(f"Total Fase 1 = {fmt_currency(c['p1'])} x {pc} = {fmt_currency(c['desemb1'])}")
        if c["ll_val"] > 0:
            self._formula(f"Lance Livre = {fmt_currency(c['ll_val'])}")
        if mr > 0:
            self._formula(f"Total Fase 2 = {fmt_currency(c['p2'])} x {mr} = {fmt_currency(c['desemb2'])}")
        self._formula(f"TOTAL DESEMBOLSADO = {fmt_currency(c['total_des'])}")

        if corr > 0:
            self._subtitle("Com correção (valores corrigidos)")
            self._text(
                "Quando aplicamos a correção anual mês a mês, os totais mudam porque "
                "as parcelas crescem ao longo do tempo:"
            )
            self._formula(f"Total F1 corrigido = somatório(Parcela F1 x Fator) = {fmt_currency(c['desemb1_corr'])}")
            self._formula(f"Total F2 corrigido = somatório(Parcela F2 x Fator) = {fmt_currency(c['desemb2_corr'])}")
            self._formula(f"TOTAL CORRIGIDO = {fmt_currency(c['total_des_corr'])}")
            self._result_box("Total Desembolsado (com correção)", fmt_currency(c["total_des_corr"]))
        else:
            self._result_box("Total Desembolsado", fmt_currency(c["total_des"]))

        self._text(
            f"As taxas e encargos totalizam {fmt_currency(c['total_tax'])} ao longo dos {pz} meses "
            f"({fmt_pct(c['total_tax'] / vc * 100 if vc > 0 else 0)} sobre o valor da carta)."
        )

        # ============ CUSTO EFETIVO ============
        ce_section = str(int(totais_section) + 1)
        self._title(f"{ce_section}. Custo Efetivo do Consórcio")

        self._text(
            "O custo efetivo mede quanto o consórcio custa em relação ao crédito recebido. "
            "É a forma mais direta de entender o preço real da operação: de cada real de crédito, "
            "quantos centavos são encargos?"
        )

        self._subtitle(f"{ce_section}.1 Custo total sobre o crédito")
        total_ref = c["total_des_corr"] if c["tem_correção"] else c["total_des"]
        custo_total = total_ref - c["carta_liq"]

        self._text(
            f"O custo total é a diferença entre tudo que você paga e o crédito que recebe:"
        )
        self._formula(f"Custo = Total Pago - Crédito = {fmt_currency(total_ref)} - {fmt_currency(c['carta_liq'])} = {fmt_currency(custo_total)}")
        self._formula(f"Custo % = Custo / Crédito = {fmt_currency(custo_total)} / {fmt_currency(c['carta_liq'])} = {fmt_pct(c['ce_pct'])}")
        self._text(
            f"Ou seja, para cada R$ 1,00 de crédito, você paga {fmt_pct(c['ce_pct'])} a mais em encargos "
            f"ao longo dos {pz} meses."
        )

        self._subtitle(f"{ce_section}.2 Taxa composta (anualizada)")
        self._text(
            "Para expressar esse custo como uma taxa anual composta (similar a como se expressa "
            "a rentabilidade de um investimento), calculamos:"
        )
        self._formula(f"Taxa mensal = (Total Pago / Crédito)^(1/Prazo) - 1")
        self._formula(f"Taxa mensal = ({fmt_currency(total_ref)} / {fmt_currency(c['carta_liq'])})^(1/{pz}) - 1 = {fmt_pct(c['ce_mensal'] * 100)} a.m.")
        self._formula(f"Taxa anual = (1 + Taxa mensal)^12 - 1 = {fmt_pct(c['ce_anual'] * 100)} a.a.")

        self._subtitle(f"{ce_section}.3 O que isso significa")
        if c["ce_anual"] > 0:
            self._text(
                f"A taxa de {fmt_pct(c['ce_anual'] * 100)} ao ano significa que o custo do consórcio "
                f"equivale a um encargo composto de {fmt_pct(c['ce_mensal'] * 100)} por mês sobre o "
                f"valor do crédito. Esse número permite ao cliente avaliar se o consórcio e a melhor "
                f"opção para o seu perfil, comparando com outras alternativas disponíveis."
            )
        else:
            self._text(
                "O custo efetivo é zero, indicando que o total pago não excede o crédito recebido."
            )

        relação = total_ref / c["carta_liq"] if c["carta_liq"] > 0 else 0

        self._result_box("Custo sobre o Crédito", fmt_pct(c["ce_pct"]))
        self._result_box("Taxa de Custo Anual", fmt_pct(c["ce_anual"] * 100))
        self._result_box("Taxa de Custo Mensal", fmt_pct(c["ce_mensal"] * 100))
        self._result_box("Crédito Líquido Recebido", fmt_currency(c["carta_liq"]))
        self._result_box("Total Desembolsado", fmt_currency(total_ref))
        self._result_box("Custo Total (encargos)", fmt_currency(custo_total))
        self._result_box("Relação Total Pago / Crédito", f"{relação:.2f}x".replace(".", ","))

        # ============ CONCLUSÃO ============
        concl_section = str(int(ce_section) + 1)
        self._title(f"{concl_section}. Conclusão")
        self._text(
            f"Para adquirir um bem de {fmt_currency(c['carta_liq'])} via consórcio, o cliente "
            f"desembolsará um total estimado de {fmt_currency(total_ref)} ao longo de {pz} meses "
            f"({pz / 12:.0f} anos), resultando em um custo adicional de {fmt_currency(custo_total)} "
            f"({fmt_pct(c['ce_pct'])} sobre o crédito)."
        )

        self._text(
            f"Em termos de taxa composta, isso representa {fmt_pct(c['ce_anual'] * 100)} ao ano. "
            f"Este percentual reflete o peso total dos encargos (taxa de administração, fundo de "
            f"reserva, seguro e eventuais correções) distribuídos ao longo de todo o prazo do grupo."
        )

        self._text(
            "IMPORTANTE: Esta simulação utiliza parâmetros estimados. Os valores reais podem "
            "variar conforme regras da administradora, comportamento do grupo e índices econômicos."
        )

        # Disclaimer
        self.ln(6)
        self.set_draw_color(*MID_GRAY)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(3)
        self.set_font("Helvetica", "I", 6)
        self.set_text_color(150, 150, 150)
        self.multi_cell(
            190, 3.5,
            "Relatório gerado automaticamente pela plataforma Somus Capital. "
            "Os cálculos seguem métodologia padrão de mercado para simulação de consórcio. "
            "A taxa de custo utiliza o método de taxa composta anualizada sobre o crédito líquido. "
            f"Gerado em {datetime.now().strftime('%d/%m/%Y as %H:%M')}.",
            align="C",
        )


def generate_relatorio_consorcio(dados, output_dir):
    """Gera PDF de relatório explicativo da simulação de consórcio."""
    os.makedirs(output_dir, exist_ok=True)

    pdf = RelatorioConsorcioPDF(dados)
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.build()

    nome = sanitize_text(dados.get("cliente_nome", "Cliente"))
    nome_limpo = nome.replace(" ", "_").replace("/", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Relatório_{nome_limpo}_{timestamp}.pdf"
    filepath = os.path.join(output_dir, filename)
    pdf.output(filepath)
    return filepath
