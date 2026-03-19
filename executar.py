"""
Somus Capital - Fluxo de Renda Fixa
App empresarial moderno.
"""

import os
import re
import base64
import threading
from datetime import datetime
from version import VERSION
from updater import verificar_atualizacao, baixar_e_instalar, reiniciar_app
from collections import defaultdict

import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
import openpyxl

import matplotlib
matplotlib.use('Agg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from gerar_pdf_consorcio import generate_consorcio_pdf, generate_relatorio_consorcio

# Import dos módulos internos
import sys
import tempfile
_SALDOS_MODULE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Mesa Produtos", "Envio Saldos")
if _SALDOS_MODULE not in sys.path:
    sys.path.insert(0, _SALDOS_MODULE)
_SEGUROS_MODULE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Seguros", "Renovacoes Anuais")
if _SEGUROS_MODULE not in sys.path:
    sys.path.insert(0, _SEGUROS_MODULE)

# Import do consolidador (projeto agente investimentos - na pasta Documents)
_DOCS_DIR = os.path.join(os.path.expanduser("~"), "Documents")
_AGENTE_PROJECT = os.path.join(_DOCS_DIR, "Projeto - AGENTE INVESTIMENTOS")
if os.path.isdir(_AGENTE_PROJECT) and _AGENTE_PROJECT not in sys.path:
    sys.path.insert(0, _AGENTE_PROJECT)

# === PATHS ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Shared
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo_somus.png")
BASE_FILE = os.path.join(BASE_DIR, "BASE", "BASE EMAILS.xlsx")

# Mesa Produtos
ENTRADA_FILE = os.path.join(BASE_DIR, "Mesa Produtos", "Fluxo RF", "ENTRADA", "Agenda de eventos.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "Mesa Produtos", "Fluxo RF", "PDFs")
AGIO_DIR = os.path.join(BASE_DIR, "Mesa Produtos", "Info Agio", "Agio")
RECEITA_FILE = os.path.join(BASE_DIR, "Mesa Produtos", "Ctrl Receita", "CONTROLE RECEITA TOTAL.xlsx")
ORGANIZADOR_OUTPUT_DIR = os.path.join(BASE_DIR, "Mesa Produtos", "Organizador", "ORGANIZADO")
CONSOLIDADOR_OUTPUT_DIR = os.path.join(BASE_DIR, "Mesa Produtos", "Consolidador", "PDFs")
SALDOS_DIR = os.path.join(BASE_DIR, "Mesa Produtos", "Envio Saldos")
ANIVERSARIOS_DIR = os.path.join(BASE_DIR, "Mesa Produtos", "Envio Aniversarios")

# Corporate
CONSÓRCIO_OUTPUT_DIR = os.path.join(BASE_DIR, "Corporate", "Simulador", "PDFs")

# Seguros
SEGUROS_DIR = os.path.join(os.path.expanduser("~"), "SOMUS CAPITAL ASSESSOR DE INVESTIMENTOS LTDA", "Somus Capital - Seguros - Documentos")

# === PALETA SOMUS ===
BG_PRIMARY = "#f7f8fa"
BG_SECONDARY = "#ffffff"
BG_SIDEBAR = "#004d33"
BG_SIDEBAR_HOVER = "#006644"
BG_SIDEBAR_ACTIVE = "#00835a"
BG_HEADER = "#ffffff"
BG_CARD = "#ffffff"
BG_LOG = "#1a1f2e"
BG_LOG_HEADER = "#141824"
BG_PROGRESS_TRACK = "#e4e7ec"
BG_INPUT = "#f2f4f7"

ACCENT_GREEN = "#004d33"
ACCENT_BLUE = "#1863DC"
ACCENT_TEAL = "#00a86b"
ACCENT_ORANGE = "#e6832a"
ACCENT_PURPLE = "#7c5cbf"
ACCENT_RED = "#dc3545"

TEXT_PRIMARY = "#111827"
TEXT_SECONDARY = "#4b5563"
TEXT_TERTIARY = "#9ca3af"
TEXT_WHITE = "#ffffff"
TEXT_SIDEBAR = "#c8e6d8"
TEXT_SIDEBAR_ACTIVE = "#ffffff"
TEXT_LOG = "#c8d0dc"

BORDER_LIGHT = "#e5e7eb"
BORDER_CARD = "#e8ecf0"
SHADOW_CARD = "#d1d5db"


# === CTRL RECEITA: DETAIL SHEETS MAPPING ===
CTRL_DETAIL_SHEETS = [
    {"produto": "Sec RF", "sheet": "SECUND\u00c1RIO RF", "col_assessor": 3, "col_receita": 14},
    {"produto": "Prim RF", "sheet": "PRIMARIO RF", "col_assessor": 2, "col_receita": None},
    {"produto": "RV Corretagem", "sheet": "RENDA VARIAVEL | CORRETAGEM", "col_assessor": 4, "col_receita": 8},
    {"produto": "RV Estruturadas", "sheet": "RENDA VARIAVEL | ESTRUTURADAS", "col_assessor": 9, "col_receita": 7},
    {"produto": "Cambio", "sheet": "C\u00c2MBIO", "col_assessor": 7, "col_receita": 15},
    {"produto": "Fundos Fechados", "sheet": "FUNDOS FECHADOS", "col_assessor": 10, "col_receita": 9},
    {"produto": "Sec Fundos", "sheet": "SECUNDARIO FUNDOS", "col_assessor": 4, "col_receita": 15},
    {"produto": "COE", "sheet": "COE", "col_assessor": 2, "col_receita": 11},
]


# === DATA HELPERS ===

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
    return text


def load_base_emails():
    wb = openpyxl.load_workbook(BASE_FILE)
    ws = wb[wb.sheetnames[0]]
    assessores = {}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if not row[0]:
            continue
        código = str(row[0]).strip()
        if not código.startswith("A"):
            continue
        assessores[código] = {
            "nome": sanitize_text(row[1]) if row[1] else "",
            "email": str(row[2]).strip() if row[2] else "",
            "assistente": sanitize_text(row[3]) if row[3] else "-",
            "email_assistente": str(row[4]).strip() if row[4] else "-",
        }
    wb.close()
    return assessores


def load_equipe_mapping():
    """Carrega mapeamento codigo_assessor -> equipe da planilha CONTROLE RECEITA."""
    mapping = {}
    try:
        if not os.path.exists(RECEITA_FILE):
            return mapping
        wb = openpyxl.load_workbook(RECEITA_FILE, data_only=True)
        ws = wb["CONSOLIDADO"]
        for r in range(7, ws.max_row + 1):
            cod = ws.cell(r, 2).value
            if not cod:
                continue
            equipe = sanitize_text(ws.cell(r, 3).value)
            nome = sanitize_text(ws.cell(r, 4).value).upper()
            # Agrupar PRODUTOS e CORPORATE dentro de SP
            if equipe in ("PRODUTOS", "CORPORATE"):
                equipe = "SP"
            # Reagrupar BACKOFFICE individualmente
            if equipe == "BACKOFFICE":
                if "JULIA" in nome:
                    equipe = "LEBLON"
                elif "GUILHERME" in nome and "PRAXEDES" in nome:
                    equipe = "SP"
                else:
                    equipe = "SP"  # fallback para outros backoffice
            cod_str = str(int(cod)) if isinstance(cod, (int, float)) else str(cod).strip()
            mapping[f"A{cod_str}"] = equipe
        wb.close()
    except Exception:
        pass
    return mapping


def load_eventos_summary():
    wb = openpyxl.load_workbook(ENTRADA_FILE)
    ws = wb[wb.sheetnames[0]]
    total_eventos = 0
    assessores_set = set()
    ativos = set()
    valor_total = 0.0
    datas = set()
    tipos = defaultdict(lambda: {"count": 0, "total": 0.0})
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        cod = row[1]
        if not cod or not isinstance(cod, str) or not cod.startswith("A"):
            continue
        data = row[0]
        if not isinstance(data, datetime):
            continue
        total_eventos += 1
        assessores_set.add(cod)
        if row[8]:
            ativos.add(row[8])
        val = float(row[7]) if row[7] else 0
        valor_total += val
        datas.add(data)
        tipo = sanitize_text(row[9]) if row[9] else "OUTRO"
        tipos[tipo]["count"] += 1
        tipos[tipo]["total"] += val
    wb.close()
    datas_sorted = sorted(datas) if datas else []
    return {
        "total_eventos": total_eventos,
        "assessores": len(assessores_set),
        "ativos": len(ativos),
        "valor_total": valor_total,
        "data_min": datas_sorted[0] if datas_sorted else None,
        "data_max": datas_sorted[-1] if datas_sorted else None,
        "datas_unicas": len(datas_sorted),
        "tipos": dict(tipos),
    }


def find_pdf_for_assessor(código, nome):
    nome_limpo = nome.replace(" ", "_").replace("/", "_")
    expected = f"Fluxo_RF_{código}_{nome_limpo}.pdf"
    path = os.path.join(OUTPUT_DIR, expected)
    if os.path.exists(path):
        return path
    if os.path.exists(OUTPUT_DIR):
        for f in os.listdir(OUTPUT_DIR):
            if f.startswith(f"Fluxo_RF_{código}_") and f.endswith(".pdf"):
                return os.path.join(OUTPUT_DIR, f)
    return None


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


def _load_logo_b64():
    """Carrega logo Somus como base64 para embedding inline no email."""
    if not os.path.exists(LOGO_PATH):
        return ""
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode()


# Tag HTML referenciando a logo via CID (funciona no Outlook)
LOGO_CID = "somus_logo"
LOGO_TAG_CID = (
    f'<img src="cid:{LOGO_CID}" width="155" height="38"'
    f' style="vertical-align:middle;margin-right:16px;" alt="Somus Capital">'
)


def _attach_logo_cid(mail):
    """Anexa logo como imagem oculta com CID para exibir no corpo HTML."""
    if not os.path.exists(LOGO_PATH):
        return
    att = mail.Attachments.Add(os.path.abspath(LOGO_PATH))
    # PR_ATTACH_CONTENT_ID (DASL property)
    att.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        LOGO_CID)
    # PR_ATTACHMENT_HIDDEN — esconde da lista de anexos
    att.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
        True)


def build_email_body(nome_assessor):
    hoje = datetime.now().strftime("%d/%m/%Y")
    primeiro_nome = nome_assessor.split()[0] if nome_assessor else "Assessor"

    logo_tag = LOGO_TAG_CID if os.path.exists(LOGO_PATH) else ""

    return f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:960px;margin:0 auto;">

<!-- Header -->
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding:16px 0 12px 0;vertical-align:middle;">
      {logo_tag}<!--
      --><span style="font-size:16pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:300;"> | </span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:bold;">Produtos</span>
    </td>
  </tr>
  <tr>
    <td style="padding:0 0 16px 0;">
      <hr style="border:none;border-top:2px solid #004d33;margin:0;">
    </td>
  </tr>
</table>

<!-- Corpo -->
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px;">
  <tr><td style="padding:0 4px;">

    <p style="font-size:11pt;color:#1a1a2e;margin-bottom:4px;">
      Prezado(a) <b>{primeiro_nome}</b>, tudo bem?
    </p>
    <p style="font-size:10.5pt;color:#4b5563;margin-top:0;">
      Segue em anexo o relat&oacute;rio de <b>Fluxo de Renda Fixa</b> com os eventos previstos para os pr&oacute;ximos dias.
    </p>

  </td></tr>

  <!-- Secao: Conteudo do relatório -->
  <tr><td style="padding:22px 0 8px 0;">
    <table cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background-color:#00b876;width:4px;border-radius:4px;">&nbsp;</td>
      <td style="padding-left:12px;">
        <span style="font-family:Calibri,Arial,sans-serif;font-size:12.5pt;color:#004d33;font-weight:bold;letter-spacing:0.3px;">Conte&uacute;do do Relat&oacute;rio</span>
      </td>
    </tr></table>
  </td></tr>

  <tr><td style="padding:0 4px;">
    <table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:9.5pt;margin-bottom:6px;">
      <tr style="background:#f7faf9;">
        <td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1.5px solid #00785a;white-space:nowrap;">&#8226;</td>
        <td style="padding:5px 12px;color:#1a1a2e;border-bottom:1.5px solid #00785a;">Calend&aacute;rio com todas as datas de eventos</td>
      </tr>
      <tr style="background:#ffffff;">
        <td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1px solid #eef1ef;">&#8226;</td>
        <td style="padding:5px 12px;color:#1a1a2e;border-bottom:1px solid #eef1ef;">Detalhamento por data com ativo, tipo, quantidade e valor</td>
      </tr>
      <tr style="background:#f7faf9;">
        <td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1px solid #eef1ef;">&#8226;</td>
        <td style="padding:5px 12px;color:#1a1a2e;border-bottom:1px solid #eef1ef;">Resumo consolidado por ativo</td>
      </tr>
    </table>
  </td></tr>

  <tr><td style="padding:14px 4px 0 4px;">
    <p style="font-size:10.5pt;color:#4b5563;margin:0;">
      Em caso de d&uacute;vidas, estamos &agrave; disposi&ccedil;&atilde;o.
    </p>
  </td></tr>

  <!-- Footer -->
  <tr><td style="padding:28px 0 0 0;">
    <hr style="border:none;border-top:2px solid #004d33;margin:0 0 12px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="vertical-align:middle;">
          <span style="font-size:10pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:300;"> | </span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:bold;">Produtos</span>
        </td>
        <td style="text-align:right;">
          <span style="font-size:8.5pt;color:#6b7280;">
            Fluxo de Renda Fixa &middot; {hoje}
          </span><br>
          <span style="font-size:7.5pt;color:#9ca3af;">
            E-mail gerado automaticamente
          </span>
        </td>
      </tr>
    </table>
  </td></tr>

</table>
</div>
"""


def build_email_consorcio(nome_cliente, dados_resumo):
    hoje = datetime.now().strftime("%d/%m/%Y")
    return f"""<html><head><style>
body{{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#2d2d2d}}
.hb{{background:#004d33;padding:12px 20px}}.hb span{{color:#fff;font-size:14pt;font-weight:bold}}
.ct{{padding:15px 20px}}.ft{{padding:10px 20px;font-size:9pt;color:#888;border-top:1px solid #e0e0e0;margin-top:20px}}
.hl{{color:#004d33;font-weight:bold}}
table{{border-collapse:collapse;width:100%;margin:10px 0}}
td{{padding:6px 10px;border-bottom:1px solid #eee;font-size:10pt}}
td.lb{{color:#666;width:55%}}td.vl{{font-weight:bold;color:#2d2d2d}}
.tag{{display:inline-block;background:#004d33;color:#fff;padding:3px 10px;border-radius:4px;font-size:9pt;margin-bottom:8px}}
</style></head><body>
<div class="hb"><span>Somus Capital</span></div>
<div class="ct">
<p>Prezado(a) <span class="hl">{nome_cliente}</span>,</p>
<p>Segue em anexo a <b>simula&ccedil;&atilde;o de cons&oacute;rcio</b> elaborada pela Somus Capital.</p>
<span class="tag">Cons&oacute;rcio - {dados_resumo.get('tipo_bem', '')}</span>
<table>
<tr><td class="lb">Carta de Cr&eacute;dito</td><td class="vl">{dados_resumo.get('valor_carta_fmt', '')}</td></tr>
<tr><td class="lb">Administradora</td><td class="vl">{dados_resumo.get('administradora', '')}</td></tr>
<tr><td class="lb">Prazo do Grupo</td><td class="vl">{dados_resumo.get('prazo', '')} meses</td></tr>
<tr><td class="lb">Parcela Fase 1</td><td class="vl">{dados_resumo.get('parcela_f1', '')}/m&ecirc;s</td></tr>
<tr><td class="lb">Custo Efetivo</td><td class="vl">{dados_resumo.get('ce_anual', '')} a.a.</td></tr>
</table>
<p>Os documentos anexos cont&ecirc;m:</p>
<ul>
<li><b>Proposta de Simula&ccedil;&atilde;o</b> - valores detalhados, cronograma e resumo financeiro</li>
<li><b>Relat&oacute;rio Explicativo</b> - métodologia de c&aacute;lculo passo a passo</li>
</ul>
<p>Fico &agrave; disposi&ccedil;&atilde;o para esclarecer qualquer d&uacute;vida.</p>
<p>Atenciosamente,<br><span class="hl">Equipe Somus Capital</span><br>Cons&oacute;rcios</p>
</div>
<div class="ft">Simula&ccedil;&atilde;o gerada em {hoje}. Valores estimados, sujeitos a altera&ccedil;&atilde;o pela administradora.</div>
</body></html>"""


# =====================================================================
#  APP
# =====================================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Somus Capital")
        self.geometry("1100x720")
        self.minsize(1000, 650)
        self.configure(fg_color=BG_PRIMARY)
        self._set_icon()

        self.current_page = "dashboard"
        self.current_module = "Mesa Produtos"
        self.sidebar_buttons = {}
        self.mp_nav_keys = []
        self.cs_nav_keys = []
        self.sg_nav_keys = []
        self.pages = {}

        # Grid: sidebar | content
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_pages()
        self._show_page("dashboard")
        self._load_data_async()
        # Force initial module state after UI is ready
        self.after(50, lambda: self._on_module_change("Mesa Produtos"))

        # Verificar atualizacoes em background (4s de delay para nao travar o carregamento)
        self.after(4000, self._checar_atualizacao)

    def _set_icon(self):
        try:
            ico_path = os.path.join(BASE_DIR, "assets", "icon_somus.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
        except Exception:
            pass

    # =================================================================
    #  AUTO UPDATE
    # =================================================================

    def _checar_atualizacao(self):
        def _on_disponivel(versao, url):
            self.after(0, lambda: self._mostrar_update_dialog(versao, url))
        verificar_atualizacao(_on_disponivel)

    def _mostrar_update_dialog(self, versao, url):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Atualizacao disponivel")
        dialog.geometry("420x220")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.configure(fg_color=BG_SECONDARY)
        dialog.after(200, dialog.lift)

        ctk.CTkLabel(
            dialog, text=f"Nova versao disponivel: v{versao}",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=ACCENT_GREEN
        ).pack(pady=(28, 4))
        ctk.CTkLabel(
            dialog, text="Deseja atualizar agora? O app sera reiniciado.",
            font=ctk.CTkFont(size=12), text_color=TEXT_SECONDARY
        ).pack(pady=(0, 16))

        barra = ctk.CTkProgressBar(dialog, width=340)
        barra.set(0)

        label_status = ctk.CTkLabel(dialog, text="", font=ctk.CTkFont(size=11), text_color=TEXT_TERTIARY)

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=4)

        def _iniciar():
            btn_sim.configure(state="disabled")
            btn_nao.configure(state="disabled")
            barra.pack(pady=(4, 2))
            label_status.pack()

            def _progresso(pct, msg):
                self.after(0, lambda: barra.set(pct / 100))
                self.after(0, lambda: label_status.configure(text=msg))

            def _concluido():
                self.after(0, lambda: label_status.configure(text="Reiniciando..."))
                self.after(800, reiniciar_app)

            def _erro(msg):
                self.after(0, lambda: label_status.configure(
                    text=f"Erro: {msg}", text_color=ACCENT_RED
                ))
                self.after(0, lambda: btn_nao.configure(state="normal", text="Fechar"))

            baixar_e_instalar(url, _progresso, _concluido, _erro)

        btn_sim = ctk.CTkButton(
            btn_frame, text="Atualizar agora", width=160,
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            command=_iniciar
        )
        btn_sim.pack(side="left", padx=8)

        btn_nao = ctk.CTkButton(
            btn_frame, text="Agora nao", width=120,
            fg_color=BG_INPUT, hover_color=BORDER_LIGHT,
            text_color=TEXT_PRIMARY, command=dialog.destroy
        )
        btn_nao.pack(side="left", padx=8)

    # =================================================================
    #  SIDEBAR
    # =================================================================
    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, fg_color=BG_SIDEBAR, width=245, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)
        sidebar.grid_rowconfigure(20, weight=1)
        self._sidebar = sidebar  # prevent GC

        # Logo (row 0)
        logo_frame = ctk.CTkFrame(sidebar, fg_color="transparent", height=78)
        logo_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        logo_frame.grid_propagate(False)
        self._logo_frame = logo_frame  # prevent GC

        try:
            logo_img = Image.open(LOGO_PATH)
            self._logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(155, 38))
            ctk.CTkLabel(logo_frame, image=self._logo_ctk, text="").place(relx=0.5, rely=0.5, anchor="center")
        except Exception:
            ctk.CTkLabel(
                logo_frame, text="SOMUS CAPITAL",
                font=("Segoe UI", 16, "bold"), text_color=TEXT_WHITE
            ).place(relx=0.5, rely=0.5, anchor="center")

        # Separador (row 1)
        ctk.CTkFrame(sidebar, fg_color="#006644", height=1, corner_radius=0).grid(
            row=1, column=0, sticky="ew", padx=18
        )

        # ---- SELETOR DE MODULO (row 2) ----
        selector_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        selector_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(12, 10))
        self._selector_frame = selector_frame  # prevent GC

        ctk.CTkLabel(
            selector_frame, text="MÓDULO",
            font=("Segoe UI", 9, "bold"), text_color="#5a9a7a", anchor="w"
        ).pack(fill="x", padx=4, pady=(0, 4))

        self.module_selector = ctk.CTkOptionMenu(
            selector_frame,
            values=["Mesa Produtos", "Corporate", "Seguros"],
            font=("Segoe UI", 12, "bold"),
            fg_color="#003d2b",
            button_color=ACCENT_BLUE,
            button_hover_color="#1555bb",
            dropdown_fg_color="#002a1a",
            dropdown_hover_color=BG_SIDEBAR_HOVER,
            dropdown_text_color=TEXT_WHITE,
            text_color=TEXT_WHITE,
            corner_radius=8,
            height=36,
            anchor="w",
            command=self._on_module_change,
        )
        self.module_selector.pack(fill="x")
        self.module_selector.set("Mesa Produtos")

        # Separador 2 (row 3)
        ctk.CTkFrame(sidebar, fg_color="#006644", height=1, corner_radius=0).grid(
            row=3, column=0, sticky="ew", padx=18
        )

        # Secao label (row 4)
        ctk.CTkLabel(
            sidebar, text="MENU", font=("Segoe UI", 9, "bold"),
            text_color="#5a9a7a", anchor="w"
        ).grid(row=4, column=0, sticky="w", padx=24, pady=(12, 6))

        # === MESA PRODUTOS nav buttons (rows 5-7) ===
        mp_nav_items = [
            ("dashboard", "Dashboard", "\u25a3"),
            ("fluxo_rf", "FLUXO - RF", "\u2913"),
            ("informativo", "Informativo", "\u2709"),
            ("info_agio", "Info - Agio", "\u25b2"),
            ("envio_ordens", "Envio de Ordens", "\u2191"),
            ("ctrl_receita", "Ctrl Receita", "$"),
            ("organizador", "Organizador", "\u25a6"),
            ("consolidador", "Consolidador", "\u229a"),
            ("saldos", "Envio Saldos", "\u21c6"),
            ("envio_mesa", "Envio Mesa", "\u2709"),
            ("envio_aniversarios", "Envio Aniversários", "\u2605"),
            ("tarefas", "Tarefas", "\u2611"),
        ]
        for i, (key, label, icon) in enumerate(mp_nav_items):
            btn = ctk.CTkButton(
                sidebar,
                text=f"  {icon}    {label}",
                font=("Segoe UI", 13),
                fg_color="transparent",
                hover_color=BG_SIDEBAR_HOVER,
                text_color=TEXT_SIDEBAR,
                anchor="w",
                height=42,
                corner_radius=8,
                command=lambda k=key: self._on_nav(k),
            )
            btn.grid(row=5 + i, column=0, sticky="ew", padx=12, pady=2)
            self.sidebar_buttons[key] = btn
            self.mp_nav_keys.append(key)

        # === CORPORATE nav buttons (rows 5-6, ocultos inicialmente) ===
        cs_nav_items = [
            ("corp_dashboard", "Dashboard", "\u25a3"),
            ("simulador", "Simulador", "\u2630"),
        ]
        for i, (key, label, icon) in enumerate(cs_nav_items):
            btn = ctk.CTkButton(
                sidebar,
                text=f"  {icon}    {label}",
                font=("Segoe UI", 13),
                fg_color="transparent",
                hover_color=BG_SIDEBAR_HOVER,
                text_color=TEXT_SIDEBAR,
                anchor="w",
                height=42,
                corner_radius=8,
                command=lambda k=key: self._on_nav(k),
            )
            btn.grid(row=5 + i, column=0, sticky="ew", padx=12, pady=2)
            btn.grid_remove()
            self.sidebar_buttons[key] = btn
            self.cs_nav_keys.append(key)

        # === SEGUROS nav buttons (rows 5+, ocultos inicialmente) ===
        sg_nav_items = [
            ("seg_renovacoes", "Renovações Anuais", "\u26e8"),
        ]
        for i, (key, label, icon) in enumerate(sg_nav_items):
            btn = ctk.CTkButton(
                sidebar,
                text=f"  {icon}    {label}",
                font=("Segoe UI", 13),
                fg_color="transparent",
                hover_color=BG_SIDEBAR_HOVER,
                text_color=TEXT_SIDEBAR,
                anchor="w",
                height=42,
                corner_radius=8,
                command=lambda k=key: self._on_nav(k),
            )
            btn.grid(row=5 + i, column=0, sticky="ew", padx=12, pady=2)
            btn.grid_remove()
            self.sidebar_buttons[key] = btn
            self.sg_nav_keys.append(key)

        # Spacer (row 20 tem weight=1)

        # Separator inferior (row 21)
        ctk.CTkFrame(sidebar, fg_color="#006644", height=1, corner_radius=0).grid(
            row=21, column=0, sticky="ew", padx=18, pady=(0, 0)
        )

        # Reportar Erro (row 22) - visivel em ambos os modulos
        ctk.CTkButton(
            sidebar,
            text="  \u26a0    Reportar Erro",
            font=("Segoe UI", 11),
            fg_color="transparent",
            hover_color="#4d1a1a",
            text_color="#e07070",
            anchor="w",
            height=36,
            corner_radius=8,
            command=self._on_reportar_erro,
        ).grid(row=22, column=0, sticky="ew", padx=12, pady=(8, 4))

        # Versao (row 23)
        ctk.CTkLabel(
            sidebar, text="SomusApp - BETA",
            font=("Segoe UI", 8), text_color="#3a6a50"
        ).grid(row=23, column=0, pady=(4, 14))

    def _on_nav(self, page_key):
        if page_key == "simulador":
            self._show_page("consorcio")
        elif page_key == "corp_dashboard":
            self._show_page("corp_dashboard")
        elif page_key == "seg_renovacoes":
            self._show_page("seg_renovacoes")
        else:
            self._show_page(page_key)

    def _on_module_change(self, value):
        self.current_module = value
        all_nav = [
            (self.mp_nav_keys, "Mesa Produtos", "dashboard"),
            (self.cs_nav_keys, "Corporate", "corp_dashboard"),
            (self.sg_nav_keys, "Seguros", "seg_renovacoes"),
        ]
        for keys, module, default_page in all_nav:
            if value == module:
                for key in keys:
                    self.sidebar_buttons[key].grid()
            else:
                for key in keys:
                    self.sidebar_buttons[key].grid_remove()
        default_pages = {"Mesa Produtos": "dashboard", "Corporate": "corp_dashboard", "Seguros": "seg_renovacoes"}
        self._show_page(default_pages.get(value, "dashboard"))

    def _update_sidebar_active(self, active_key):
        for key, btn in self.sidebar_buttons.items():
            if key == active_key:
                btn.configure(fg_color=BG_SIDEBAR_ACTIVE, text_color=TEXT_WHITE)
            else:
                btn.configure(fg_color="transparent", text_color=TEXT_SIDEBAR)

    # =================================================================
    #  TOPBAR HELPER
    # =================================================================
    def _make_topbar(self, parent, title, subtitle="", show_date=True,
                     back_btn=False, back_cmd=None):
        """Cria topbar profissional padronizado com logo no canto direito."""
        wrapper = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        wrapper.pack(fill="x")

        topbar = ctk.CTkFrame(wrapper, fg_color=BG_HEADER, height=68, corner_radius=0)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        # --- Lado esquerdo: título + subtítulo ---
        left = ctk.CTkFrame(topbar, fg_color="transparent")
        left.pack(side="left", fill="y", padx=28)

        title_lbl = ctk.CTkLabel(
            left, text=title,
            font=("Segoe UI", 21, "bold"), text_color=TEXT_PRIMARY
        )
        title_lbl.pack(side="left", anchor="w", pady=0)

        if subtitle:
            sep_lbl = ctk.CTkLabel(
                left, text="  |  ",
                font=("Segoe UI", 13), text_color=BORDER_CARD
            )
            sep_lbl.pack(side="left", anchor="w")
            sub_lbl = ctk.CTkLabel(
                left, text=subtitle,
                font=("Segoe UI", 12), text_color=TEXT_TERTIARY
            )
            sub_lbl.pack(side="left", anchor="w")

        # --- Lado direito: logo + data ---
        right = ctk.CTkFrame(topbar, fg_color="transparent")
        right.pack(side="right", fill="y", padx=24)

        try:
            logo_img = Image.open(LOGO_PATH)
            logo_topbar = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(140, 34))
            logo_lbl = ctk.CTkLabel(right, image=logo_topbar, text="")
            logo_lbl.pack(side="right", anchor="e", padx=(16, 4))
            logo_lbl._topbar_logo = logo_topbar  # manter referencia
        except Exception:
            pass

        if show_date:
            date_lbl = ctk.CTkLabel(
                right, text=datetime.now().strftime("%d/%m/%Y   %H:%M"),
                font=("Segoe UI", 10), text_color=TEXT_TERTIARY
            )
            date_lbl.pack(side="right", anchor="e", padx=(0, 8))

        if back_btn and back_cmd:
            ctk.CTkButton(
                right, text="\u2190  Voltar",
                font=("Segoe UI", 10), fg_color=BG_INPUT,
                hover_color=BORDER_CARD, text_color=TEXT_SECONDARY,
                height=30, corner_radius=8, width=90,
                command=back_cmd
            ).pack(side="right", anchor="e", padx=(0, 12))

        # --- Filete verde embaixo ---
        accent = ctk.CTkFrame(wrapper, fg_color=ACCENT_GREEN, height=3, corner_radius=0)
        accent.pack(fill="x")

        return wrapper, title_lbl, date_lbl if show_date else None

    # =================================================================
    #  PAGES
    # =================================================================
    def _build_pages(self):
        self.pages["dashboard"] = self._build_receita_dashboard_page()
        self.pages["fluxo_rf"] = self._build_dashboard_page()
        self.pages["operations"] = self._build_operations_page()
        self.pages["informativo"] = self._build_informativo_page()
        self.pages["info_agio"] = self._build_info_agio_page()
        self.pages["envio_ordens"] = self._build_envio_ordens_page()
        self.pages["consorcio"] = self._build_consorcio_page()
        self.pages["corp_dashboard"] = self._build_corp_dashboard_page()
        self.pages["ctrl_receita"] = self._build_ctrl_receita_page()
        self.pages["organizador"] = self._build_organizador_page()
        self.pages["consolidador"] = self._build_consolidador_page()
        self.pages["saldos"] = self._build_saldos_page()
        self.pages["envio_mesa"] = self._build_envio_mesa_page()
        self.pages["envio_aniversarios"] = self._build_envio_aniversarios_page()
        self.pages["tarefas"] = self._build_tarefas_page()
        self.pages["seg_renovacoes"] = self._build_renovacoes_page()

    def _show_page(self, page_key):
        for key, frame in self.pages.items():
            frame.grid_remove()
        self.pages[page_key].grid(row=0, column=1, sticky="nsew")
        self.current_page = page_key

        # Lazy load ctrl_receita data on first visit
        if page_key == "ctrl_receita" and not self._cr_loading and self._cr_data is None:
            self._cr_loading = True
            threading.Thread(target=self._load_ctrl_receita_data, daemon=True).start()

        # Lazy load info_agio data on first visit
        if page_key == "info_agio" and not getattr(self, "_ag_loaded", False):
            self._ag_loaded = True
            threading.Thread(target=self._ag_load_data_thread, daemon=True).start()

        nav_map = {
            "dashboard": "dashboard",
            "fluxo_rf": "fluxo_rf",
            "ctrl_receita": "ctrl_receita",
            "operations": "fluxo_rf",
            "corp_dashboard": "corp_dashboard",
            "consorcio": "simulador",
            "seg_renovacoes": "seg_renovacoes",
        }
        self._update_sidebar_active(nav_map.get(page_key, page_key))

    # -----------------------------------------------------------------
    #  PAGE: RECEITA DASHBOARD
    # -----------------------------------------------------------------
    def _build_receita_dashboard_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        _, _, _ = self._make_topbar(page, "Dashboard", subtitle="Mesa de Produtos")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Status banner
        self.rc_banner = ctk.CTkFrame(content, fg_color=ACCENT_GREEN, corner_radius=10, height=44)
        self.rc_banner.pack(fill="x", pady=(0, 18))
        self.rc_banner.pack_propagate(False)

        self.rc_banner_text = ctk.CTkLabel(
            self.rc_banner, text="  Carregando dados de receita...",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        )
        self.rc_banner_text.pack(side="left", padx=18, pady=10)

        self.rc_banner_date = ctk.CTkLabel(
            self.rc_banner, text="",
            font=("Segoe UI", 11), text_color="#80c0a0"
        )
        self.rc_banner_date.pack(side="right", padx=18)

        # RECEITA TOTAL destaque
        total_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        total_card.pack(fill="x", pady=(0, 18))
        total_inner = ctk.CTkFrame(total_card, fg_color="transparent")
        total_inner.pack(fill="x", padx=24, pady=16)

        ctk.CTkLabel(
            total_inner, text="RECEITA TOTAL",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY
        ).pack(anchor="w")

        self.rc_total_label = ctk.CTkLabel(
            total_inner, text="Carregando...",
            font=("Segoe UI", 34, "bold"), text_color=ACCENT_GREEN
        )
        self.rc_total_label.pack(anchor="w", pady=(2, 0))

        self.rc_total_sub = ctk.CTkLabel(
            total_inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY
        )
        self.rc_total_sub.pack(anchor="w")

        # KPI cards - 4 maiores categorias
        kpi_frame = ctk.CTkFrame(content, fg_color="transparent")
        kpi_frame.pack(fill="x", pady=(0, 8))
        kpi_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.rc_kpi_rf = self._make_kpi_card(kpi_frame, "Renda Fixa", "...", ACCENT_GREEN, 0)
        self.rc_kpi_rv_corr = self._make_kpi_card(kpi_frame, "RV Corretagem", "...", ACCENT_BLUE, 1)
        self.rc_kpi_fundos = self._make_kpi_card(kpi_frame, "Fundos Fechados", "...", ACCENT_PURPLE, 2)
        self.rc_kpi_rv_pe = self._make_kpi_card(kpi_frame, "RV Estruturadas", "...", ACCENT_ORANGE, 3)

        # KPI row 2 - menores categorias
        kpi_frame2 = ctk.CTkFrame(content, fg_color="transparent")
        kpi_frame2.pack(fill="x", pady=(0, 18))
        kpi_frame2.columnconfigure((0, 1, 2), weight=1)

        self.rc_kpi_cambio = self._make_kpi_card(kpi_frame2, "Cambio", "...", ACCENT_TEAL, 0)
        self.rc_kpi_coe = self._make_kpi_card(kpi_frame2, "COE", "...", ACCENT_RED, 1)
        self.rc_kpi_assessores = self._make_kpi_card(kpi_frame2, "Assessores", "...", "#6b7280", 2)

        # Distribuição por Produto (barra horizontal visual)
        ctk.CTkLabel(
            content, text="Distribuição por Produto",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        self.rc_bars_frame = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        self.rc_bars_frame.pack(fill="x", pady=(0, 18))

        self.rc_bars_inner = ctk.CTkFrame(self.rc_bars_frame, fg_color="transparent")
        self.rc_bars_inner.pack(fill="x", padx=20, pady=16)

        # Ranking Assessores
        ctk.CTkLabel(
            content, text="Ranking por Assessor",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        self.rc_ranking_frame = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        self.rc_ranking_frame.pack(fill="x", pady=(0, 18))

        self.rc_ranking_inner = ctk.CTkFrame(self.rc_ranking_frame, fg_color="transparent")
        self.rc_ranking_inner.pack(fill="x", padx=20, pady=16)

        # Placeholder
        self.rc_ranking_placeholder = ctk.CTkLabel(
            self.rc_ranking_inner, text="Carregando ranking...",
            font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.rc_ranking_placeholder.pack(anchor="w")

        # Ranking por Equipe
        ctk.CTkLabel(
            content, text="Receita por Equipe",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        self.rc_equipe_frame = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        self.rc_equipe_frame.pack(fill="x", pady=(0, 18))

        self.rc_equipe_inner = ctk.CTkFrame(self.rc_equipe_frame, fg_color="transparent")
        self.rc_equipe_inner.pack(fill="x", padx=20, pady=16)

        # Botao Atualizar
        self._make_atualizar_btn(content, RECEITA_FILE)

        # Load data async (delay to ensure mainloop is running)
        self.after(300, lambda: threading.Thread(target=self._load_receita_data, daemon=True).start())

        return page

    def _load_receita_data(self):
        """Carrega dados de CONTROLE RECEITA TOTAL.xlsx e atualiza o dashboard."""
        try:
            if not os.path.exists(RECEITA_FILE):
                self.after(0, lambda: self.rc_banner_text.configure(
                    text="  Arquivo CONTROLE RECEITA TOTAL.xlsx não encontrado"))
                return

            wb = openpyxl.load_workbook(RECEITA_FILE, data_only=True)
            ws = wb["CONSOLIDADO"]

            # Data base
            data_base = ws.cell(3, 2).value
            if data_base and hasattr(data_base, "strftime"):
                data_str = data_base.strftime("%d/%m/%Y")
            else:
                data_str = str(data_base) if data_base else "N/A"

            # Headers na row 6 (B=2..K=11)
            # B=Assessor, C=Equipe, D=Nome, E=RF, F=RV Corr, G=RV PE, H=Cambio, I=Fundos, J=COE, K=Receita Geral
            assessores = []
            for r in range(7, ws.max_row + 1):
                cod = ws.cell(r, 2).value
                if not cod:
                    continue
                nome = sanitize_text(ws.cell(r, 4).value)
                equipe = sanitize_text(ws.cell(r, 3).value)
                rf = float(ws.cell(r, 5).value or 0)
                rv_corr = float(ws.cell(r, 6).value or 0)
                rv_pe = float(ws.cell(r, 7).value or 0)
                cambio = float(ws.cell(r, 8).value or 0)
                fundos = float(ws.cell(r, 9).value or 0)
                coe = float(ws.cell(r, 10).value or 0)
                receita = float(ws.cell(r, 11).value or 0)
                assessores.append({
                    "cod": str(cod), "nome": nome, "equipe": equipe,
                    "rf": rf, "rv_corr": rv_corr, "rv_pe": rv_pe,
                    "cambio": cambio, "fundos": fundos, "coe": coe,
                    "receita": receita,
                })
            wb.close()

            # Totais por categoria
            total_rf = sum(a["rf"] for a in assessores)
            total_rv_corr = sum(a["rv_corr"] for a in assessores)
            total_rv_pe = sum(a["rv_pe"] for a in assessores)
            total_cambio = sum(a["cambio"] for a in assessores)
            total_fundos = sum(a["fundos"] for a in assessores)
            total_coe = sum(a["coe"] for a in assessores)
            total_geral = sum(a["receita"] for a in assessores)
            n_assessores = len([a for a in assessores if a["receita"] > 0])

            # Equipes
            equipes = defaultdict(float)
            for a in assessores:
                if a["receita"] > 0:
                    equipes[a["equipe"]] += a["receita"]

            # Ranking (top 10 assessores por receita)
            ranking = sorted(assessores, key=lambda x: x["receita"], reverse=True)[:10]

            # Categorias para barras
            categorias = [
                ("Renda Fixa", total_rf, ACCENT_GREEN),
                ("RV Corretagem", total_rv_corr, ACCENT_BLUE),
                ("Fundos Fechados", total_fundos, ACCENT_PURPLE),
                ("RV Estruturadas", total_rv_pe, ACCENT_ORANGE),
                ("Cambio", total_cambio, ACCENT_TEAL),
                ("COE", total_coe, ACCENT_RED),
            ]

            # Update UI on main thread
            def update_ui():
                self.rc_banner_text.configure(text=f"  Receita consolidada - {len(assessores)} assessores")
                self.rc_banner_date.configure(text=f"Base: {data_str}")

                self.rc_total_label.configure(text=fmt_currency(total_geral))
                self.rc_total_sub.configure(text=f"{n_assessores} assessores com receita ativa")

                self.rc_kpi_rf.configure(text=fmt_currency(total_rf))
                self.rc_kpi_rv_corr.configure(text=fmt_currency(total_rv_corr))
                self.rc_kpi_fundos.configure(text=fmt_currency(total_fundos))
                self.rc_kpi_rv_pe.configure(text=fmt_currency(total_rv_pe))
                self.rc_kpi_cambio.configure(text=fmt_currency(total_cambio))
                self.rc_kpi_coe.configure(text=fmt_currency(total_coe))
                self.rc_kpi_assessores.configure(text=str(n_assessores))

                # Barras de distribuição
                for w in self.rc_bars_inner.winfo_children():
                    w.destroy()

                max_val = max(v for _, v, _ in categorias) if categorias else 1
                for cat_name, cat_val, cat_color in categorias:
                    row_frame = ctk.CTkFrame(self.rc_bars_inner, fg_color="transparent")
                    row_frame.pack(fill="x", pady=3)

                    pct = (cat_val / total_geral * 100) if total_geral > 0 else 0

                    lbl = ctk.CTkLabel(
                        row_frame, text=cat_name, font=("Segoe UI", 11),
                        text_color=TEXT_PRIMARY, width=120, anchor="w"
                    )
                    lbl.pack(side="left")

                    bar_track = ctk.CTkFrame(row_frame, fg_color=BG_PROGRESS_TRACK, height=18, corner_radius=4)
                    bar_track.pack(side="left", fill="x", expand=True, padx=(8, 8))
                    bar_track.pack_propagate(False)

                    bar_width = max(cat_val / max_val, 0.02) if max_val > 0 else 0.02
                    bar_fill = ctk.CTkFrame(bar_track, fg_color=cat_color, corner_radius=4)
                    bar_fill.place(relx=0, rely=0, relwidth=bar_width, relheight=1.0)

                    val_lbl = ctk.CTkLabel(
                        row_frame, text=f"{fmt_currency(cat_val)}  ({pct:.1f}%)",
                        font=("Segoe UI", 10, "bold"), text_color=cat_color, width=180, anchor="e"
                    )
                    val_lbl.pack(side="right")

                # Ranking assessores
                self.rc_ranking_placeholder.destroy()
                for w in self.rc_ranking_inner.winfo_children():
                    w.destroy()

                # Header
                hdr = ctk.CTkFrame(self.rc_ranking_inner, fg_color="transparent")
                hdr.pack(fill="x", pady=(0, 8))
                ctk.CTkLabel(hdr, text="#", font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY, width=30, anchor="w").pack(side="left")
                ctk.CTkLabel(hdr, text="Assessor", font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY, width=200, anchor="w").pack(side="left")
                ctk.CTkLabel(hdr, text="Equipe", font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY, width=100, anchor="w").pack(side="left")
                ctk.CTkLabel(hdr, text="Receita Total", font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY, anchor="e").pack(side="right")

                sep = ctk.CTkFrame(self.rc_ranking_inner, fg_color=BORDER_LIGHT, height=1)
                sep.pack(fill="x", pady=(0, 4))

                medal_colors = [("#FFD700", "#B8860B"), ("#C0C0C0", "#808080"), ("#CD7F32", "#8B5A2B")]

                for idx, a in enumerate(ranking):
                    if a["receita"] <= 0:
                        continue
                    row = ctk.CTkFrame(self.rc_ranking_inner, fg_color="transparent")
                    row.pack(fill="x", pady=2)

                    pos = idx + 1
                    if idx < 3:
                        pos_color = medal_colors[idx][1]
                        pos_font = ("Segoe UI", 11, "bold")
                    else:
                        pos_color = TEXT_TERTIARY
                        pos_font = ("Segoe UI", 10)

                    ctk.CTkLabel(row, text=str(pos), font=pos_font, text_color=pos_color, width=30, anchor="w").pack(side="left")
                    ctk.CTkLabel(row, text=a["nome"], font=("Segoe UI", 11), text_color=TEXT_PRIMARY, width=200, anchor="w").pack(side="left")
                    ctk.CTkLabel(row, text=a["equipe"], font=("Segoe UI", 10), text_color=TEXT_SECONDARY, width=100, anchor="w").pack(side="left")

                    val_color = ACCENT_GREEN if idx < 3 else TEXT_PRIMARY
                    ctk.CTkLabel(row, text=fmt_currency(a["receita"]), font=("Segoe UI", 11, "bold"), text_color=val_color, anchor="e").pack(side="right")

                    if idx < len(ranking) - 1 and ranking[idx + 1]["receita"] > 0:
                        ctk.CTkFrame(self.rc_ranking_inner, fg_color=BORDER_LIGHT, height=1).pack(fill="x")

                # Equipes
                for w in self.rc_equipe_inner.winfo_children():
                    w.destroy()

                equipe_sorted = sorted(equipes.items(), key=lambda x: x[1], reverse=True)
                equipe_colors = [ACCENT_GREEN, ACCENT_BLUE, ACCENT_ORANGE, ACCENT_PURPLE, ACCENT_TEAL, ACCENT_RED, "#6b7280", "#9b59b6"]
                max_equipe = equipe_sorted[0][1] if equipe_sorted else 1

                for ei, (eq_name, eq_val) in enumerate(equipe_sorted):
                    row_frame = ctk.CTkFrame(self.rc_equipe_inner, fg_color="transparent")
                    row_frame.pack(fill="x", pady=4)

                    eq_color = equipe_colors[ei % len(equipe_colors)]
                    pct = (eq_val / total_geral * 100) if total_geral > 0 else 0

                    ctk.CTkLabel(
                        row_frame, text=eq_name, font=("Segoe UI", 11, "bold"),
                        text_color=TEXT_PRIMARY, width=120, anchor="w"
                    ).pack(side="left")

                    bar_track = ctk.CTkFrame(row_frame, fg_color=BG_PROGRESS_TRACK, height=22, corner_radius=5)
                    bar_track.pack(side="left", fill="x", expand=True, padx=(8, 8))
                    bar_track.pack_propagate(False)

                    bar_width = max(eq_val / max_equipe, 0.02) if max_equipe > 0 else 0.02
                    bar_fill = ctk.CTkFrame(bar_track, fg_color=eq_color, corner_radius=5)
                    bar_fill.place(relx=0, rely=0, relwidth=bar_width, relheight=1.0)

                    ctk.CTkLabel(
                        row_frame, text=f"{fmt_currency(eq_val)}  ({pct:.1f}%)",
                        font=("Segoe UI", 10, "bold"), text_color=eq_color, width=180, anchor="e"
                    ).pack(side="right")

            self.after(0, update_ui)

        except Exception as e:
            import traceback; traceback.print_exc()
            err_msg = f"  Erro ao carregar receita: {str(e)[:60]}"
            self.after(0, lambda m=err_msg: self.rc_banner_text.configure(text=m))

    # -----------------------------------------------------------------
    #  PAGE: CTRL RECEITA (BI Dashboard)
    # -----------------------------------------------------------------
    def _build_ctrl_receita_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        _, _, _ = self._make_topbar(page, "Controle Receita", subtitle="Dashboard BI")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Banner
        self.cr_banner = ctk.CTkFrame(content, fg_color=ACCENT_GREEN, corner_radius=10, height=44)
        self.cr_banner.pack(fill="x", pady=(0, 14))
        self.cr_banner.pack_propagate(False)

        self.cr_banner_text = ctk.CTkLabel(
            self.cr_banner, text="  Carregando dados...",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        )
        self.cr_banner_text.pack(side="left", padx=18, pady=10)

        self.cr_banner_date = ctk.CTkLabel(
            self.cr_banner, text="",
            font=("Segoe UI", 11), text_color="#80c0a0"
        )
        self.cr_banner_date.pack(side="right", padx=18)

        # Filter
        filter_frame = ctk.CTkFrame(content, fg_color="transparent")
        filter_frame.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            filter_frame, text="Assessor:",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY
        ).pack(side="left", padx=(0, 8))

        self.cr_assessor_var = ctk.StringVar(value="Todos")
        self.cr_assessor_menu = ctk.CTkOptionMenu(
            filter_frame, variable=self.cr_assessor_var,
            values=["Todos"], width=300,
            font=("Segoe UI", 12),
            fg_color=BG_INPUT, button_color=ACCENT_GREEN,
            button_hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY,
            command=self._on_ctrl_assessor_change,
        )
        self.cr_assessor_menu.pack(side="left")

        # === OVERVIEW FRAME ===
        self.cr_overview_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.cr_overview_frame.pack(fill="x")

        # RECEITA TOTAL card
        cr_total_card = ctk.CTkFrame(
            self.cr_overview_frame, fg_color=BG_CARD,
            corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        cr_total_card.pack(fill="x", pady=(0, 14))
        cr_total_inner = ctk.CTkFrame(cr_total_card, fg_color="transparent")
        cr_total_inner.pack(fill="x", padx=24, pady=16)

        ctk.CTkLabel(
            cr_total_inner, text="RECEITA TOTAL",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY
        ).pack(anchor="w")

        self.cr_total_label = ctk.CTkLabel(
            cr_total_inner, text="...",
            font=("Segoe UI", 34, "bold"), text_color=ACCENT_GREEN
        )
        self.cr_total_label.pack(anchor="w", pady=(2, 0))

        self.cr_total_sub = ctk.CTkLabel(
            cr_total_inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY
        )
        self.cr_total_sub.pack(anchor="w")

        # KPI Row 1
        cr_kpi1 = ctk.CTkFrame(self.cr_overview_frame, fg_color="transparent")
        cr_kpi1.pack(fill="x", pady=(0, 8))
        cr_kpi1.columnconfigure((0, 1, 2, 3), weight=1)

        self.cr_kpi_rf = self._make_kpi_card(cr_kpi1, "Renda Fixa", "...", ACCENT_GREEN, 0)
        self.cr_kpi_rv_corr = self._make_kpi_card(cr_kpi1, "RV Corretagem", "...", ACCENT_BLUE, 1)
        self.cr_kpi_fundos = self._make_kpi_card(cr_kpi1, "Fundos", "...", ACCENT_PURPLE, 2)
        self.cr_kpi_rv_pe = self._make_kpi_card(cr_kpi1, "RV Estruturadas", "...", ACCENT_ORANGE, 3)

        # KPI Row 2
        cr_kpi2 = ctk.CTkFrame(self.cr_overview_frame, fg_color="transparent")
        cr_kpi2.pack(fill="x", pady=(0, 14))
        cr_kpi2.columnconfigure((0, 1, 2), weight=1)

        self.cr_kpi_cambio = self._make_kpi_card(cr_kpi2, "Cambio", "...", ACCENT_TEAL, 0)
        self.cr_kpi_coe = self._make_kpi_card(cr_kpi2, "COE", "...", ACCENT_RED, 1)
        self.cr_kpi_n_assessores = self._make_kpi_card(cr_kpi2, "Assessores", "...", "#6b7280", 2)

        # Charts row (pie + bar)
        cr_charts = ctk.CTkFrame(self.cr_overview_frame, fg_color="transparent")
        cr_charts.pack(fill="x", pady=(0, 14))
        cr_charts.columnconfigure((0, 1), weight=1)

        pie_card, self.cr_ov_pie_fig, self.cr_ov_pie_ax, self.cr_ov_pie_canvas = \
            self._make_chart_card(cr_charts, "Distribuição por Produto", 4.5, 3)
        pie_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        bar_card, self.cr_ov_bar_fig, self.cr_ov_bar_ax, self.cr_ov_bar_canvas = \
            self._make_chart_card(cr_charts, "Top 10 Assessores", 4.5, 3)
        bar_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # Distribuição por Produto (barras)
        ctk.CTkLabel(
            self.cr_overview_frame, text="Distribuição por Produto",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        self.cr_ov_bars_card = ctk.CTkFrame(
            self.cr_overview_frame, fg_color=BG_CARD,
            corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        self.cr_ov_bars_card.pack(fill="x", pady=(0, 14))
        self.cr_ov_bars_inner = ctk.CTkFrame(self.cr_ov_bars_card, fg_color="transparent")
        self.cr_ov_bars_inner.pack(fill="x", padx=20, pady=16)

        # Receita por Equipe
        ctk.CTkLabel(
            self.cr_overview_frame, text="Receita por Equipe",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        self.cr_ov_equipe_card = ctk.CTkFrame(
            self.cr_overview_frame, fg_color=BG_CARD,
            corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        self.cr_ov_equipe_card.pack(fill="x", pady=(0, 14))
        self.cr_ov_equipe_inner = ctk.CTkFrame(self.cr_ov_equipe_card, fg_color="transparent")
        self.cr_ov_equipe_inner.pack(fill="x", padx=20, pady=16)

        # === DEEPDIVE FRAME (hidden initially) ===
        self.cr_deepdive_frame = ctk.CTkFrame(content, fg_color="transparent")

        # Header card
        self.cr_dd_header = ctk.CTkFrame(
            self.cr_deepdive_frame, fg_color=BG_CARD,
            corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        self.cr_dd_header.pack(fill="x", pady=(0, 14))

        dd_hdr_inner = ctk.CTkFrame(self.cr_dd_header, fg_color="transparent")
        dd_hdr_inner.pack(fill="x", padx=24, pady=16)

        self.cr_dd_nome = ctk.CTkLabel(
            dd_hdr_inner, text="",
            font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY
        )
        self.cr_dd_nome.pack(anchor="w")

        self.cr_dd_info = ctk.CTkLabel(
            dd_hdr_inner, text="",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.cr_dd_info.pack(anchor="w", pady=(2, 0))

        self.cr_dd_total = ctk.CTkLabel(
            dd_hdr_inner, text="",
            font=("Segoe UI", 28, "bold"), text_color=ACCENT_GREEN
        )
        self.cr_dd_total.pack(anchor="w", pady=(8, 0))

        self.cr_dd_rank = ctk.CTkLabel(
            dd_hdr_inner, text="",
            font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.cr_dd_rank.pack(anchor="w")

        # KPI Row 1 (deepdive)
        dd_kpi1 = ctk.CTkFrame(self.cr_deepdive_frame, fg_color="transparent")
        dd_kpi1.pack(fill="x", pady=(0, 8))
        dd_kpi1.columnconfigure((0, 1, 2, 3), weight=1)

        self.cr_dd_kpi_rf = self._make_kpi_card(dd_kpi1, "Renda Fixa", "...", ACCENT_GREEN, 0)
        self.cr_dd_kpi_rv_corr = self._make_kpi_card(dd_kpi1, "RV Corretagem", "...", ACCENT_BLUE, 1)
        self.cr_dd_kpi_fundos = self._make_kpi_card(dd_kpi1, "Fundos", "...", ACCENT_PURPLE, 2)
        self.cr_dd_kpi_rv_pe = self._make_kpi_card(dd_kpi1, "RV Estruturadas", "...", ACCENT_ORANGE, 3)

        # KPI Row 2 (deepdive)
        dd_kpi2 = ctk.CTkFrame(self.cr_deepdive_frame, fg_color="transparent")
        dd_kpi2.pack(fill="x", pady=(0, 14))
        dd_kpi2.columnconfigure((0, 1, 2, 3), weight=1)

        self.cr_dd_kpi_cambio = self._make_kpi_card(dd_kpi2, "Cambio", "...", ACCENT_TEAL, 0)
        self.cr_dd_kpi_coe = self._make_kpi_card(dd_kpi2, "COE", "...", ACCENT_RED, 1)
        self.cr_dd_kpi_pct = self._make_kpi_card(dd_kpi2, "% do Total", "...", "#6b7280", 2)
        self.cr_dd_kpi_ranking = self._make_kpi_card(dd_kpi2, "Ranking", "...", "#6b7280", 3)

        # Charts (deepdive)
        dd_charts = ctk.CTkFrame(self.cr_deepdive_frame, fg_color="transparent")
        dd_charts.pack(fill="x", pady=(0, 14))
        dd_charts.columnconfigure((0, 1), weight=1)

        dd_pie_card, self.cr_dd_pie_fig, self.cr_dd_pie_ax, self.cr_dd_pie_canvas = \
            self._make_chart_card(dd_charts, "Mix Receita", 4.5, 3)
        dd_pie_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        dd_bar_card, self.cr_dd_bar_fig, self.cr_dd_bar_ax, self.cr_dd_bar_canvas = \
            self._make_chart_card(dd_charts, "Receita por Produto", 4.5, 3)
        dd_bar_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # Detail table
        ctk.CTkLabel(
            self.cr_deepdive_frame, text="Operações Detalhadas",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        self.cr_dd_table_card = ctk.CTkFrame(
            self.cr_deepdive_frame, fg_color=BG_CARD,
            corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        self.cr_dd_table_card.pack(fill="x", pady=(0, 14))

        self.cr_dd_table_inner = ctk.CTkFrame(self.cr_dd_table_card, fg_color="transparent")
        self.cr_dd_table_inner.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            self.cr_dd_table_inner, text="Selecione um assessor para ver operações",
            font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        ).pack(anchor="w")

        # Botao Atualizar
        self._make_atualizar_btn(content, RECEITA_FILE)

        # Data cache
        self._cr_data = None
        self._cr_detail_cache = {}
        self._cr_loading = False

        return page

    # -- Chart helpers --
    def _make_chart_card(self, parent, title, fig_width=4, fig_height=3):
        card = ctk.CTkFrame(
            parent, fg_color=BG_CARD, corner_radius=12,
            border_width=1, border_color=BORDER_CARD
        )
        ctk.CTkLabel(
            card, text=title, font=("Segoe UI", 12, "bold"),
            text_color=TEXT_PRIMARY
        ).pack(anchor="w", padx=16, pady=(12, 4))

        fig = Figure(figsize=(fig_width, fig_height), dpi=90, facecolor='white')
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=card)
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=(0, 8))

        return card, fig, ax, canvas

    def _update_pie_chart(self, fig, ax, canvas, labels, values, colors):
        ax.clear()
        if not values or sum(values) == 0:
            ax.text(0.5, 0.5, 'Sem dados', ha='center', va='center',
                    fontsize=12, color='#9ca3af')
            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)
            ax.axis('off')
            canvas.draw()
            return
        wedges, texts, autotexts = ax.pie(
            values, labels=None, colors=colors,
            autopct='%1.1f%%', startangle=90,
            pctdistance=0.78, wedgeprops=dict(width=0.42)
        )
        for t in autotexts:
            t.set_fontsize(7)
            t.set_color('#333333')
        ax.legend(labels, loc='center left', bbox_to_anchor=(-0.15, 0.5),
                  fontsize=7, frameon=False)
        fig.tight_layout()
        canvas.draw()

    def _update_bar_chart(self, fig, ax, canvas, names, values, colors):
        ax.clear()
        if not values:
            ax.text(0.5, 0.5, 'Sem dados', ha='center', va='center',
                    fontsize=12, color='#9ca3af')
            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)
            ax.axis('off')
            canvas.draw()
            return
        y_pos = range(len(names))
        bars = ax.barh(y_pos, values, color=colors, height=0.6)
        ax.set_yticks(list(y_pos))
        ax.set_yticklabels(names, fontsize=7)
        ax.invert_yaxis()
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.tick_params(axis='x', labelsize=7)
        max_v = max(values) if values else 1
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + max_v * 0.02,
                    bar.get_y() + bar.get_height() / 2,
                    fmt_currency(val), va='center', fontsize=6)
        fig.tight_layout()
        canvas.draw()

    # -- Ctrl Receita data loading --
    def _load_ctrl_receita_data(self):
        try:
            if not os.path.exists(RECEITA_FILE):
                self.after(0, lambda: self.cr_banner_text.configure(
                    text="  Arquivo CONTROLE RECEITA TOTAL.xlsx não encontrado"))
                return

            wb = openpyxl.load_workbook(RECEITA_FILE, data_only=True)
            ws = wb["CONSOLIDADO"]

            data_base = ws.cell(3, 2).value
            if data_base and hasattr(data_base, "strftime"):
                data_str = data_base.strftime("%d/%m/%Y")
            else:
                data_str = str(data_base) if data_base else "N/A"

            assessores = []
            for r in range(7, ws.max_row + 1):
                cod = ws.cell(r, 2).value
                if not cod:
                    continue
                nome = sanitize_text(ws.cell(r, 4).value)
                equipe = sanitize_text(ws.cell(r, 3).value)
                rf = float(ws.cell(r, 5).value or 0)
                rv_corr = float(ws.cell(r, 6).value or 0)
                rv_pe = float(ws.cell(r, 7).value or 0)
                cambio = float(ws.cell(r, 8).value or 0)
                fundos = float(ws.cell(r, 9).value or 0)
                coe = float(ws.cell(r, 10).value or 0)
                receita = float(ws.cell(r, 11).value or 0)
                assessores.append({
                    "cod": str(cod), "nome": nome, "equipe": equipe,
                    "rf": rf, "rv_corr": rv_corr, "rv_pe": rv_pe,
                    "cambio": cambio, "fundos": fundos, "coe": coe,
                    "receita": receita,
                })
            wb.close()

            total_geral = sum(a["receita"] for a in assessores)
            ranking = sorted(assessores, key=lambda x: x["receita"], reverse=True)

            self._cr_data = {
                "assessores": assessores,
                "data_str": data_str,
                "total_geral": total_geral,
                "ranking": ranking,
            }

            names = ["Todos"] + [f'{a["cod"]} - {a["nome"]}' for a in ranking]

            def update_ui():
                self.cr_assessor_menu.configure(values=names)
                self.cr_banner_text.configure(
                    text=f"  Receita consolidada - {len(assessores)} assessores")
                self.cr_banner_date.configure(text=f"Base: {data_str}")
                self._update_ctrl_overview()

            self.after(0, update_ui)

        except Exception as e:
            import traceback; traceback.print_exc()
            err_msg = f"  Erro: {str(e)[:60]}"
            self.after(0, lambda m=err_msg: self.cr_banner_text.configure(text=m))

    def _on_ctrl_assessor_change(self, value):
        if value == "Todos":
            self.cr_deepdive_frame.pack_forget()
            self.cr_overview_frame.pack(fill="x")
            self._update_ctrl_overview()
        else:
            self.cr_overview_frame.pack_forget()
            self.cr_deepdive_frame.pack(fill="x")
            cod = value.split(" - ")[0].strip()
            self._update_ctrl_deepdive(cod)

    def _update_ctrl_overview(self):
        if not self._cr_data:
            return

        data = self._cr_data
        assessores = data["assessores"]
        total_geral = data["total_geral"]
        ranking = data["ranking"]

        total_rf = sum(a["rf"] for a in assessores)
        total_rv_corr = sum(a["rv_corr"] for a in assessores)
        total_rv_pe = sum(a["rv_pe"] for a in assessores)
        total_cambio = sum(a["cambio"] for a in assessores)
        total_fundos = sum(a["fundos"] for a in assessores)
        total_coe = sum(a["coe"] for a in assessores)
        n_assessores = len([a for a in assessores if a["receita"] > 0])

        self.cr_total_label.configure(text=fmt_currency(total_geral))
        self.cr_total_sub.configure(text=f"{n_assessores} assessores com receita ativa")

        self.cr_kpi_rf.configure(text=fmt_currency(total_rf))
        self.cr_kpi_rv_corr.configure(text=fmt_currency(total_rv_corr))
        self.cr_kpi_fundos.configure(text=fmt_currency(total_fundos))
        self.cr_kpi_rv_pe.configure(text=fmt_currency(total_rv_pe))
        self.cr_kpi_cambio.configure(text=fmt_currency(total_cambio))
        self.cr_kpi_coe.configure(text=fmt_currency(total_coe))
        self.cr_kpi_n_assessores.configure(text=str(n_assessores))

        # Pie chart
        cat_labels = ["RF", "RV Corr", "RV Estrut", "Fundos", "Cambio", "COE"]
        cat_values = [total_rf, total_rv_corr, total_rv_pe, total_fundos,
                      total_cambio, total_coe]
        cat_colors = [ACCENT_GREEN, ACCENT_BLUE, ACCENT_ORANGE, ACCENT_PURPLE,
                      ACCENT_TEAL, ACCENT_RED]

        pie_data = [(l, v, c) for l, v, c in zip(cat_labels, cat_values, cat_colors) if v > 0]
        if pie_data:
            self._update_pie_chart(
                self.cr_ov_pie_fig, self.cr_ov_pie_ax, self.cr_ov_pie_canvas,
                [d[0] for d in pie_data], [d[1] for d in pie_data],
                [d[2] for d in pie_data]
            )

        # Bar chart - top 10
        top10 = [a for a in ranking[:10] if a["receita"] > 0]
        top10_names = [a["nome"][:20] for a in top10]
        top10_values = [a["receita"] for a in top10]
        top10_colors = [ACCENT_GREEN] * len(top10_names)

        self._update_bar_chart(
            self.cr_ov_bar_fig, self.cr_ov_bar_ax, self.cr_ov_bar_canvas,
            top10_names, top10_values, top10_colors
        )

        # Barras distribuição produto
        categorias = [
            ("Renda Fixa", total_rf, ACCENT_GREEN),
            ("RV Corretagem", total_rv_corr, ACCENT_BLUE),
            ("Fundos", total_fundos, ACCENT_PURPLE),
            ("RV Estruturadas", total_rv_pe, ACCENT_ORANGE),
            ("Cambio", total_cambio, ACCENT_TEAL),
            ("COE", total_coe, ACCENT_RED),
        ]

        for w in self.cr_ov_bars_inner.winfo_children():
            w.destroy()

        max_val = max(v for _, v, _ in categorias) if categorias else 1
        for cat_name, cat_val, cat_color in categorias:
            row_frame = ctk.CTkFrame(self.cr_ov_bars_inner, fg_color="transparent")
            row_frame.pack(fill="x", pady=3)

            pct = (cat_val / total_geral * 100) if total_geral > 0 else 0

            ctk.CTkLabel(
                row_frame, text=cat_name, font=("Segoe UI", 11),
                text_color=TEXT_PRIMARY, width=120, anchor="w"
            ).pack(side="left")

            bar_track = ctk.CTkFrame(
                row_frame, fg_color=BG_PROGRESS_TRACK, height=18, corner_radius=4
            )
            bar_track.pack(side="left", fill="x", expand=True, padx=(8, 8))
            bar_track.pack_propagate(False)

            bar_w = max(cat_val / max_val, 0.02) if max_val > 0 else 0.02
            bar_fill = ctk.CTkFrame(bar_track, fg_color=cat_color, corner_radius=4)
            bar_fill.place(relx=0, rely=0, relwidth=bar_w, relheight=1.0)

            ctk.CTkLabel(
                row_frame, text=f"{fmt_currency(cat_val)}  ({pct:.1f}%)",
                font=("Segoe UI", 10, "bold"), text_color=cat_color,
                width=180, anchor="e"
            ).pack(side="right")

        # Equipes
        equipes = defaultdict(float)
        for a in assessores:
            if a["receita"] > 0:
                equipes[a["equipe"]] += a["receita"]

        for w in self.cr_ov_equipe_inner.winfo_children():
            w.destroy()

        equipe_sorted = sorted(equipes.items(), key=lambda x: x[1], reverse=True)
        equipe_colors = [ACCENT_GREEN, ACCENT_BLUE, ACCENT_ORANGE, ACCENT_PURPLE,
                         ACCENT_TEAL, ACCENT_RED, "#6b7280"]
        max_equipe = equipe_sorted[0][1] if equipe_sorted else 1

        for ei, (eq_name, eq_val) in enumerate(equipe_sorted):
            row_frame = ctk.CTkFrame(self.cr_ov_equipe_inner, fg_color="transparent")
            row_frame.pack(fill="x", pady=4)

            eq_color = equipe_colors[ei % len(equipe_colors)]
            pct = (eq_val / total_geral * 100) if total_geral > 0 else 0

            ctk.CTkLabel(
                row_frame, text=eq_name, font=("Segoe UI", 11, "bold"),
                text_color=TEXT_PRIMARY, width=120, anchor="w"
            ).pack(side="left")

            bar_track = ctk.CTkFrame(
                row_frame, fg_color=BG_PROGRESS_TRACK, height=22, corner_radius=5
            )
            bar_track.pack(side="left", fill="x", expand=True, padx=(8, 8))
            bar_track.pack_propagate(False)

            bar_w = max(eq_val / max_equipe, 0.02) if max_equipe > 0 else 0.02
            bar_fill = ctk.CTkFrame(bar_track, fg_color=eq_color, corner_radius=5)
            bar_fill.place(relx=0, rely=0, relwidth=bar_w, relheight=1.0)

            ctk.CTkLabel(
                row_frame, text=f"{fmt_currency(eq_val)}  ({pct:.1f}%)",
                font=("Segoe UI", 10, "bold"), text_color=eq_color,
                width=180, anchor="e"
            ).pack(side="right")

    def _update_ctrl_deepdive(self, cod):
        if not self._cr_data:
            return

        data = self._cr_data
        ranking = data["ranking"]
        total_geral = data["total_geral"]

        assessor = None
        rank_pos = 0
        for i, a in enumerate(ranking):
            if a["cod"] == cod:
                assessor = a
                rank_pos = i + 1
                break
        if not assessor:
            return

        # Header
        self.cr_dd_nome.configure(text=assessor["nome"])
        self.cr_dd_info.configure(
            text=f'Equipe: {assessor["equipe"]}  |  Código: {assessor["cod"]}'
        )
        self.cr_dd_total.configure(text=fmt_currency(assessor["receita"]))
        self.cr_dd_rank.configure(text=f'Ranking #{rank_pos} de {len(ranking)}')

        # KPIs
        self.cr_dd_kpi_rf.configure(text=fmt_currency(assessor["rf"]))
        self.cr_dd_kpi_rv_corr.configure(text=fmt_currency(assessor["rv_corr"]))
        self.cr_dd_kpi_fundos.configure(text=fmt_currency(assessor["fundos"]))
        self.cr_dd_kpi_rv_pe.configure(text=fmt_currency(assessor["rv_pe"]))
        self.cr_dd_kpi_cambio.configure(text=fmt_currency(assessor["cambio"]))
        self.cr_dd_kpi_coe.configure(text=fmt_currency(assessor["coe"]))

        pct = (assessor["receita"] / total_geral * 100) if total_geral > 0 else 0
        self.cr_dd_kpi_pct.configure(text=f"{pct:.1f}%")
        self.cr_dd_kpi_ranking.configure(text=f"#{rank_pos}")

        # Pie chart
        labels = ["RF", "RV Corr", "RV Estrut", "Fundos", "Cambio", "COE"]
        values = [assessor["rf"], assessor["rv_corr"], assessor["rv_pe"],
                  assessor["fundos"], assessor["cambio"], assessor["coe"]]
        colors = [ACCENT_GREEN, ACCENT_BLUE, ACCENT_ORANGE,
                  ACCENT_PURPLE, ACCENT_TEAL, ACCENT_RED]

        pie_data = [(l, v, c) for l, v, c in zip(labels, values, colors) if v > 0]
        if pie_data:
            self._update_pie_chart(
                self.cr_dd_pie_fig, self.cr_dd_pie_ax, self.cr_dd_pie_canvas,
                [d[0] for d in pie_data], [d[1] for d in pie_data],
                [d[2] for d in pie_data]
            )
        else:
            self.cr_dd_pie_ax.clear()
            self.cr_dd_pie_ax.text(
                0.5, 0.5, 'Sem dados', ha='center', va='center',
                fontsize=12, color='#9ca3af'
            )
            self.cr_dd_pie_ax.axis('off')
            self.cr_dd_pie_canvas.draw()

        # Bar chart
        bar_data = [(l, v, c) for l, v, c in zip(labels, values, colors) if v > 0]
        if bar_data:
            self._update_bar_chart(
                self.cr_dd_bar_fig, self.cr_dd_bar_ax, self.cr_dd_bar_canvas,
                [d[0] for d in bar_data], [d[1] for d in bar_data],
                [d[2] for d in bar_data]
            )
        else:
            self.cr_dd_bar_ax.clear()
            self.cr_dd_bar_ax.text(
                0.5, 0.5, 'Sem dados', ha='center', va='center',
                fontsize=12, color='#9ca3af'
            )
            self.cr_dd_bar_ax.axis('off')
            self.cr_dd_bar_canvas.draw()

        # Load detail table async
        threading.Thread(
            target=self._load_ctrl_assessor_details, args=(cod,), daemon=True
        ).start()

    def _load_ctrl_assessor_details(self, cod):
        if cod in self._cr_detail_cache:
            self.after(0, lambda: self._update_ctrl_detail_table(cod))
            return

        try:
            self.after(0, lambda: self._set_ctrl_detail_loading("Carregando operações..."))

            wb = openpyxl.load_workbook(RECEITA_FILE, data_only=True)
            details = []

            for sheet_info in CTRL_DETAIL_SHEETS:
                sheet_name = sheet_info["sheet"]
                if sheet_name not in wb.sheetnames:
                    continue

                ws = wb[sheet_name]
                col_a = sheet_info["col_assessor"]
                col_r = sheet_info["col_receita"]

                for r in range(2, ws.max_row + 1):
                    cell_val = ws.cell(r, col_a).value
                    if not cell_val:
                        continue
                    cell_str = str(cell_val).strip()
                    if cell_str != cod and cod not in cell_str:
                        continue

                    # Grab first columns for description
                    desc_parts = []
                    for c in range(1, min(ws.max_column + 1, 8)):
                        v = ws.cell(r, c).value
                        if v is not None and c != col_a:
                            desc_parts.append(sanitize_text(str(v))[:30])

                    receita_val = None
                    if col_r:
                        try:
                            receita_val = float(ws.cell(r, col_r).value or 0)
                        except (ValueError, TypeError):
                            receita_val = 0

                    details.append({
                        "produto": sheet_info["produto"],
                        "desc": " | ".join(desc_parts[:3]),
                        "receita": receita_val,
                    })

            wb.close()

            self._cr_detail_cache[cod] = details
            self.after(0, lambda: self._update_ctrl_detail_table(cod))

        except Exception as e:
            self.after(0, lambda: self._set_ctrl_detail_loading(
                f"Erro: {str(e)[:50]}"))

    def _set_ctrl_detail_loading(self, text):
        for w in self.cr_dd_table_inner.winfo_children():
            w.destroy()
        ctk.CTkLabel(
            self.cr_dd_table_inner, text=text,
            font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        ).pack(anchor="w")

    def _update_ctrl_detail_table(self, cod):
        details = self._cr_detail_cache.get(cod, [])

        for w in self.cr_dd_table_inner.winfo_children():
            w.destroy()

        if not details:
            ctk.CTkLabel(
                self.cr_dd_table_inner, text="Nenhuma operação encontrada",
                font=("Segoe UI", 11), text_color=TEXT_TERTIARY
            ).pack(anchor="w")
            return

        # Header
        hdr = ctk.CTkFrame(self.cr_dd_table_inner, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(
            hdr, text="Produto", font=("Segoe UI", 10, "bold"),
            text_color=TEXT_TERTIARY, width=120, anchor="w"
        ).pack(side="left")
        ctk.CTkLabel(
            hdr, text="Descrição", font=("Segoe UI", 10, "bold"),
            text_color=TEXT_TERTIARY, anchor="w"
        ).pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            hdr, text="Receita", font=("Segoe UI", 10, "bold"),
            text_color=TEXT_TERTIARY, width=120, anchor="e"
        ).pack(side="right")

        ctk.CTkFrame(
            self.cr_dd_table_inner, fg_color=BORDER_LIGHT, height=1
        ).pack(fill="x", pady=(0, 4))

        prod_colors = {
            "Sec RF": ACCENT_GREEN, "Prim RF": ACCENT_GREEN,
            "RV Corretagem": ACCENT_BLUE, "RV Estruturadas": ACCENT_ORANGE,
            "Cambio": ACCENT_TEAL, "Fundos Fechados": ACCENT_PURPLE,
            "Sec Fundos": ACCENT_PURPLE, "COE": ACCENT_RED,
        }

        for d in details[:50]:
            row = ctk.CTkFrame(self.cr_dd_table_inner, fg_color="transparent")
            row.pack(fill="x", pady=1)

            color = prod_colors.get(d["produto"], TEXT_PRIMARY)
            ctk.CTkLabel(
                row, text=d["produto"], font=("Segoe UI", 10),
                text_color=color, width=120, anchor="w"
            ).pack(side="left")
            ctk.CTkLabel(
                row, text=d["desc"][:60], font=("Segoe UI", 9),
                text_color=TEXT_SECONDARY, anchor="w"
            ).pack(side="left", fill="x", expand=True)

            rec_text = fmt_currency(d["receita"]) if d["receita"] is not None else "-"
            ctk.CTkLabel(
                row, text=rec_text, font=("Segoe UI", 10, "bold"),
                text_color=TEXT_PRIMARY, width=120, anchor="e"
            ).pack(side="right")

        if len(details) > 50:
            ctk.CTkLabel(
                self.cr_dd_table_inner,
                text=f"... e mais {len(details) - 50} operações",
                font=("Segoe UI", 10), text_color=TEXT_TERTIARY
            ).pack(anchor="w", pady=(4, 0))

    # -----------------------------------------------------------------
    #  PAGE: FLUXO - RF
    # -----------------------------------------------------------------
    def _build_dashboard_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        # Top bar profissional
        _, _, self.topbar_date = self._make_topbar(
            page, "FLUXO - RF", subtitle="Mesa de Produtos"
        )

        # Scroll area
        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Periodo banner
        self.período_banner = ctk.CTkFrame(content, fg_color=ACCENT_GREEN, corner_radius=10, height=44)
        self.período_banner.pack(fill="x", pady=(0, 18))
        self.período_banner.pack_propagate(False)

        self.período_text = ctk.CTkLabel(
            self.período_banner,
            text="  Carregando dados...",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        )
        self.período_text.pack(side="left", padx=18, pady=10)

        self.período_pdfs = ctk.CTkLabel(
            self.período_banner, text="",
            font=("Segoe UI", 11), text_color="#80c0a0"
        )
        self.período_pdfs.pack(side="right", padx=18)

        # KPI Cards row
        kpi_frame = ctk.CTkFrame(content, fg_color="transparent")
        kpi_frame.pack(fill="x", pady=(0, 18))
        kpi_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.kpi_assessores = self._make_kpi_card(kpi_frame, "Assessores", "...", ACCENT_GREEN, 0)
        self.kpi_eventos = self._make_kpi_card(kpi_frame, "Eventos", "...", ACCENT_BLUE, 1)
        self.kpi_ativos = self._make_kpi_card(kpi_frame, "Ativos Distintos", "...", ACCENT_ORANGE, 2)
        self.kpi_datas = self._make_kpi_card(kpi_frame, "Datas com Eventos", "...", ACCENT_PURPLE, 3)

        # Total destaque
        total_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        total_card.pack(fill="x", pady=(0, 18))

        total_inner = ctk.CTkFrame(total_card, fg_color="transparent")
        total_inner.pack(fill="x", padx=24, pady=16)

        ctk.CTkLabel(
            total_inner, text="VALOR TOTAL ESTIMADO",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_TERTIARY
        ).pack(anchor="w")

        self.total_value_label = ctk.CTkLabel(
            total_inner, text="Carregando...",
            font=("Segoe UI", 32, "bold"), text_color=ACCENT_GREEN
        )
        self.total_value_label.pack(anchor="w", pady=(2, 0))

        self.total_sub_label = ctk.CTkLabel(
            total_inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY
        )
        self.total_sub_label.pack(anchor="w")

        # Breakdown por tipo
        tipo_label = ctk.CTkLabel(
            content, text="Distribuição por Tipo de Evento",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        tipo_label.pack(fill="x", pady=(0, 10))

        self.tipo_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.tipo_frame.pack(fill="x", pady=(0, 18))
        self.tipo_frame.columnconfigure((0, 1), weight=1)

        # Placeholder - será preenchido quando os dados carregarem
        self.tipo_cards = {}

        # Ações rapidas
        act_label = ctk.CTkLabel(
            content, text="Ações Rápidas",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        act_label.pack(fill="x", pady=(0, 10))

        act_frame = ctk.CTkFrame(content, fg_color="transparent")
        act_frame.pack(fill="x", pady=(0, 10))
        act_frame.columnconfigure((0, 1, 2), weight=1)

        self._make_action_card(
            act_frame, "\u2913", "Gerar PDFs",
            "Gera um relatório PDF para cada assessor com todos os fluxos de renda fixa detalhados por data.",
            ACCENT_GREEN, 0, self._go_gerar_pdfs
        )
        self._make_action_card(
            act_frame, "\u2709", "Gerar Emails",
            "Cria rascunhos no Outlook com o PDF anexado, prontos para revisão e envio.",
            ACCENT_BLUE, 1, self._go_gerar_emails
        )
        self._make_action_card(
            act_frame, "\u2750", "Abrir Pasta",
            "Abre a pasta de PDFs gerados no Windows Explorer.",
            "#6b7280", 2, self._on_abrir_pasta
        )

        # Botao Atualizar
        self._make_atualizar_btn(content, ENTRADA_FILE)

        return page

    def _make_kpi_card(self, parent, label, value, color, col):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        card.grid(row=0, column=col, padx=(0 if col == 0 else 6, 0 if col == 3 else 6), sticky="nsew")

        # Indicador de cor
        bar = ctk.CTkFrame(card, fg_color=color, height=4, corner_radius=2)
        bar.pack(fill="x", padx=14, pady=(14, 0))

        ctk.CTkLabel(
            card, text=label, font=("Segoe UI", 10),
            text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=16, pady=(10, 0))

        val_label = ctk.CTkLabel(
            card, text=value, font=("Segoe UI", 26, "bold"),
            text_color=TEXT_PRIMARY
        )
        val_label.pack(anchor="w", padx=16, pady=(0, 14))

        return val_label

    def _make_action_card(self, parent, icon, title, desc, color, col, command):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        card.grid(row=0, column=col, padx=(0 if col == 0 else 5, 0 if col == 2 else 5), sticky="nsew")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=18, pady=16)

        header = ctk.CTkFrame(inner, fg_color="transparent")
        header.pack(fill="x")

        ctk.CTkLabel(
            header, text=icon, font=("Segoe UI", 22),
            text_color=color
        ).pack(side="left")

        ctk.CTkLabel(
            header, text=title, font=("Segoe UI", 14, "bold"),
            text_color=TEXT_PRIMARY
        ).pack(side="left", padx=(10, 0))

        ctk.CTkLabel(
            inner, text=desc, font=("Segoe UI", 10),
            text_color=TEXT_SECONDARY, wraplength=250, justify="left", anchor="w"
        ).pack(fill="x", pady=(8, 12))

        ctk.CTkButton(
            inner, text=f"Executar", font=("Segoe UI", 11, "bold"),
            fg_color=color, hover_color=self._darken(color),
            height=36, corner_radius=8, command=command
        ).pack(anchor="w")

    def _make_tipo_card(self, parent, tipo_name, count, total, color, row, col):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=BORDER_CARD)
        card.grid(row=row, column=col, padx=(0 if col == 0 else 4, 0 if col == 1 else 4), pady=4, sticky="nsew")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=12)

        # Dot + name
        header = ctk.CTkFrame(inner, fg_color="transparent")
        header.pack(fill="x")

        ctk.CTkLabel(header, text="\u25cf", font=("Segoe UI", 10), text_color=color).pack(side="left")
        ctk.CTkLabel(
            header, text=f"  {tipo_name}", font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY
        ).pack(side="left")
        ctk.CTkLabel(
            header, text=f"{count} eventos", font=("Segoe UI", 10), text_color=TEXT_TERTIARY
        ).pack(side="right")

        # Value
        ctk.CTkLabel(
            inner, text=fmt_currency(total),
            font=("Segoe UI", 16, "bold"), text_color=color, anchor="w"
        ).pack(fill="x", pady=(4, 0))

    def _make_atualizar_btn(self, parent, file_path):
        """Botao verde 'Atualizar' que abre a pasta da planilha fonte."""
        folder = os.path.dirname(os.path.abspath(file_path))
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(24, 8))
        ctk.CTkButton(
            btn_frame, text="\u21BB  Atualizar",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_WHITE, height=42, corner_radius=10,
            command=lambda: os.startfile(folder)
        ).pack(fill="x")

    # -----------------------------------------------------------------
    #  PAGE: OPERATIONS
    # -----------------------------------------------------------------
    def _build_operations_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        # Top bar profissional
        _, self.ops_title, _ = self._make_topbar(
            page, "Operações", subtitle="Geração de PDFs e Emails",
            back_btn=True, back_cmd=lambda: self._show_page("fluxo_rf")
        )

        # Content
        content = ctk.CTkFrame(page, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=28, pady=20)

        # Status card
        status_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD)
        status_card.pack(fill="x", pady=(0, 14))

        status_inner = ctk.CTkFrame(status_card, fg_color="transparent")
        status_inner.pack(fill="x", padx=24, pady=18)

        self.ops_status_dot = ctk.CTkLabel(
            status_inner, text="\u25cf", font=("Segoe UI", 14), text_color=TEXT_TERTIARY
        )
        self.ops_status_dot.pack(side="left")

        self.ops_status_text = ctk.CTkLabel(
            status_inner, text="  Aguardando...",
            font=("Segoe UI", 13), text_color=TEXT_PRIMARY
        )
        self.ops_status_text.pack(side="left")

        self.ops_counter = ctk.CTkLabel(
            status_inner, text="",
            font=("Segoe UI", 12, "bold"), text_color=ACCENT_GREEN
        )
        self.ops_counter.pack(side="right")

        # Progress
        self.ops_progress = ctk.CTkProgressBar(
            content, fg_color=BG_PROGRESS_TRACK, progress_color=ACCENT_GREEN,
            height=8, corner_radius=4
        )
        self.ops_progress.pack(fill="x", pady=(0, 4))
        self.ops_progress.set(0)

        self.ops_progress_text = ctk.CTkLabel(
            content, text="", font=("Segoe UI", 9), text_color=TEXT_TERTIARY, anchor="w"
        )
        self.ops_progress_text.pack(fill="x", pady=(0, 12))

        # Log
        log_card = ctk.CTkFrame(content, fg_color=BG_LOG, corner_radius=12, border_width=1, border_color="#2a3040")
        log_card.pack(fill="both", expand=True)

        log_header = ctk.CTkFrame(log_card, fg_color=BG_LOG_HEADER, corner_radius=0)
        log_header.pack(fill="x")

        ctk.CTkLabel(
            log_header, text="  Terminal",
            font=("Consolas", 10, "bold"), text_color="#5a6a7a"
        ).pack(side="left", padx=12, pady=6)

        ctk.CTkButton(
            log_header, text="Limpar", font=("Segoe UI", 9),
            fg_color="transparent", hover_color="#2a3040",
            text_color="#5a6a7a", width=50, height=22,
            corner_radius=4, command=self._clear_ops_log
        ).pack(side="right", padx=8, pady=4)

        self.ops_log = ctk.CTkTextbox(
            log_card, font=("Consolas", 10), fg_color=BG_LOG,
            text_color=TEXT_LOG, corner_radius=0, wrap="word",
            border_width=0
        )
        self.ops_log.pack(fill="both", expand=True, padx=2, pady=(0, 2))
        self.ops_log.configure(state="disabled")

        return page

    # -----------------------------------------------------------------
    #  PAGE: INFORMATIVO - E-MAIL
    # -----------------------------------------------------------------
    def _build_informativo_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Informativo", subtitle="E-mail")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Banner
        banner = ctk.CTkFrame(content, fg_color=ACCENT_BLUE, corner_radius=10, height=44)
        banner.pack(fill="x", pady=(0, 18))
        banner.pack_propagate(False)

        ctk.CTkLabel(
            banner, text="  Compose e envie informativos por e-mail aos assessores",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        ).pack(side="left", padx=18, pady=10)

        # === ASSUNTO ===
        ctk.CTkLabel(
            content, text="Assunto",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.info_assunto = ctk.CTkEntry(
            content, font=("Segoe UI", 12), height=40,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: Informativo Semanal - Mesa de Produtos"
        )
        self.info_assunto.pack(fill="x", pady=(0, 14))

        # === CORPO DO EMAIL ===
        ctk.CTkLabel(
            content, text="Corpo do E-mail",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.info_corpo = ctk.CTkTextbox(
            content, font=("Segoe UI", 11), height=220,
            corner_radius=8, border_width=1, border_color=BORDER_CARD,
            fg_color=BG_CARD, text_color=TEXT_PRIMARY, wrap="word"
        )
        self.info_corpo.pack(fill="x", pady=(0, 14))

        # === CC - CÓPIA ===
        ctk.CTkLabel(
            content, text="CC - Cópia",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        cc_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=10,
                               border_width=1, border_color=BORDER_CARD)
        cc_card.pack(fill="x", pady=(0, 14))

        cc_inner = ctk.CTkFrame(cc_card, fg_color="transparent")
        cc_inner.pack(fill="x", padx=16, pady=12)

        # Backoffice - assistente da base por assessor (checkbox pré-marcado)
        self.info_cc_backoffice_var = ctk.IntVar(value=1)
        ctk.CTkCheckBox(
            cc_inner, text="Backoffice / Assistente (da planilha BASE EMAILS)",
            font=("Segoe UI", 11), text_color=TEXT_PRIMARY,
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            variable=self.info_cc_backoffice_var,
        ).pack(anchor="w", pady=(0, 8))

        # Separador
        ctk.CTkFrame(cc_inner, fg_color=BORDER_CARD, height=1).pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            cc_inner, text="Pessoas em cópia:",
            font=("Segoe UI", 10, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(anchor="w", pady=(0, 6))

        # Lista de pessoas selecionáveis
        self._info_cc_pessoas = [
            ("Leonardo Dellatorre", "Leonardo.dellatorre@somuscapital.com.br"),
            ("João Morais", "joao.morais@somuscapital.com.br"),
            ("Alexandre Achui", "alexandre.achui@somuscapital.com.br"),
            ("Alessandra Lima", "alessandra.lima@somuscapital.com.br"),
        ]
        self._info_cc_vars = []

        cc_grid = ctk.CTkFrame(cc_inner, fg_color="transparent")
        cc_grid.pack(fill="x", pady=(0, 8))

        for i, (nome, email_addr) in enumerate(self._info_cc_pessoas):
            var = ctk.IntVar(value=0)
            self._info_cc_vars.append(var)
            cb = ctk.CTkCheckBox(
                cc_grid, text=f"{nome}",
                font=("Segoe UI", 11), text_color=TEXT_PRIMARY,
                fg_color=ACCENT_BLUE, hover_color="#1555bb",
                variable=var,
            )
            row = i // 2
            col = i % 2
            cb.grid(row=row, column=col, sticky="w", padx=(0, 24), pady=2)

        cc_grid.columnconfigure(0, weight=1)
        cc_grid.columnconfigure(1, weight=1)

        # Separador
        ctk.CTkFrame(cc_inner, fg_color=BORDER_CARD, height=1).pack(fill="x", pady=(4, 8))

        # Campo manual para CC adicional
        ctk.CTkLabel(
            cc_inner, text="Outros (separar com ponto e vírgula):",
            font=("Segoe UI", 10), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(anchor="w", pady=(0, 4))

        self.info_cc_manual = ctk.CTkEntry(
            cc_inner, font=("Segoe UI", 11), height=36,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: nome@email.com; outro@email.com"
        )
        self.info_cc_manual.pack(fill="x")

        # === ANEXO ===
        ctk.CTkLabel(
            content, text="Anexo (opcional)",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        anexo_frame = ctk.CTkFrame(content, fg_color="transparent")
        anexo_frame.pack(fill="x", pady=(0, 14))

        self.info_anexo_label = ctk.CTkLabel(
            anexo_frame, text="Nenhum arquivo selecionado",
            font=("Segoe UI", 11), text_color=TEXT_TERTIARY, anchor="w"
        )
        self.info_anexo_label.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            anexo_frame, text="Selecionar",
            font=("Segoe UI", 11, "bold"),
            fg_color=BG_INPUT, hover_color=BORDER_CARD,
            text_color=TEXT_SECONDARY, height=34, corner_radius=8, width=110,
            command=self._on_info_browse_anexo,
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            anexo_frame, text="Limpar",
            font=("Segoe UI", 10),
            fg_color="transparent", hover_color=BORDER_CARD,
            text_color=TEXT_TERTIARY, height=34, corner_radius=8, width=70,
            command=self._on_info_clear_anexo,
        ).pack(side="right")

        # === DESTINATARIOS ===
        ctk.CTkLabel(
            content, text="Destinatarios",
            font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        dest_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=10,
                                 border_width=1, border_color=BORDER_CARD)
        dest_card.pack(fill="x", pady=(0, 14))

        dest_inner = ctk.CTkFrame(dest_card, fg_color="transparent")
        dest_inner.pack(fill="x", padx=16, pady=12)

        # Seletor de Time/Equipe
        team_row = ctk.CTkFrame(dest_inner, fg_color="transparent")
        team_row.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            team_row, text="Enviar para:",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(side="left", padx=(0, 10))

        # Carregar equipes disponiveis da base
        self._info_equipe_map = load_equipe_mapping()
        equipes_disponiveis = sorted(set(self._info_equipe_map.values()))
        opcoes_time = ["Todos"] + equipes_disponiveis

        self.info_team_var = ctk.StringVar(value="Todos")
        self.info_team_selector = ctk.CTkSegmentedButton(
            team_row,
            values=opcoes_time,
            variable=self.info_team_var,
            font=("Segoe UI", 11),
            selected_color=ACCENT_GREEN,
            selected_hover_color=BG_SIDEBAR_HOVER,
            unselected_color=BG_INPUT,
            unselected_hover_color=BORDER_CARD,
            text_color=TEXT_WHITE,
            corner_radius=8,
            command=self._on_info_team_changed,
        )
        self.info_team_selector.pack(side="left", fill="x", expand=True)

        self.info_dest_count = ctk.CTkLabel(
            dest_inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        )
        self.info_dest_count.pack(anchor="w", pady=(4, 0))

        # Lista de assessores do time selecionado
        self.info_team_list = ctk.CTkLabel(
            dest_inner, text="",
            font=("Segoe UI", 9), text_color=TEXT_SECONDARY, anchor="w",
            justify="left", wraplength=600
        )
        self.info_team_list.pack(anchor="w", pady=(2, 0))

        # Carregar contagem inicial
        self._info_assessores_base = {}
        try:
            self._info_assessores_base = load_base_emails()
            self._on_info_team_changed("Todos")
        except Exception:
            self.info_dest_count.configure(text="Erro ao carregar base")

        # === AÇÕES ===
        act_frame = ctk.CTkFrame(content, fg_color="transparent")
        act_frame.pack(fill="x", pady=(6, 14))

        self.info_enviar_btn = ctk.CTkButton(
            act_frame, text="\u2709  Gerar Rascunhos no Outlook",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=44, corner_radius=10,
            command=self._on_info_gerar_emails,
        )
        self.info_enviar_btn.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            act_frame, text="Limpar Tudo",
            font=("Segoe UI", 11),
            fg_color="#6b7280", hover_color="#555d6a",
            height=44, corner_radius=10, width=120,
            command=self._on_info_limpar,
        ).pack(side="left")

        # === STATUS ===
        self.info_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.info_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.info_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.info_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.info_status_dot.pack(side="left")

        self.info_status_text = ctk.CTkLabel(
            si, text="  Preencha os campos e clique em Gerar Rascunhos",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.info_status_text.pack(side="left")

        # Botao Atualizar
        self._make_atualizar_btn(content, BASE_FILE)

        # Internal state
        self._info_anexo_path = None

        return page

    # -----------------------------------------------------------------
    #  INFORMATIVO: Ações
    # -----------------------------------------------------------------
    def _on_info_team_changed(self, selected):
        """Atualiza contagem e lista de assessores quando o time muda."""
        assessores = self._info_assessores_base
        equipe_map = self._info_equipe_map
        filtered = self._get_info_filtered_assessores()
        n = len(filtered)
        if selected == "Todos":
            self.info_dest_count.configure(text=f"{n} assessores com e-mail cadastrado")
            self.info_team_list.configure(text="")
        else:
            nomes = [filtered[c]["nome"] for c in sorted(filtered.keys())]
            self.info_dest_count.configure(text=f"{n} assessores na equipe {selected}")
            self.info_team_list.configure(text=", ".join(nomes))

    def _get_info_filtered_assessores(self):
        """Retorna assessores filtrados pelo time selecionado."""
        assessores = self._info_assessores_base
        selected = self.info_team_var.get()
        filtered = {}
        for cod, info in assessores.items():
            email = info.get("email", "")
            if not email or email == "-":
                continue
            if selected == "Todos":
                filtered[cod] = info
            else:
                equipe = self._info_equipe_map.get(cod, "")
                if equipe == selected:
                    filtered[cod] = info
        return filtered

    def _on_info_browse_anexo(self):
        path = filedialog.askopenfilename(
            title="Selecionar anexo",
            filetypes=[("Todos", "*.*"), ("PDF", "*.pdf"), ("Excel", "*.xlsx"), ("Imagem", "*.png *.jpg")]
        )
        if path:
            self._info_anexo_path = path
            self.info_anexo_label.configure(
                text=os.path.basename(path), text_color=ACCENT_GREEN
            )

    def _on_info_clear_anexo(self):
        self._info_anexo_path = None
        self.info_anexo_label.configure(
            text="Nenhum arquivo selecionado", text_color=TEXT_TERTIARY
        )

    def _on_info_limpar(self):
        self.info_assunto.delete(0, "end")
        self.info_corpo.delete("1.0", "end")
        self.info_cc_manual.delete(0, "end")
        self.info_cc_backoffice_var.set(1)  # Backoffice sempre marcado por padrão
        for var in self._info_cc_vars:
            var.set(0)
        self._on_info_clear_anexo()
        self.info_team_var.set("Todos")
        self._on_info_team_changed("Todos")
        self.info_status_dot.configure(text_color=TEXT_TERTIARY)
        self.info_status_text.configure(text="  Preencha os campos e clique em Gerar Rascunhos")

    def _on_info_gerar_emails(self):
        assunto = self.info_assunto.get().strip()
        corpo = self.info_corpo.get("1.0", "end").strip()

        if not assunto:
            messagebox.showwarning("Campo obrigatorio", "Preencha o assunto do e-mail.")
            return
        if not corpo:
            messagebox.showwarning("Campo obrigatorio", "Preencha o corpo do e-mail.")
            return

        filtered = self._get_info_filtered_assessores()
        if not filtered:
            messagebox.showwarning("Sem destinatarios", "Nenhum assessor encontrado para o time selecionado.")
            return

        team_sel = self.info_team_var.get()
        team_label = f"equipe {team_sel}" if team_sel != "Todos" else "todos os assessores"

        resp = messagebox.askyesno(
            "Gerar Rascunhos",
            f"Sera criado um rascunho no Outlook para {len(filtered)} assessor(es) ({team_label}).\n\n"
            "Os e-mails NAO serao enviados, apenas salvos como rascunho.\n\n"
            "O Outlook precisa estar aberto. Deseja continuar?"
        )
        if not resp:
            return

        # Montar lista de CC extra (checkboxes + manual)
        include_backoffice = bool(self.info_cc_backoffice_var.get())
        cc_extras = []
        for i, var in enumerate(self._info_cc_vars):
            if var.get():
                cc_extras.append(self._info_cc_pessoas[i][1])
        cc_manual = self.info_cc_manual.get().strip()
        if cc_manual:
            cc_extras.append(cc_manual)
        cc_extra_str = "; ".join(cc_extras)

        self.info_enviar_btn.configure(state="disabled")
        self.info_status_dot.configure(text_color=ACCENT_ORANGE)
        self.info_status_text.configure(text=f"  Gerando rascunhos para {team_label}...")
        threading.Thread(target=self._run_info_emails, args=(assunto, corpo, filtered, cc_extra_str, include_backoffice), daemon=True).start()

    def _run_info_emails(self, assunto, corpo, filtered_assessores, cc_manual="", include_backoffice=True):
        try:
            import win32com.client as win32
            import pythoncom
            pythoncom.CoInitialize()

            # Recarregar base de emails para pegar alterações recentes
            try:
                self._info_assessores_base = load_base_emails()
            except Exception:
                pass

            outlook = win32.Dispatch("Outlook.Application")

            hoje = datetime.now().strftime("%d/%m/%Y")
            logo_tag = LOGO_TAG_CID if os.path.exists(LOGO_PATH) else ""

            # Converter quebras de linha em <br>
            corpo_html = corpo.replace("\n", "<br>")

            criados = 0
            erros = 0
            codes = sorted(filtered_assessores.keys())

            for código in codes:
                info = filtered_assessores[código]
                email = info.get("email", "")
                if not email or email == "-":
                    continue

                nome = info["nome"]
                primeiro_nome = nome.split()[0] if nome else "Assessor"

                html_body = f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:960px;margin:0 auto;">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding:16px 0 12px 0;vertical-align:middle;">
      {logo_tag}<!--
      --><span style="font-size:16pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:300;"> | </span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:bold;">Produtos</span>
    </td>
  </tr>
  <tr>
    <td style="padding:0 0 16px 0;">
      <hr style="border:none;border-top:2px solid #004d33;margin:0;">
    </td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px;">
  <tr><td style="padding:0 4px;">
    <p style="font-size:11pt;color:#1a1a2e;margin-bottom:4px;">
      Prezado(a) <b>{primeiro_nome}</b>, tudo bem?
    </p>
    <p style="font-size:10.5pt;color:#000000;margin-top:0;">
      {corpo_html}
    </p>
  </td></tr>
  <tr><td style="padding:28px 0 0 0;">
    <hr style="border:none;border-top:2px solid #004d33;margin:0 0 12px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="vertical-align:middle;">
          <span style="font-size:10pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:300;"> | </span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:bold;">Produtos</span>
        </td>
        <td style="text-align:right;">
          <span style="font-size:8.5pt;color:#6b7280;">Informativo &middot; {hoje}</span><br>
          <span style="font-size:7.5pt;color:#9ca3af;">E-mail gerado automaticamente</span>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</div>
"""
                try:
                    mail = outlook.CreateItem(0)
                    mail.To = email

                    # Montar CC: backoffice/assistente da base (se marcado) + CC extras
                    cc_parts = []
                    if include_backoffice:
                        email_assist = info.get("email_assistente", "-")
                        if email_assist and email_assist != "-":
                            cc_parts.append(email_assist)
                    if cc_manual:
                        cc_parts.append(cc_manual)
                    if cc_parts:
                        mail.CC = "; ".join(cc_parts)

                    mail.Subject = assunto
                    mail.HTMLBody = html_body
                    _attach_logo_cid(mail)
                    if self._info_anexo_path and os.path.exists(self._info_anexo_path):
                        mail.Attachments.Add(os.path.abspath(self._info_anexo_path))
                    mail.Save()
                    criados += 1
                except Exception:
                    erros += 1

            pythoncom.CoUninitialize()

            def _done():
                self.info_enviar_btn.configure(state="normal")
                self.info_status_dot.configure(text_color="#00a86b")
                self.info_status_text.configure(
                    text=f"  Concluído - {criados} rascunhos criados, {erros} erros"
                )
                messagebox.showinfo(
                    "Rascunhos Criados",
                    f"{criados} rascunhos criados no Outlook!\n{erros} erros."
                )
            self.after(0, _done)

        except Exception as e:
            def _err():
                self.info_enviar_btn.configure(state="normal")
                self.info_status_dot.configure(text_color=ACCENT_RED)
                self.info_status_text.configure(text=f"  Erro: {e}")
                messagebox.showerror("Erro", str(e))
            self.after(0, _err)

    # -----------------------------------------------------------------
    #  PAGE: INFORMATIVO - AGIO
    # -----------------------------------------------------------------
    def _build_info_agio_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Informativo - Agio", subtitle="Mesa de Produtos")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Banner
        banner = ctk.CTkFrame(content, fg_color=ACCENT_GREEN, corner_radius=10, height=44)
        banner.pack(fill="x", pady=(0, 18))
        banner.pack_propagate(False)

        ctk.CTkLabel(
            banner, text="  Informativo de Agio/Desagio - Dashboard e Envio",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        ).pack(side="left", padx=18, pady=10)

        ctk.CTkButton(
            banner, text="\u21BB  Recarregar",
            font=("Segoe UI", 10, "bold"),
            fg_color=BG_SIDEBAR_HOVER, hover_color=BG_SIDEBAR_ACTIVE,
            text_color=TEXT_WHITE, height=28, corner_radius=6, width=110,
            command=self._ag_load_data,
        ).pack(side="right", padx=18)

        # ======== KPI CARDS ========
        kpi_frame = ctk.CTkFrame(content, fg_color="transparent")
        kpi_frame.pack(fill="x", pady=(0, 8))
        kpi_frame.columnconfigure((0, 1, 2, 3), weight=1)

        # KPI 1: Total R$ em Agio
        kpi1 = ctk.CTkFrame(kpi_frame, fg_color=BG_CARD, corner_radius=12,
                             border_width=1, border_color=BORDER_CARD)
        kpi1.grid(row=0, column=0, padx=(0, 6), sticky="nsew")
        ctk.CTkFrame(kpi1, fg_color=ACCENT_GREEN, height=4,
                     corner_radius=2).pack(fill="x", padx=14, pady=(14, 0))
        ctk.CTkLabel(kpi1, text="Total Agio (R$)", font=("Segoe UI", 10),
                     text_color=TEXT_SECONDARY).pack(anchor="w", padx=16, pady=(10, 0))
        self.ag_kpi_total = ctk.CTkLabel(
            kpi1, text="...", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY)
        self.ag_kpi_total.pack(anchor="w", padx=16, pady=(0, 14))

        # KPI 2: Assessores
        kpi2 = ctk.CTkFrame(kpi_frame, fg_color=BG_CARD, corner_radius=12,
                             border_width=1, border_color=BORDER_CARD)
        kpi2.grid(row=0, column=1, padx=6, sticky="nsew")
        ctk.CTkFrame(kpi2, fg_color=ACCENT_BLUE, height=4,
                     corner_radius=2).pack(fill="x", padx=14, pady=(14, 0))
        ctk.CTkLabel(kpi2, text="Assessores", font=("Segoe UI", 10),
                     text_color=TEXT_SECONDARY).pack(anchor="w", padx=16, pady=(10, 0))
        self.ag_kpi_assessores = ctk.CTkLabel(
            kpi2, text="...", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY)
        self.ag_kpi_assessores.pack(anchor="w", padx=16, pady=(0, 14))

        # KPI 3: Papeis em Agio
        kpi3 = ctk.CTkFrame(kpi_frame, fg_color=BG_CARD, corner_radius=12,
                             border_width=1, border_color=BORDER_CARD)
        kpi3.grid(row=0, column=2, padx=6, sticky="nsew")
        ctk.CTkFrame(kpi3, fg_color=ACCENT_ORANGE, height=4,
                     corner_radius=2).pack(fill="x", padx=14, pady=(14, 0))
        ctk.CTkLabel(kpi3, text="Papeis em Agio", font=("Segoe UI", 10),
                     text_color=TEXT_SECONDARY).pack(anchor="w", padx=16, pady=(10, 0))
        self.ag_kpi_papeis = ctk.CTkLabel(
            kpi3, text="...", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY)
        self.ag_kpi_papeis.pack(anchor="w", padx=16, pady=(0, 14))

        # KPI 4: Principal Indexador
        kpi4 = ctk.CTkFrame(kpi_frame, fg_color=BG_CARD, corner_radius=12,
                             border_width=1, border_color=BORDER_CARD)
        kpi4.grid(row=0, column=3, padx=(6, 0), sticky="nsew")
        ctk.CTkFrame(kpi4, fg_color=ACCENT_PURPLE, height=4,
                     corner_radius=2).pack(fill="x", padx=14, pady=(14, 0))
        ctk.CTkLabel(kpi4, text="Principal Indexador", font=("Segoe UI", 10),
                     text_color=TEXT_SECONDARY).pack(anchor="w", padx=16, pady=(10, 0))
        self.ag_kpi_indexador = ctk.CTkLabel(
            kpi4, text="...", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY)
        self.ag_kpi_indexador.pack(anchor="w", padx=16, pady=(0, 14))

        # ======== TOTAL DESTAQUE ========
        total_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                  border_width=1, border_color=BORDER_CARD)
        total_card.pack(fill="x", pady=(0, 18))
        total_inner = ctk.CTkFrame(total_card, fg_color="transparent")
        total_inner.pack(fill="x", padx=24, pady=16)
        ctk.CTkLabel(total_inner, text="VALOR TOTAL ESTIMADO EM AGIO",
                     font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_TERTIARY).pack(anchor="w")
        self.ag_total_value = ctk.CTkLabel(
            total_inner, text="Carregando...",
            font=("Segoe UI", 32, "bold"), text_color=ACCENT_GREEN)
        self.ag_total_value.pack(anchor="w", pady=(2, 0))
        self.ag_total_sub = ctk.CTkLabel(
            total_inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY)
        self.ag_total_sub.pack(anchor="w")

        # ======== TOP ASSESSORES ========
        ctk.CTkLabel(
            content, text="Top Assessores - Papeis em Agio",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        ranking_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                    border_width=1, border_color=BORDER_CARD)
        ranking_card.pack(fill="x", pady=(0, 18))
        self.ag_ranking_inner = ctk.CTkFrame(ranking_card, fg_color="transparent")
        self.ag_ranking_inner.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(self.ag_ranking_inner, text="Carregando...",
                     font=("Segoe UI", 11), text_color=TEXT_TERTIARY).pack(anchor="w")

        # ======== DISTRIBUICAO POR INDEXADOR ========
        ctk.CTkLabel(
            content, text="Distribuição por Indexador",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 10))

        dist_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                 border_width=1, border_color=BORDER_CARD)
        dist_card.pack(fill="x", pady=(0, 18))
        self.ag_dist_inner = ctk.CTkFrame(dist_card, fg_color="transparent")
        self.ag_dist_inner.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(self.ag_dist_inner, text="Carregando...",
                     font=("Segoe UI", 11), text_color=TEXT_TERTIARY).pack(anchor="w")

        # ======== ASSUNTO ========
        ctk.CTkLabel(
            content, text="Configurações do Envio",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 8))

        config_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=BORDER_CARD)
        config_card.pack(fill="x", pady=(0, 16))

        ci = ctk.CTkFrame(config_card, fg_color="transparent")
        ci.pack(fill="x", padx=20, pady=16)

        ctk.CTkLabel(
            ci, text="Tipo de Ativo",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 6))

        self.ag_tipo_ativo = ctk.CTkSegmentedButton(
            ci,
            values=["Tesouro Direto", "Crédito Privado", "Título Publico"],
            font=("Segoe UI", 11, "bold"),
            fg_color=BG_INPUT,
            selected_color=ACCENT_GREEN,
            selected_hover_color=BG_SIDEBAR_HOVER,
            unselected_color=BG_INPUT,
            unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY,
            corner_radius=8,
            height=38,
        )
        self.ag_tipo_ativo.pack(fill="x")
        self.ag_tipo_ativo.set("Crédito Privado")

        # ======== AÇÕES ========
        act_frame = ctk.CTkFrame(content, fg_color="transparent")
        act_frame.pack(fill="x", pady=(14, 14))

        self.ag_enviar_btn = ctk.CTkButton(
            act_frame, text="\u2709  Gerar Rascunhos no Outlook",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            height=44, corner_radius=10,
            command=self._on_ag_gerar_emails,
        )
        self.ag_enviar_btn.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            act_frame, text="\u2702  Limpar Planilha",
            font=("Segoe UI", 11, "bold"),
            fg_color=ACCENT_RED, hover_color="#b02a37",
            height=44, corner_radius=10, width=160,
            command=self._on_ag_limpar_xlsx,
        ).pack(side="left")

        # ======== STATUS ========
        self.ag_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.ag_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.ag_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.ag_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.ag_status_dot.pack(side="left")

        self.ag_status_text = ctk.CTkLabel(
            si, text="  Preencha o assunto e gere os rascunhos",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.ag_status_text.pack(side="left")

        # Atualizar (abre pasta da planilha)
        os.makedirs(AGIO_DIR, exist_ok=True)
        self._make_atualizar_btn(content, os.path.join(AGIO_DIR, "agio.xlsx"))

        # Internal state
        self._ag_data = []
        self._ag_headers = []
        self._ag_loaded = False

        return page

    # -----------------------------------------------------------------
    #  INFORMATIVO AGIO: Carregamento e Dashboard
    # -----------------------------------------------------------------
    def _ag_load_data_thread(self):
        """Le XLSX em background thread (read_only) e pre-computa dashboard."""
        from collections import Counter, defaultdict

        os.makedirs(AGIO_DIR, exist_ok=True)
        xlsx_files = [f for f in os.listdir(AGIO_DIR)
                      if f.lower().endswith(".xlsx") and not f.startswith("~")]
        if not xlsx_files:
            self._ag_data = []
            self._ag_headers = []
            self.after(0, lambda: self._ag_update_dashboard(None))
            return

        path = os.path.join(AGIO_DIR, xlsx_files[0])
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            rows_iter = ws.iter_rows(values_only=True)

            first_row = next(rows_iter, None)
            if not first_row:
                wb.close()
                self._ag_data = []
                self._ag_headers = []
                self.after(0, lambda: self._ag_update_dashboard(None))
                return

            headers = [str(h) if h else f"Col{i}"
                       for i, h in enumerate(first_row, 1)]
            rows = []
            for row_vals in rows_iter:
                if any(v is not None for v in row_vals):
                    rows.append(dict(zip(headers, row_vals)))
            wb.close()

            self._ag_data = rows
            self._ag_headers = headers

            # --- Pre-computar dados do dashboard na thread ---
            def _find_col(*keywords):
                for h in headers:
                    hl = h.lower()
                    if all(k in hl for k in keywords):
                        return h
                return None

            col_assessor = _find_col("assessor")
            col_agio_pct = _find_col("gio", "%")
            col_agio_rs = _find_col("gio", "r$")
            col_index = _find_col("indexador")

            filtered = []
            for row in rows:
                val = row.get(col_agio_pct, 0) if col_agio_pct else 0
                try:
                    num = float(val) if val is not None else 0
                except (ValueError, TypeError):
                    continue
                if num > 0.005:
                    filtered.append(row)

            total_rs = 0.0
            for r in filtered:
                try:
                    total_rs += float(r.get(col_agio_rs, 0) or 0)
                except (ValueError, TypeError):
                    pass

            assessores_set = set()
            for r in filtered:
                cod = str(r.get(col_assessor, "")).strip() if col_assessor else ""
                if cod:
                    assessores_set.add(cod)

            idx_counts = Counter()
            if col_index:
                for r in filtered:
                    idx = str(r.get(col_index, "")).strip()
                    if idx:
                        idx_counts[idx] += 1
            top_idx = idx_counts.most_common(1)[0][0] if idx_counts else "--"

            by_assessor = defaultdict(int)
            for r in filtered:
                cod = str(r.get(col_assessor, "")).strip() if col_assessor else ""
                if cod:
                    by_assessor[cod] += 1

            info = {
                "path": path,
                "total_rs": total_rs,
                "n_assessores": len(assessores_set),
                "n_papeis": len(filtered),
                "top_idx": top_idx,
                "top5": sorted(by_assessor.items(), key=lambda x: -x[1])[:5],
                "top_idxs": idx_counts.most_common(6),
                "n_total": len(rows),
            }
            self.after(0, lambda: self._ag_update_dashboard(info))

        except Exception as e:
            self._ag_data = []
            self._ag_headers = []
            self.after(0, lambda err=e: self._ag_update_dashboard({"error": err}))

    def _ag_load_data(self):
        """Carrega dados (sync) — chamado pelo botao Recarregar e Gerar Emails."""
        self._ag_loaded = True
        threading.Thread(target=self._ag_load_data_thread, daemon=True).start()

    def _ag_update_dashboard(self, info):
        """Atualiza widgets do dashboard com dados pre-computados (main thread)."""
        if info is None:
            self.ag_kpi_total.configure(text="--")
            self.ag_kpi_assessores.configure(text="--")
            self.ag_kpi_papeis.configure(text="--")
            self.ag_kpi_indexador.configure(text="--")
            self.ag_total_value.configure(text="Sem dados")
            self.ag_total_sub.configure(
                text="Insira planilha .xlsx via botao Atualizar")
            for w in self.ag_ranking_inner.winfo_children():
                w.destroy()
            ctk.CTkLabel(self.ag_ranking_inner, text="Sem dados",
                         font=("Segoe UI", 11),
                         text_color=TEXT_TERTIARY).pack(anchor="w")
            for w in self.ag_dist_inner.winfo_children():
                w.destroy()
            ctk.CTkLabel(self.ag_dist_inner, text="Sem dados",
                         font=("Segoe UI", 11),
                         text_color=TEXT_TERTIARY).pack(anchor="w")
            self.ag_status_dot.configure(text_color=ACCENT_RED)
            self.ag_status_text.configure(
                text="  Nenhuma planilha encontrada em Mesa Produtos/Info Agio/Agio/")
            return

        if "error" in info:
            self.ag_status_dot.configure(text_color=ACCENT_RED)
            self.ag_status_text.configure(
                text=f"  Erro ao carregar: {info['error']}")
            return

        total_rs = info["total_rs"]
        n_assessores = info["n_assessores"]
        n_papeis = info["n_papeis"]
        top_idx = info["top_idx"]
        top5 = info["top5"]
        top_idxs = info["top_idxs"]
        n_total = info["n_total"]
        path = info["path"]

        # --- KPI cards ---
        fmt_total = f"R$ {abs(total_rs):,.0f}".replace(",", ".")
        self.ag_kpi_total.configure(text=fmt_total)
        self.ag_kpi_assessores.configure(text=str(n_assessores))
        self.ag_kpi_papeis.configure(text=str(n_papeis))
        self.ag_kpi_indexador.configure(text=top_idx)

        # --- Total destaque ---
        fmt_full = f"R$ {abs(total_rs):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.ag_total_value.configure(text=fmt_full)
        self.ag_total_sub.configure(
            text=f"{n_papeis} papeis com agio > 0,50%  |  "
                 f"{n_assessores} assessores  |  "
                 f"{n_total} registros totais")

        # --- Top Assessores ranking ---
        for w in self.ag_ranking_inner.winfo_children():
            w.destroy()
        max_count = top5[0][1] if top5 else 1
        for i, (cod, count) in enumerate(top5):
            row_f = ctk.CTkFrame(self.ag_ranking_inner, fg_color="transparent")
            row_f.pack(fill="x", pady=3)
            ctk.CTkLabel(
                row_f, text=f"{i + 1}.", font=("Segoe UI", 11, "bold"),
                text_color=TEXT_TERTIARY, width=24
            ).pack(side="left")
            ctk.CTkLabel(
                row_f, text=cod, font=("Segoe UI", 11, "bold"),
                text_color=TEXT_PRIMARY, width=80
            ).pack(side="left", padx=(0, 12))
            bar_bg = ctk.CTkFrame(
                row_f, fg_color=BG_PROGRESS_TRACK, height=18, corner_radius=4)
            bar_bg.pack(side="left", fill="x", expand=True, padx=(0, 12))
            bar_bg.pack_propagate(False)
            pct = count / max_count
            bar_fill = ctk.CTkFrame(
                bar_bg, fg_color=ACCENT_GREEN, corner_radius=4)
            bar_fill.place(relx=0, rely=0, relwidth=pct, relheight=1)
            ctk.CTkLabel(
                row_f, text=str(count), font=("Segoe UI", 11, "bold"),
                text_color=ACCENT_GREEN, width=40
            ).pack(side="right")

        # --- Distribuição por Indexador ---
        for w in self.ag_dist_inner.winfo_children():
            w.destroy()
        max_idx = top_idxs[0][1] if top_idxs else 1
        colors = [ACCENT_GREEN, ACCENT_BLUE, ACCENT_ORANGE,
                  ACCENT_PURPLE, ACCENT_TEAL, ACCENT_RED]
        for i, (idx_name, count) in enumerate(top_idxs):
            row_f = ctk.CTkFrame(self.ag_dist_inner, fg_color="transparent")
            row_f.pack(fill="x", pady=3)
            ctk.CTkLabel(
                row_f, text=sanitize_text(idx_name),
                font=("Segoe UI", 11), text_color=TEXT_PRIMARY,
                width=120, anchor="w"
            ).pack(side="left", padx=(0, 12))
            bar_bg = ctk.CTkFrame(
                row_f, fg_color=BG_PROGRESS_TRACK, height=18, corner_radius=4)
            bar_bg.pack(side="left", fill="x", expand=True, padx=(0, 12))
            bar_bg.pack_propagate(False)
            pct = count / max_idx
            c = colors[i % len(colors)]
            bar_fill = ctk.CTkFrame(bar_bg, fg_color=c, corner_radius=4)
            bar_fill.place(relx=0, rely=0, relwidth=pct, relheight=1)
            ctk.CTkLabel(
                row_f, text=str(count), font=("Segoe UI", 11, "bold"),
                text_color=c, width=40
            ).pack(side="right")

        # --- Status ---
        self.ag_status_dot.configure(text_color=ACCENT_GREEN)
        self.ag_status_text.configure(
            text=f"  Dados carregados - {os.path.basename(path)}"
                 f"  |  {n_total} registros  |  {n_papeis} com agio > 0,50%")

    def _on_ag_limpar_xlsx(self):
        """Remove do XLSX todas as linhas com |agio %| < 0,50%."""
        os.makedirs(AGIO_DIR, exist_ok=True)
        xlsx_files = [f for f in os.listdir(AGIO_DIR)
                      if f.lower().endswith(".xlsx") and not f.startswith("~")]
        if not xlsx_files:
            messagebox.showwarning(
                "Sem planilha",
                "Nenhuma planilha .xlsx encontrada em Mesa Produtos/Info Agio/Agio/.\n"
                "Clique em 'Atualizar' para inserir o arquivo.")
            return

        path = os.path.join(AGIO_DIR, xlsx_files[0])
        resp = messagebox.askyesno(
            "Limpar Planilha",
            f"Isso vai REMOVER do arquivo todas as linhas com\n"
            f"|agio| inferior a 0,50%.\n\n"
            f"Arquivo: {os.path.basename(path)}\n\n"
            "Esta acao não pode ser desfeita. Continuar?")
        if not resp:
            return

        self.ag_status_dot.configure(text_color=ACCENT_ORANGE)
        self.ag_status_text.configure(text="  Limpando planilha...")

        threading.Thread(
            target=self._run_limpar_xlsx, args=(path,), daemon=True
        ).start()

    def _run_limpar_xlsx(self, path):
        """Thread: reescreve XLSX mantendo so linhas com agio >= 0,50%."""
        try:
            from openpyxl import Workbook
            wb_src = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws_src = wb_src[wb_src.sheetnames[0]]
            rows_iter = ws_src.iter_rows(values_only=True)

            header_row = next(rows_iter, None)
            if not header_row:
                wb_src.close()
                self.after(0, lambda: messagebox.showerror(
                    "Erro", "Planilha vazia."))
                return

            # Encontrar coluna de agio %
            agio_col = None
            for i, val in enumerate(header_row):
                if val and "gio" in str(val).lower() and "%" in str(val):
                    agio_col = i
                    break

            if agio_col is None:
                wb_src.close()
                self.after(0, lambda: messagebox.showerror(
                    "Erro", "Coluna de Agio (%) não encontrada."))
                return

            # Ler todas as linhas e filtrar
            kept_rows = []
            total_src = 0
            for row_vals in rows_iter:
                if not any(v is not None for v in row_vals):
                    continue
                total_src += 1
                val = row_vals[agio_col]
                try:
                    num = float(val) if val is not None else 0
                except (ValueError, TypeError):
                    num = 0
                if num > 0.005:
                    kept_rows.append(row_vals)

            wb_src.close()

            # Criar novo workbook com linhas filtradas
            wb_new = Workbook()
            ws_new = wb_new.active
            ws_new.title = "Export"
            for c, val in enumerate(header_row, 1):
                ws_new.cell(1, c, val)
            for r, row_vals in enumerate(kept_rows, 2):
                for c, val in enumerate(row_vals, 1):
                    ws_new.cell(r, c, val)

            kept = len(kept_rows)
            removed = total_src - kept

            wb_new.save(path)
            wb_new.close()

            def _done():
                self.ag_status_dot.configure(text_color=ACCENT_GREEN)
                self.ag_status_text.configure(
                    text=f"  Limpeza concluída - {removed} removidas"
                         f"  |  {kept} restantes")
                messagebox.showinfo(
                    "Limpeza Concluida",
                    f"{removed} linhas com |agio| < 0,50% removidas.\n"
                    f"{kept} linhas restantes no arquivo.")
                self._ag_load_data()
            self.after(0, _done)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror(
                "Erro ao limpar", str(e)))
            self.after(0, lambda: self.ag_status_dot.configure(
                text_color=ACCENT_RED))
            self.after(0, lambda: self.ag_status_text.configure(
                text=f"  Erro: {e}"))

    def _on_ag_gerar_emails(self):
        if not self._ag_data:
            messagebox.showwarning(
                "Dados não carregados",
                "Aguarde o carregamento ou clique em Recarregar.")
            return

        tipo_ativo = self.ag_tipo_ativo.get()

        # Contar registros filtrados e assessores (dados ja em memoria)
        headers = getattr(self, "_ag_headers", [])
        col_agio_pct = None
        col_assessor = None
        for h in headers:
            hl = h.lower()
            if "gio" in hl and "%" in hl:
                col_agio_pct = h
            if "assessor" in hl:
                col_assessor = h

        n_filtrados = 0
        assessores_set = set()
        for row in self._ag_data:
            agio_val = row.get(col_agio_pct, 0) if col_agio_pct else 0
            try:
                agio_num = float(agio_val) if agio_val is not None else 0
            except (ValueError, TypeError):
                continue
            if agio_num > 0.005:
                n_filtrados += 1
                cod = str(row.get(col_assessor, "")).strip() if col_assessor else ""
                if cod:
                    assessores_set.add(cod)

        if n_filtrados == 0:
            messagebox.showinfo("Sem dados", "Nenhum registro com agio acima de 0,50%.")
            return

        resp = messagebox.askyesno(
            "Gerar Rascunhos",
            f"Serao criados rascunhos no Outlook:\n\n"
            f"Tipo: {tipo_ativo}\n"
            f"{n_filtrados} registros com agio > 0,50%\n"
            f"{len(assessores_set)} assessores\n\n"
            "O Outlook precisa estar aberto. Continuar?"
        )
        if not resp:
            return

        self.ag_enviar_btn.configure(state="disabled")
        self.ag_status_dot.configure(text_color=ACCENT_ORANGE)
        self.ag_status_text.configure(text="  Gerando rascunhos...")
        threading.Thread(
            target=self._run_ag_emails,
            args=(tipo_ativo,),
            daemon=True
        ).start()

    def _run_ag_emails(self, tipo_ativo):
        try:
            import win32com.client as win32
            import pythoncom
            from collections import defaultdict
            pythoncom.CoInitialize()

            outlook = win32.Dispatch("Outlook.Application")
            assessores = load_base_emails()

            hoje = datetime.now().strftime("%d/%m/%Y")
            logo_tag = LOGO_TAG_CID if os.path.exists(LOGO_PATH) else ""

            headers = self._ag_headers

            # --- Localizar colunas por correspondencia parcial ---
            def _find_col(*keywords):
                for h in headers:
                    hl = h.lower()
                    if all(k in hl for k in keywords):
                        return h
                return None

            col_assessor = _find_col("assessor") or headers[1]
            col_agio_pct = _find_col("gio", "%") or headers[-2]
            col_agio_rs = _find_col("gio", "r$") or headers[-1]
            col_conta = _find_col("conta")
            col_ativo = _find_col("nome", "ativo") or _find_col("ativo")
            col_ticker = _find_col("ticker")
            col_venc = _find_col("vencimento")
            col_index = _find_col("indexador")
            col_valor = _find_col("valor", "aplic")
            col_posição = _find_col("posi")

            # Colunas exibidas no email (chave_original, label_email)
            email_cols = []
            if col_conta:    email_cols.append((col_conta, "Conta"))
            if col_ativo:    email_cols.append((col_ativo, "Ativo"))
            if col_ticker:   email_cols.append((col_ticker, "Ticker"))
            if col_venc:     email_cols.append((col_venc, "Vencimento"))
            if col_index:    email_cols.append((col_index, "Indexador"))
            if col_valor:    email_cols.append((col_valor, "Valor Aplicado"))
            if col_posição:  email_cols.append((col_posição, "Pos. Atual"))
            if col_agio_pct: email_cols.append((col_agio_pct, "Agio (%)"))
            if col_agio_rs:  email_cols.append((col_agio_rs, "Agio (R$)"))

            money_cols = {col_valor, col_posição, col_agio_rs}
            pct_cols = {col_agio_pct}
            date_cols = {col_venc}

            # --- Filtrar: |agio %| > 0,50% (0.005 em decimal) ---
            filtered = []
            for row in self._ag_data:
                agio_val = row.get(col_agio_pct, 0)
                try:
                    agio_num = float(agio_val) if agio_val is not None else 0
                except (ValueError, TypeError):
                    continue
                if agio_num > 0.005:
                    filtered.append(row)

            # --- Agrupar por assessor e ordenar por financeiro (Agio R$) desc ---
            by_assessor = defaultdict(list)
            for row in filtered:
                cod = str(row.get(col_assessor, "")).strip()
                if cod:
                    by_assessor[cod].append(row)
            for cod in by_assessor:
                by_assessor[cod].sort(
                    key=lambda r: abs(float(r.get(col_agio_rs, 0) or 0)),
                    reverse=True)

            # --- Montar tabela HTML para um assessor ---
            def _build_table(rows):
                hdr_html = "".join(
                    f'<td style="padding:4px 10px;font-weight:bold;color:#00785a;'
                    f'font-size:8.5pt;border-bottom:1.5px solid #00785a;white-space:nowrap;">'
                    f'{sanitize_text(label)}</td>'
                    for _col_key, label in email_cols
                )
                rows_html = ""
                for i, row in enumerate(rows):
                    bg = "#f7faf9" if i % 2 == 0 else "#ffffff"
                    cells = ""
                    for col_key, _label in email_cols:
                        val = row.get(col_key, "")
                        color = "#1a1a2e"
                        if col_key in date_cols and val is not None:
                            try:
                                if hasattr(val, "strftime"):
                                    val = val.strftime("%d/%m/%y")
                                else:
                                    s = str(val).strip()[:10]
                                    dt = datetime.strptime(s, "%Y-%m-%d")
                                    val = dt.strftime("%d/%m/%y")
                            except (ValueError, TypeError):
                                val = str(val) if val else ""
                        elif col_key in pct_cols and val is not None:
                            try:
                                v = float(val)
                                val = f"{v * 100:.2f}%"
                                if v < 0:
                                    color = "#c0392b"
                            except (ValueError, TypeError):
                                val = str(val) if val else ""
                        elif col_key in money_cols and val is not None:
                            try:
                                v = float(val)
                                val = f"R$ {v:,.2f}"
                                if v < 0:
                                    color = "#c0392b"
                            except (ValueError, TypeError):
                                val = str(val) if val else ""
                        else:
                            val = str(val) if val is not None else ""
                        cells += (
                            f'<td style="padding:3px 10px;color:{color};'
                            f'border-bottom:1px solid #eef1ef;font-size:8.5pt;'
                            f'white-space:nowrap;">{sanitize_text(val)}</td>'
                        )
                    rows_html += f'<tr style="background:{bg};">{cells}</tr>\n'
                return (
                    f'<table cellpadding="0" cellspacing="0" border="0" '
                    f'style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;'
                    f'font-size:9pt;margin-bottom:6px;">'
                    f'<tr>{hdr_html}</tr>\n{rows_html}</table>'
                )

            # --- HTML completo do email (padrao Saldos Diários) ---
            def _build_html(primeiro_nome, data_table):
                return f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:960px;margin:0 auto;">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding:16px 0 12px 0;vertical-align:middle;">
      {logo_tag}<!--
      --><span style="font-size:16pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:300;"> | </span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:bold;">Produtos</span>
    </td>
  </tr>
  <tr>
    <td style="padding:0 0 16px 0;">
      <hr style="border:none;border-top:2px solid #004d33;margin:0;">
    </td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px;">
  <tr><td style="padding:0 4px;">
    <p style="font-size:11pt;color:#1a1a2e;margin-bottom:4px;">
      Prezado(a) <b>{primeiro_nome}</b>, tudo bem?
    </p>
    <p style="font-size:10.5pt;color:#4b5563;margin-top:0;">
      Segue abaixo o informativo de &aacute;gio/des&aacute;gio dos seus clientes.
    </p>
  </td></tr>
  <tr><td style="padding:22px 0 8px 0;">
    <table cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background-color:#00b876;width:4px;border-radius:4px;">&nbsp;</td>
      <td style="padding-left:12px;">
        <span style="font-size:12.5pt;color:#004d33;font-weight:bold;letter-spacing:0.3px;">
          Ativos com &Aacute;gio/Des&aacute;gio acima de 0,50%
        </span>
      </td>
    </tr></table>
  </td></tr>
  <tr><td style="padding:0 4px;">
    {data_table}
  </td></tr>
  <tr><td style="padding:28px 0 0 0;">
    <hr style="border:none;border-top:2px solid #004d33;margin:0 0 12px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="vertical-align:middle;">
          <span style="font-size:10pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:300;"> | </span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:bold;">Produtos</span>
        </td>
        <td style="text-align:right;">
          <span style="font-size:8.5pt;color:#6b7280;">Informativo &Aacute;gio &middot; {hoje}</span><br>
          <span style="font-size:7.5pt;color:#9ca3af;">E-mail gerado automaticamente</span>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</div>
"""

            criados = 0
            erros = 0
            ignorados = 0

            for cod_assess, rows in sorted(by_assessor.items()):
                info = assessores.get(cod_assess)
                if not info:
                    ignorados += 1
                    continue
                email = info.get("email", "")
                if not email or email == "-":
                    ignorados += 1
                    continue

                nome = info["nome"]
                primeiro = nome.split()[0] if nome else "Assessor"
                tabela = _build_table(rows)

                try:
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    email_assist = info.get("email_assistente", "-")
                    if email_assist and email_assist != "-":
                        mail.CC = email_assist
                    mail.Subject = (
                        f"{cod_assess} | Segue a sua relação de agio de "
                        f"{tipo_ativo}")
                    mail.HTMLBody = _build_html(primeiro, tabela)
                    _attach_logo_cid(mail)
                    mail.Save()
                    criados += 1
                except Exception:
                    erros += 1

            pythoncom.CoUninitialize()

            def _done():
                self.ag_enviar_btn.configure(state="normal")
                self.ag_status_dot.configure(text_color="#00a86b")
                self.ag_status_text.configure(
                    text=f"  Concluído - {criados} rascunhos, {erros} erros, {ignorados} sem email"
                )
                messagebox.showinfo("Rascunhos Criados",
                    f"{criados} rascunhos criados no Outlook!\n"
                    f"{erros} erros\n{ignorados} assessores sem email cadastrado")
            self.after(0, _done)

        except Exception as e:
            def _err():
                self.ag_enviar_btn.configure(state="normal")
                self.ag_status_dot.configure(text_color=ACCENT_RED)
                self.ag_status_text.configure(text=f"  Erro: {e}")
                messagebox.showerror("Erro", str(e))
            self.after(0, _err)

    # -----------------------------------------------------------------
    #  PAGE: ENVIO DE ORDENS
    # -----------------------------------------------------------------
    def _build_envio_ordens_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Envio de Ordens", subtitle="Mesa de Produtos")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Banner
        banner = ctk.CTkFrame(content, fg_color=ACCENT_GREEN, corner_radius=10, height=44)
        banner.pack(fill="x", pady=(0, 18))
        banner.pack_propagate(False)

        ctk.CTkLabel(
            banner, text="  Envio de ordens por e-mail a partir de planilha",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        ).pack(side="left", padx=18, pady=10)

        # ======== INSERIR LISTA DE EMAILS (botao topo) ========
        ctk.CTkLabel(
            content, text="Lista de Destinatarios",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 6))

        self.eo_lista_card = ctk.CTkFrame(
            content, fg_color="#f0f4ff", corner_radius=12,
            border_width=2, border_color=ACCENT_BLUE
        )
        self.eo_lista_card.pack(fill="x", pady=(0, 16))

        lista_inner = ctk.CTkFrame(self.eo_lista_card, fg_color="transparent")
        lista_inner.pack(fill="x", padx=20, pady=16)

        lista_left = ctk.CTkFrame(lista_inner, fg_color="transparent")
        lista_left.pack(side="left", fill="x", expand=True)

        self.eo_lista_icon = ctk.CTkLabel(
            lista_left, text="\u2709",
            font=("Segoe UI", 20), text_color=ACCENT_BLUE
        )
        self.eo_lista_icon.pack(side="left")

        self.eo_lista_info = ctk.CTkLabel(
            lista_left, text="  Nenhuma lista carregada",
            font=("Segoe UI", 12), text_color=TEXT_SECONDARY, anchor="w"
        )
        self.eo_lista_info.pack(side="left", padx=(6, 0))

        ctk.CTkButton(
            lista_inner, text="  Inserir Lista de Emails",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=38, corner_radius=8, width=200,
            command=self._on_eo_browse_lista,
        ).pack(side="right")

        # Preview da lista carregada
        self.eo_preview_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.eo_preview_frame.pack(fill="x", pady=(0, 16))
        self.eo_preview_frame.pack_forget()

        self.eo_preview_text = ctk.CTkTextbox(
            self.eo_preview_frame, font=("Consolas", 9),
            fg_color=BG_CARD, text_color=TEXT_PRIMARY,
            corner_radius=10, height=140, wrap="none"
        )
        self.eo_preview_text.pack(fill="x", padx=4, pady=4)
        self.eo_preview_text.configure(state="disabled")

        # ======== CONFIGURAÇÕES ========
        ctk.CTkLabel(
            content, text="Configurações do Envio",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 8))

        config_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=BORDER_CARD)
        config_card.pack(fill="x", pady=(0, 16))

        ci = ctk.CTkFrame(config_card, fg_color="transparent")
        ci.pack(fill="x", padx=20, pady=16)

        # Row 1: Modo de envio + Tipo de ativo
        row1 = ctk.CTkFrame(ci, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 14))
        row1.columnconfigure((0, 1), weight=1)

        # -- Modo de envio --
        modo_frame = ctk.CTkFrame(row1, fg_color="transparent")
        modo_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        ctk.CTkLabel(
            modo_frame, text="Modo de Envio",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 6))

        self.eo_modo = ctk.CTkSegmentedButton(
            modo_frame,
            values=["Grupo", "Individual"],
            font=("Segoe UI", 11, "bold"),
            fg_color=BG_INPUT,
            selected_color=ACCENT_GREEN,
            selected_hover_color=BG_SIDEBAR_HOVER,
            unselected_color=BG_INPUT,
            unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY,
            corner_radius=8,
            height=38,
        )
        self.eo_modo.pack(fill="x")
        self.eo_modo.set("Individual")

        # -- Tipo de ativo --
        tipo_frame = ctk.CTkFrame(row1, fg_color="transparent")
        tipo_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        ctk.CTkLabel(
            tipo_frame, text="Tipo de Ativo",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 6))

        self.eo_tipo_ativo = ctk.CTkSegmentedButton(
            tipo_frame,
            values=["Fundos", "Renda Variavel", "Renda Fixa", "Alternativo"],
            font=("Segoe UI", 11, "bold"),
            fg_color=BG_INPUT,
            selected_color=ACCENT_BLUE,
            selected_hover_color="#1555bb",
            unselected_color=BG_INPUT,
            unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY,
            corner_radius=8,
            height=38,
        )
        self.eo_tipo_ativo.pack(fill="x")
        self.eo_tipo_ativo.set("Fundos")

        # Row 2: Dados da Ordem (Ativo, Quantidade, Financeiro, Cotação)
        ctk.CTkLabel(
            ci, text="Dados da Ordem",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(14, 6))

        ordem_row = ctk.CTkFrame(ci, fg_color="transparent")
        ordem_row.pack(fill="x", pady=(0, 14))
        ordem_row.columnconfigure((0, 1, 2, 3), weight=1)

        # Ativo *
        ativo_frame = ctk.CTkFrame(ordem_row, fg_color="transparent")
        ativo_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(
            ativo_frame, text="Ativo *",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")

        self.eo_ativo = ctk.CTkEntry(
            ativo_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: XPML11"
        )
        self.eo_ativo.pack(fill="x", pady=(2, 0))

        # Quantidade *
        qtd_frame = ctk.CTkFrame(ordem_row, fg_color="transparent")
        qtd_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 6))

        ctk.CTkLabel(
            qtd_frame, text="Quantidade *",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")

        self.eo_quantidade = ctk.CTkEntry(
            qtd_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: 100"
        )
        self.eo_quantidade.pack(fill="x", pady=(2, 0))

        # Financeiro *
        fin_frame = ctk.CTkFrame(ordem_row, fg_color="transparent")
        fin_frame.grid(row=0, column=2, sticky="nsew", padx=(6, 6))

        ctk.CTkLabel(
            fin_frame, text="Financeiro *",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")

        self.eo_financeiro = ctk.CTkEntry(
            fin_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: 10000.00"
        )
        self.eo_financeiro.pack(fill="x", pady=(2, 0))

        # Cotação (opcional)
        cot_frame = ctk.CTkFrame(ordem_row, fg_color="transparent")
        cot_frame.grid(row=0, column=3, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(
            cot_frame, text="Cotação",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")

        self.eo_cotação = ctk.CTkEntry(
            cot_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Opcional"
        )
        self.eo_cotação.pack(fill="x", pady=(2, 0))

        # ===== VENDA — Ativo de Saída (opcional) =====
        venda_sep = ctk.CTkFrame(ci, fg_color="#fff3e8", corner_radius=8,
                                 border_width=1, border_color=ACCENT_ORANGE)
        venda_sep.pack(fill="x", pady=(18, 10))

        venda_sep_inner = ctk.CTkFrame(venda_sep, fg_color="transparent")
        venda_sep_inner.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(
            venda_sep_inner, text="\u2193  Venda \u2014 Ativo de Sa\u00edda",
            font=("Segoe UI", 11, "bold"), text_color=ACCENT_ORANGE, anchor="w"
        ).pack(side="left")

        ctk.CTkLabel(
            venda_sep_inner,
            text="(opcional \u2014 preencha caso o cliente precise sair de algum ativo)",
            font=("Segoe UI", 9), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(side="left", padx=(10, 0))

        venda_row = ctk.CTkFrame(ci, fg_color="transparent")
        venda_row.pack(fill="x", pady=(0, 4))
        venda_row.columnconfigure((0, 1, 2, 3), weight=1)

        # Ativo de saída
        vativo_frame = ctk.CTkFrame(venda_row, fg_color="transparent")
        vativo_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ctk.CTkLabel(
            vativo_frame, text="Ativo",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")
        self.eo_venda_ativo = ctk.CTkEntry(
            vativo_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=ACCENT_ORANGE,
            placeholder_text="Ex: FIXA2028"
        )
        self.eo_venda_ativo.pack(fill="x", pady=(2, 0))

        # Quantidade de saída
        vqtd_frame = ctk.CTkFrame(venda_row, fg_color="transparent")
        vqtd_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 6))
        ctk.CTkLabel(
            vqtd_frame, text="Quantidade",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")
        self.eo_venda_quantidade = ctk.CTkEntry(
            vqtd_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=ACCENT_ORANGE,
            placeholder_text="Ex: 100"
        )
        self.eo_venda_quantidade.pack(fill="x", pady=(2, 0))

        # Financeiro de saída
        vfin_frame = ctk.CTkFrame(venda_row, fg_color="transparent")
        vfin_frame.grid(row=0, column=2, sticky="nsew", padx=(6, 6))
        ctk.CTkLabel(
            vfin_frame, text="Financeiro",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")
        self.eo_venda_financeiro = ctk.CTkEntry(
            vfin_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=ACCENT_ORANGE,
            placeholder_text="Ex: 10000.00"
        )
        self.eo_venda_financeiro.pack(fill="x", pady=(2, 0))

        # Cotação de saída
        vcot_frame = ctk.CTkFrame(venda_row, fg_color="transparent")
        vcot_frame.grid(row=0, column=3, sticky="nsew", padx=(6, 0))
        ctk.CTkLabel(
            vcot_frame, text="Cota\u00e7\u00e3o",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        ).pack(fill="x")
        self.eo_venda_cotacao = ctk.CTkEntry(
            vcot_frame, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=ACCENT_ORANGE,
            placeholder_text="Opcional"
        )
        self.eo_venda_cotacao.pack(fill="x", pady=(2, 0))

        # Row 3: Assunto
        ctk.CTkLabel(
            ci, text="Assunto do E-mail",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.eo_assunto = ctk.CTkEntry(
            ci, font=("Segoe UI", 12), height=40,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: Ordem de Compra - Fundos"
        )
        self.eo_assunto.pack(fill="x", pady=(0, 14))

        # Row 3: Corpo
        ctk.CTkLabel(
            ci, text="Corpo do E-mail",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.eo_corpo = ctk.CTkTextbox(
            ci, font=("Segoe UI", 11), height=160,
            corner_radius=8, border_width=1, border_color=BORDER_CARD,
            fg_color=BG_PRIMARY, text_color=TEXT_PRIMARY, wrap="word"
        )
        self.eo_corpo.pack(fill="x", pady=(0, 14))

        # Row 4: Anexo
        anexo_row = ctk.CTkFrame(ci, fg_color="transparent")
        anexo_row.pack(fill="x", pady=(0, 0))

        ctk.CTkLabel(
            anexo_row, text="Anexo (opcional)",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(side="left")

        self.eo_anexo_label = ctk.CTkLabel(
            anexo_row, text="Nenhum arquivo",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="e"
        )
        self.eo_anexo_label.pack(side="right", padx=(0, 8))

        ctk.CTkButton(
            anexo_row, text="Selecionar",
            font=("Segoe UI", 10), fg_color=BG_INPUT,
            hover_color=BORDER_CARD, text_color=TEXT_SECONDARY,
            height=30, corner_radius=6, width=90,
            command=self._on_eo_browse_anexo,
        ).pack(side="right", padx=(0, 4))

        ctk.CTkButton(
            anexo_row, text="Limpar",
            font=("Segoe UI", 10), fg_color="transparent",
            hover_color=BORDER_CARD, text_color=TEXT_TERTIARY,
            height=30, corner_radius=6, width=60,
            command=self._on_eo_clear_anexo,
        ).pack(side="right")

        # ======== AÇÕES ========
        act_frame = ctk.CTkFrame(content, fg_color="transparent")
        act_frame.pack(fill="x", pady=(6, 14))

        self.eo_enviar_btn = ctk.CTkButton(
            act_frame, text="\u2191  Gerar Rascunhos no Outlook",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            height=44, corner_radius=10,
            command=self._on_eo_gerar_emails,
        )
        self.eo_enviar_btn.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            act_frame, text="Limpar Tudo",
            font=("Segoe UI", 11),
            fg_color="#6b7280", hover_color="#555d6a",
            height=44, corner_radius=10, width=120,
            command=self._on_eo_limpar,
        ).pack(side="left")

        # ======== STATUS ========
        self.eo_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.eo_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.eo_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.eo_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.eo_status_dot.pack(side="left")

        self.eo_status_text = ctk.CTkLabel(
            si, text="  Carregue a lista de emails e configure o envio",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.eo_status_text.pack(side="left")

        # Internal state
        self._eo_lista_path = None
        self._eo_destinatarios = []
        self._eo_anexo_path = None

        return page

    # -----------------------------------------------------------------
    #  ENVIO DE ORDENS: Ações
    # -----------------------------------------------------------------
    def _on_eo_browse_lista(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha de emails",
            initialdir=os.path.join(BASE_DIR, "BASE"),
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        if not path:
            return

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb[wb.sheetnames[0]]

            # Detectar colunas pela header (row 1)
            headers = {}
            for col in range(1, ws.max_column + 1):
                val = ws.cell(1, col).value
                if val:
                    headers[str(val).strip().lower()] = col

            # Tentar encontrar colunas de nome e email
            col_nome = None
            col_email = None
            for key, col in headers.items():
                if not col_nome and any(k in key for k in ["nome", "assessor", "name"]):
                    col_nome = col
                if not col_email and any(k in key for k in ["email", "e-mail", "mail"]):
                    col_email = col

            if not col_email:
                # Fallback: primeira coluna texto, segunda com @
                for r in range(2, min(10, ws.max_row + 1)):
                    for c in range(1, ws.max_column + 1):
                        v = str(ws.cell(r, c).value or "")
                        if "@" in v and not col_email:
                            col_email = c
                        elif v and not col_nome and "@" not in v:
                            col_nome = c

            if not col_email:
                wb.close()
                messagebox.showwarning(
                    "Coluna não encontrada",
                    "Não foi possível encontrar uma coluna de e-mail na planilha.\n\n"
                    "A planilha deve ter uma coluna com header 'Email' ou conter enderecos de e-mail."
                )
                return

            destinatarios = []
            for r in range(2, ws.max_row + 1):
                email = str(ws.cell(r, col_email).value or "").strip()
                if not email or "@" not in email:
                    continue
                nome = sanitize_text(ws.cell(r, col_nome).value) if col_nome else ""
                destinatarios.append({"nome": nome, "email": email})

            wb.close()

            if not destinatarios:
                messagebox.showwarning("Lista vazia", "Nenhum e-mail válido encontrado na planilha.")
                return

            self._eo_lista_path = path
            self._eo_destinatarios = destinatarios

            # Atualizar UI
            self.eo_lista_icon.configure(text_color=ACCENT_GREEN)
            self.eo_lista_info.configure(
                text=f"  {len(destinatarios)} destinatarios  -  {os.path.basename(path)}",
                text_color=ACCENT_GREEN
            )

            # Preview
            self.eo_preview_frame.pack(fill="x", pady=(0, 16))
            self.eo_preview_text.configure(state="normal")
            self.eo_preview_text.delete("1.0", "end")

            header_line = f"{'Nome':<35} {'Email':<40}\n"
            sep_line = f"{'-'*35} {'-'*40}\n"
            self.eo_preview_text.insert("end", header_line)
            self.eo_preview_text.insert("end", sep_line)
            for d in destinatarios[:20]:
                line = f"{d['nome']:<35} {d['email']:<40}\n"
                self.eo_preview_text.insert("end", line)
            if len(destinatarios) > 20:
                self.eo_preview_text.insert("end", f"\n... e mais {len(destinatarios) - 20} destinatarios")
            self.eo_preview_text.configure(state="disabled")

            self.eo_status_dot.configure(text_color=ACCENT_GREEN)
            self.eo_status_text.configure(text=f"  Lista carregada - {len(destinatarios)} emails prontos")

        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))

    def _on_eo_browse_anexo(self):
        path = filedialog.askopenfilename(
            title="Selecionar anexo",
            filetypes=[("Todos", "*.*"), ("PDF", "*.pdf"), ("Excel", "*.xlsx"), ("Imagem", "*.png *.jpg")]
        )
        if path:
            self._eo_anexo_path = path
            self.eo_anexo_label.configure(text=os.path.basename(path), text_color=ACCENT_GREEN)

    def _on_eo_clear_anexo(self):
        self._eo_anexo_path = None
        self.eo_anexo_label.configure(text="Nenhum arquivo", text_color=TEXT_TERTIARY)

    def _on_eo_limpar(self):
        self._eo_lista_path = None
        self._eo_destinatarios = []
        self._eo_anexo_path = None
        self.eo_lista_icon.configure(text_color=ACCENT_BLUE)
        self.eo_lista_info.configure(text="  Nenhuma lista carregada", text_color=TEXT_SECONDARY)
        self.eo_preview_frame.pack_forget()
        self.eo_ativo.delete(0, "end")
        self.eo_quantidade.delete(0, "end")
        self.eo_financeiro.delete(0, "end")
        self.eo_cotação.delete(0, "end")
        self.eo_venda_ativo.delete(0, "end")
        self.eo_venda_quantidade.delete(0, "end")
        self.eo_venda_financeiro.delete(0, "end")
        self.eo_venda_cotacao.delete(0, "end")
        self.eo_assunto.delete(0, "end")
        self.eo_corpo.delete("1.0", "end")
        self.eo_anexo_label.configure(text="Nenhum arquivo", text_color=TEXT_TERTIARY)
        self.eo_modo.set("Individual")
        self.eo_tipo_ativo.set("Fundos")
        self.eo_status_dot.configure(text_color=TEXT_TERTIARY)
        self.eo_status_text.configure(text="  Carregue a lista de emails e configure o envio")

    def _on_eo_gerar_emails(self):
        if not self._eo_destinatarios:
            messagebox.showwarning("Lista vazia", "Insira a lista de emails primeiro.")
            return

        ativo = self.eo_ativo.get().strip()
        quantidade = self.eo_quantidade.get().strip()
        financeiro = self.eo_financeiro.get().strip()
        cotação = self.eo_cotação.get().strip()

        campos_faltando = []
        if not ativo:
            campos_faltando.append("Ativo")
        if not quantidade:
            campos_faltando.append("Quantidade")
        if not financeiro:
            campos_faltando.append("Financeiro")

        if campos_faltando:
            messagebox.showwarning(
                "Campos obrigatorios",
                f"Preencha os seguintes campos:\n\n- " + "\n- ".join(campos_faltando)
            )
            return

        assunto = self.eo_assunto.get().strip()
        corpo = self.eo_corpo.get("1.0", "end").strip()

        if not assunto:
            messagebox.showwarning("Campo obrigatorio", "Preencha o assunto do e-mail.")
            return
        if not corpo:
            messagebox.showwarning("Campo obrigatorio", "Preencha o corpo do e-mail.")
            return

        modo = self.eo_modo.get()
        tipo = self.eo_tipo_ativo.get()
        n = len(self._eo_destinatarios)

        if modo == "Grupo":
            msg = (f"Sera criado 1 rascunho com {n} destinatarios em copia.\n\n"
                   f"Tipo: {tipo}\nOutlook precisa estar aberto. Continuar?")
        else:
            msg = (f"Serao criados {n} rascunhos individuais.\n\n"
                   f"Tipo: {tipo}\nOutlook precisa estar aberto. Continuar?")

        if not messagebox.askyesno("Gerar Rascunhos", msg):
            return

        self.eo_enviar_btn.configure(state="disabled")
        self.eo_status_dot.configure(text_color=ACCENT_ORANGE)
        self.eo_status_text.configure(text="  Gerando rascunhos...")

        ordem_dados = {
            "ativo": ativo,
            "quantidade": quantidade,
            "financeiro": financeiro,
            "cotação": cotação,
        }
        venda_dados = {
            "ativo": self.eo_venda_ativo.get().strip(),
            "quantidade": self.eo_venda_quantidade.get().strip(),
            "financeiro": self.eo_venda_financeiro.get().strip(),
            "cotacao": self.eo_venda_cotacao.get().strip(),
        }
        threading.Thread(
            target=self._run_eo_emails,
            args=(assunto, corpo, modo, tipo, ordem_dados, venda_dados),
            daemon=True
        ).start()

    def _run_eo_emails(self, assunto, corpo, modo, tipo, ordem_dados, venda_dados=None):
        try:
            import win32com.client as win32
            import pythoncom
            pythoncom.CoInitialize()

            outlook = win32.Dispatch("Outlook.Application")

            hoje = datetime.now().strftime("%d/%m/%Y")
            logo_tag = LOGO_TAG_CID if os.path.exists(LOGO_PATH) else ""

            tipo_tag = f'<span style="display:inline-block;background:#004d33;color:#fff;' \
                       f'padding:3px 10px;border-radius:4px;font-size:9pt;margin-bottom:8px;">{tipo}</span>'

            def _get_primeiro(d):
                """Extrai primeiro nome do campo nome ou, como fallback, do email."""
                if d.get("nome"):
                    return d["nome"].split()[0].capitalize()
                local = d["email"].split("@")[0]        # nicolas.kersul
                parte = local.split(".")[0]             # nicolas
                return parte.capitalize()               # Nicolas

            # Tabela de dados da ordem
            ativo = ordem_dados["ativo"]
            quantidade = ordem_dados["quantidade"]
            financeiro = ordem_dados["financeiro"]
            cotação = ordem_dados["cotação"]

            ordem_rows = (
                f'<tr style="background:#f7faf9;">'
                f'<td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1.5px solid #00785a;font-size:8.5pt;">Ativo</td>'
                f'<td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1.5px solid #00785a;font-size:8.5pt;">Quantidade</td>'
                f'<td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1.5px solid #00785a;font-size:8.5pt;">Financeiro</td>'
                f'<td style="padding:5px 12px;color:#00785a;font-weight:bold;border-bottom:1.5px solid #00785a;font-size:8.5pt;">Cota&ccedil;&atilde;o</td>'
                f'</tr>'
                f'<tr style="background:#ffffff;">'
                f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;font-weight:bold;border-bottom:1px solid #eef1ef;">{ativo}</td>'
                f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;border-bottom:1px solid #eef1ef;">{quantidade}</td>'
                f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;border-bottom:1px solid #eef1ef;">{financeiro}</td>'
                f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;border-bottom:1px solid #eef1ef;">{cotação if cotação else "-"}</td>'
                f'</tr>'
            )

            ordem_table = (
                f'<table cellpadding="0" cellspacing="0" border="0" '
                f'style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;margin:10px 0 6px 0;">'
                f'{ordem_rows}</table>'
            )

            # Tabela de venda (opcional)
            venda_table = ""
            if venda_dados and any(venda_dados.get(k) for k in ("ativo", "quantidade", "financeiro")):
                v_ativo = venda_dados.get("ativo", "") or "-"
                v_qtd = venda_dados.get("quantidade", "") or "-"
                v_fin = venda_dados.get("financeiro", "") or "-"
                v_cot = venda_dados.get("cotacao", "") or "-"
                venda_rows = (
                    f'<tr style="background:#fff3e8;">'
                    f'<td style="padding:5px 12px;color:#b85c00;font-weight:bold;border-bottom:1.5px solid #e6832a;font-size:8.5pt;">Ativo</td>'
                    f'<td style="padding:5px 12px;color:#b85c00;font-weight:bold;border-bottom:1.5px solid #e6832a;font-size:8.5pt;">Quantidade</td>'
                    f'<td style="padding:5px 12px;color:#b85c00;font-weight:bold;border-bottom:1.5px solid #e6832a;font-size:8.5pt;">Financeiro</td>'
                    f'<td style="padding:5px 12px;color:#b85c00;font-weight:bold;border-bottom:1.5px solid #e6832a;font-size:8.5pt;">Cota&ccedil;&atilde;o</td>'
                    f'</tr>'
                    f'<tr style="background:#ffffff;">'
                    f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;font-weight:bold;border-bottom:1px solid #eef1ef;">{v_ativo}</td>'
                    f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;border-bottom:1px solid #eef1ef;">{v_qtd}</td>'
                    f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;border-bottom:1px solid #eef1ef;">{v_fin}</td>'
                    f'<td style="padding:5px 12px;color:#1a1a2e;font-size:9.5pt;border-bottom:1px solid #eef1ef;">{v_cot}</td>'
                    f'</tr>'
                )
                venda_table = (
                    f'<table cellpadding="0" cellspacing="0" border="0" '
                    f'style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;margin:10px 0 6px 0;">'
                    f'{venda_rows}</table>'
                )

            def _build_html(primeiro_nome, corpo_html):
                return f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:960px;margin:0 auto;">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding:16px 0 12px 0;vertical-align:middle;">
      {logo_tag}<!--
      --><span style="font-size:16pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:300;"> | </span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:bold;">Produtos</span>
    </td>
  </tr>
  <tr>
    <td style="padding:0 0 16px 0;">
      <hr style="border:none;border-top:2px solid #004d33;margin:0;">
    </td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px;">
  <tr><td style="padding:0 4px;">
    <p style="font-size:11pt;color:#1a1a2e;margin-bottom:4px;">
      Prezado(a) <b>{primeiro_nome}</b>, tudo bem?
    </p>
    {tipo_tag}
    <p style="font-size:10.5pt;color:#4b5563;margin-top:8px;">
      {corpo_html}
    </p>
  </td></tr>

  <!-- Dados da Ordem (Compra) -->
  <tr><td style="padding:22px 0 8px 0;">
    <table cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background-color:#00b876;width:4px;border-radius:4px;">&nbsp;</td>
      <td style="padding-left:12px;">
        <span style="font-size:12.5pt;color:#004d33;font-weight:bold;letter-spacing:0.3px;">Dados da Ordem</span>
      </td>
    </tr></table>
  </td></tr>
  <tr><td style="padding:0 4px;">
    {ordem_table}
  </td></tr>
  {"" if not venda_table else f"""
  <!-- Dados da Venda (Saída) -->
  <tr><td style="padding:18px 0 8px 0;">
    <table cellpadding='0' cellspacing='0' border='0'><tr>
      <td style='background-color:#e6832a;width:4px;border-radius:4px;'>&nbsp;</td>
      <td style='padding-left:12px;'>
        <span style='font-size:12.5pt;color:#b85c00;font-weight:bold;letter-spacing:0.3px;'>Ativo de Sa&iacute;da (Venda)</span>
      </td>
    </tr></table>
  </td></tr>
  <tr><td style='padding:0 4px;'>
    {venda_table}
  </td></tr>"""}
  <tr><td style="padding:28px 0 0 0;">
    <hr style="border:none;border-top:2px solid #004d33;margin:0 0 12px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="vertical-align:middle;">
          <span style="font-size:10pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:300;"> | </span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:bold;">Produtos</span>
        </td>
        <td style="text-align:right;">
          <span style="font-size:8.5pt;color:#6b7280;">Envio de Ordens &middot; {tipo} &middot; {hoje}</span><br>
          <span style="font-size:7.5pt;color:#9ca3af;">E-mail gerado automaticamente</span>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</div>
"""

            criados = 0
            erros = 0

            if modo == "Grupo":
                # Um único email com todos em To
                try:
                    primeiro = "Prezados"
                    corpo_html = corpo.replace("{nome}", primeiro).replace("\n", "<br>")
                    mail = outlook.CreateItem(0)
                    mail.To = "; ".join(d["email"] for d in self._eo_destinatarios)
                    mail.Subject = assunto
                    mail.HTMLBody = _build_html(primeiro, corpo_html)
                    _attach_logo_cid(mail)
                    if self._eo_anexo_path and os.path.exists(self._eo_anexo_path):
                        mail.Attachments.Add(os.path.abspath(self._eo_anexo_path))
                    mail.Save()
                    criados = 1
                except Exception:
                    erros = 1
            else:
                # Individual
                for d in self._eo_destinatarios:
                    primeiro = _get_primeiro(d)
                    corpo_html = corpo.replace("{nome}", primeiro).replace("\n", "<br>")
                    try:
                        mail = outlook.CreateItem(0)
                        mail.To = d["email"]
                        mail.Subject = assunto
                        mail.HTMLBody = _build_html(primeiro, corpo_html)
                        _attach_logo_cid(mail)
                        if self._eo_anexo_path and os.path.exists(self._eo_anexo_path):
                            mail.Attachments.Add(os.path.abspath(self._eo_anexo_path))
                        mail.Save()
                        criados += 1
                    except Exception:
                        erros += 1

            pythoncom.CoUninitialize()

            def _done():
                self.eo_enviar_btn.configure(state="normal")
                self.eo_status_dot.configure(text_color="#00a86b")
                self.eo_status_text.configure(
                    text=f"  Concluído - {criados} rascunho(s), {erros} erro(s)"
                )
                messagebox.showinfo("Rascunhos Criados",
                    f"{criados} rascunho(s) criado(s) no Outlook!\n{erros} erro(s).")
            self.after(0, _done)

        except Exception as e:
            def _err():
                self.eo_enviar_btn.configure(state="normal")
                self.eo_status_dot.configure(text_color=ACCENT_RED)
                self.eo_status_text.configure(text=f"  Erro: {e}")
                messagebox.showerror("Erro", str(e))
            self.after(0, _err)

    # -----------------------------------------------------------------
    #  PAGE: ENVIO MESA
    # -----------------------------------------------------------------
    def _build_envio_mesa_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Envio Mesa", subtitle="Mesa de Produtos")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # ======== DROP ZONE ========
        self.em_drop_card = ctk.CTkFrame(
            content, fg_color="#f0f4ff", corner_radius=14,
            border_width=2, border_color=ACCENT_BLUE
        )
        self.em_drop_card.pack(fill="x", pady=(0, 16))
        self.em_drop_card.bind("<Button-1>", lambda e: self._on_em_browse())

        drop_inner = ctk.CTkFrame(self.em_drop_card, fg_color="transparent")
        drop_inner.pack(fill="x", padx=30, pady=36)
        drop_inner.bind("<Button-1>", lambda e: self._on_em_browse())

        self.em_icon_label = ctk.CTkLabel(
            drop_inner, text="\u2709",
            font=("Segoe UI", 44), text_color=ACCENT_BLUE
        )
        self.em_icon_label.pack()
        self.em_icon_label.bind("<Button-1>", lambda e: self._on_em_browse())

        self.em_drop_title = ctk.CTkLabel(
            drop_inner, text="Clique para selecionar a planilha de dados",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY
        )
        self.em_drop_title.pack(pady=(8, 2))
        self.em_drop_title.bind("<Button-1>", lambda e: self._on_em_browse())

        self.em_drop_sub = ctk.CTkLabel(
            drop_inner, text="Detecta automaticamente: assessor, datas, valores (R$), taxas (%) e quantidades",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.em_drop_sub.pack()
        self.em_drop_sub.bind("<Button-1>", lambda e: self._on_em_browse())

        self.em_browse_btn = ctk.CTkButton(
            drop_inner, text="  Selecionar Arquivo",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=40, corner_radius=8, width=200,
            command=self._on_em_browse,
        )
        self.em_browse_btn.pack(pady=(14, 0))

        # ======== FILE INFO BAR ========
        self.em_file_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.em_file_frame.pack(fill="x", pady=(0, 12))
        self.em_file_frame.pack_forget()

        fi = ctk.CTkFrame(self.em_file_frame, fg_color="transparent")
        fi.pack(fill="x", padx=16, pady=10)

        self.em_file_icon = ctk.CTkLabel(
            fi, text="\u25cf", font=("Segoe UI", 12), text_color=ACCENT_GREEN
        )
        self.em_file_icon.pack(side="left")

        self.em_file_name = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 11, "bold"), text_color=TEXT_PRIMARY
        )
        self.em_file_name.pack(side="left", padx=(6, 0))

        self.em_file_info = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 10), text_color=TEXT_SECONDARY
        )
        self.em_file_info.pack(side="left", padx=(12, 0))

        ctk.CTkButton(
            fi, text="Trocar", font=("Segoe UI", 10),
            fg_color=BG_INPUT, hover_color=BORDER_CARD,
            text_color=TEXT_SECONDARY, height=28, corner_radius=6, width=70,
            command=self._on_em_browse,
        ).pack(side="right")

        # ======== PREVIEW ========
        self.em_preview_label = ctk.CTkLabel(
            content, text="Preview dos Dados",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.em_preview_label.pack(fill="x", pady=(4, 8))

        self.em_preview_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=12,
            border_width=1, border_color=BORDER_CARD
        )
        self.em_preview_frame.pack(fill="x", pady=(0, 14))

        self.em_preview_text = ctk.CTkTextbox(
            self.em_preview_frame, font=("Consolas", 9),
            fg_color=BG_CARD, text_color=TEXT_PRIMARY,
            corner_radius=10, height=200, wrap="none"
        )
        self.em_preview_text.pack(fill="x", padx=4, pady=4)
        self.em_preview_text.insert("1.0", "  Nenhum arquivo carregado...")
        self.em_preview_text.configure(state="disabled")

        # ======== CONFIGURAÇÕES ========
        ctk.CTkLabel(
            content, text="Configuracoes do Envio",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 8))

        config_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=BORDER_CARD)
        config_card.pack(fill="x", pady=(0, 16))

        ci = ctk.CTkFrame(config_card, fg_color="transparent")
        ci.pack(fill="x", padx=20, pady=16)

        # Assunto
        ctk.CTkLabel(
            ci, text="Assunto do E-mail",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.em_assunto = ctk.CTkEntry(
            ci, font=("Segoe UI", 12), height=40,
            corner_radius=8, border_color=BORDER_CARD,
            placeholder_text="Ex: Sugestao de troca CDI11 - JURO11"
        )
        self.em_assunto.pack(fill="x", pady=(0, 14))

        # Corpo
        ctk.CTkLabel(
            ci, text="Corpo do E-mail (texto antes da tabela)",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.em_corpo = ctk.CTkTextbox(
            ci, font=("Segoe UI", 11), height=120,
            corner_radius=8, border_width=1, border_color=BORDER_CARD,
            fg_color=BG_PRIMARY, text_color=TEXT_PRIMARY, wrap="word"
        )
        self.em_corpo.pack(fill="x", pady=(0, 14))

        # Remetente
        ctk.CTkLabel(
            ci, text="Enviar em nome de",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        self.em_remetente = ctk.CTkEntry(
            ci, font=("Segoe UI", 12), height=40,
            corner_radius=8, border_color=BORDER_CARD,
        )
        self.em_remetente.pack(fill="x", pady=(0, 0))
        self.em_remetente.insert(0, "oportunidades@somuscapital.com.br")

        # ======== ACTION BUTTONS ========
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(12, 12))

        self.em_process_btn = ctk.CTkButton(
            btn_frame, text="\u2191  Gerar Rascunhos no Outlook",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=self._darken(ACCENT_GREEN),
            height=44, corner_radius=10, state="disabled",
            command=self._on_em_gerar,
        )
        self.em_process_btn.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame, text="Limpar Tudo",
            font=("Segoe UI", 11),
            fg_color="#6b7280", hover_color="#555d6a",
            height=44, corner_radius=10, width=120,
            command=self._on_em_limpar,
        ).pack(side="left")

        # ======== STATUS BAR ========
        self.em_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.em_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.em_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.em_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.em_status_dot.pack(side="left")

        self.em_status_text = ctk.CTkLabel(
            si, text="  Aguardando arquivo...",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.em_status_text.pack(side="left")

        # Internal state
        self._em_input_path = None
        self._em_headers = []
        self._em_data_rows = []
        self._em_assessor_groups = {}

        return page

    # -----------------------------------------------------------------
    #  ENVIO MESA: Ações
    # -----------------------------------------------------------------
    def _on_em_browse(self):
        path = filedialog.askopenfilename(
            title="Selecionar Planilha de Dados",
            initialdir=os.path.join(BASE_DIR, "Mesa Produtos", "Envio Mesa"),
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self._em_load_file(path)

    def _em_load_file(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]

            headers = []
            data_rows = []

            # Ler primeira linha para detectar headers
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if first_row:
                headers = [str(c).strip() if c else f"Col{i+1}" for i, c in enumerate(first_row)]

            # Se a primeira célula parece ser um código de assessor (ex: A41072), não tem header
            has_header = False
            if first_row and first_row[0]:
                val = str(first_row[0]).strip()
                # Se começa com letra e contem texto descritivo (header)
                if any(kw in val.lower() for kw in ["codigo", "código", "assessor", "cod"]):
                    has_header = True

            start_row = 2 if has_header else 1
            if not has_header:
                # Sem header, gerar nomes genéricos
                ncols = len(first_row) if first_row else 0
                headers = ["Cod Assessor", "Nome Assessor"] + [f"Col{i+1}" for i in range(2, ncols)]

            for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, values_only=True):
                if not row or not row[0]:
                    continue
                data_rows.append(list(row))

            wb.close()

            if not data_rows:
                messagebox.showwarning("Planilha vazia", "Nenhum dado encontrado na planilha.")
                return

            self._em_input_path = path
            self._em_headers = headers
            self._em_data_rows = data_rows

            # Agrupar por assessor (col 0 = código, col 1 = nome)
            groups = {}
            for row in data_rows:
                code = str(row[0]).strip() if row[0] else ""
                nome = sanitize_text(row[1]) if len(row) > 1 and row[1] else ""
                key = code
                if key not in groups:
                    groups[key] = {"nome": nome, "rows": []}
                groups[key]["rows"].append(row)
            self._em_assessor_groups = groups

            # Update UI
            fname = os.path.basename(path)
            self.em_drop_title.configure(text=fname)
            self.em_drop_sub.configure(text="Arquivo carregado - clique para trocar")
            self.em_icon_label.configure(text_color=ACCENT_GREEN, text="\u2713")
            self.em_drop_card.configure(border_color=ACCENT_GREEN, fg_color="#f0fff4")

            self.em_file_name.configure(text=fname)
            info = f"{len(data_rows)} linhas  |  {len(headers)} colunas  |  {len(groups)} assessores"
            self.em_file_info.configure(text=info)
            self.em_file_frame.pack(fill="x", pady=(0, 12))

            self._em_show_preview()
            self.em_process_btn.configure(state="normal")

            self.em_status_dot.configure(text_color=ACCENT_BLUE)
            self.em_status_text.configure(text=f"  {len(groups)} assessores identificados. Configure e gere os rascunhos.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo:\n{e}")

    def _em_show_preview(self):
        self.em_preview_text.configure(state="normal")
        self.em_preview_text.delete("1.0", "end")

        # Header
        header_line = "  |  ".join(str(h)[:20] for h in self._em_headers)
        self.em_preview_text.insert("end", header_line + "\n")
        self.em_preview_text.insert("end", "-" * len(header_line) + "\n")

        # Data rows (max 20) — formata datas e valores para preview legivel
        def _preview_cell(val):
            if val is None:
                return ""
            if isinstance(val, datetime):
                return val.strftime("%d/%m/%Y")
            if hasattr(val, "strftime") and not isinstance(val, (int, float)):
                return val.strftime("%d/%m/%Y")
            if isinstance(val, float):
                if abs(val) >= 100:
                    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{val:.4f}".replace(".", ",")
            return str(val)[:20]
        for row in self._em_data_rows[:20]:
            cells = [_preview_cell(c) for c in row]
            self.em_preview_text.insert("end", "  |  ".join(cells) + "\n")
        if len(self._em_data_rows) > 20:
            self.em_preview_text.insert("end", f"\n... e mais {len(self._em_data_rows) - 20} linhas")

        self.em_preview_text.configure(state="disabled")

    def _on_em_limpar(self):
        self._em_input_path = None
        self._em_headers = []
        self._em_data_rows = []
        self._em_assessor_groups = {}
        self.em_icon_label.configure(text_color=ACCENT_BLUE, text="\u2709")
        self.em_drop_title.configure(text="Clique para selecionar a planilha de dados")
        self.em_drop_sub.configure(text="Detecta automaticamente: assessor, datas, valores (R$), taxas (%) e quantidades")
        self.em_drop_card.configure(border_color=ACCENT_BLUE, fg_color="#f0f4ff")
        self.em_file_frame.pack_forget()
        self.em_preview_text.configure(state="normal")
        self.em_preview_text.delete("1.0", "end")
        self.em_preview_text.insert("1.0", "  Nenhum arquivo carregado...")
        self.em_preview_text.configure(state="disabled")
        self.em_assunto.delete(0, "end")
        self.em_corpo.delete("1.0", "end")
        self.em_process_btn.configure(state="disabled")
        self.em_status_dot.configure(text_color=TEXT_TERTIARY)
        self.em_status_text.configure(text="  Aguardando arquivo...")

    def _on_em_gerar(self):
        if not self._em_assessor_groups:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return

        assunto = self.em_assunto.get().strip()
        corpo = self.em_corpo.get("1.0", "end").strip()

        if not assunto:
            messagebox.showwarning("Campo obrigatorio", "Preencha o assunto do e-mail.")
            return

        n = len(self._em_assessor_groups)
        msg = (f"Serao criados ate {n} rascunhos (1 por assessor).\n\n"
               f"Assunto: {assunto}\n"
               f"Outlook precisa estar aberto. Continuar?")

        if not messagebox.askyesno("Gerar Rascunhos", msg):
            return

        self.em_process_btn.configure(state="disabled")
        self.em_status_dot.configure(text_color=ACCENT_ORANGE)
        self.em_status_text.configure(text="  Gerando rascunhos...")

        # Limpa preview para usar como log
        self.em_preview_text.configure(state="normal")
        self.em_preview_text.delete("1.0", "end")
        self.em_preview_text.configure(state="disabled")

        remetente = self.em_remetente.get().strip()
        threading.Thread(
            target=self._run_em_emails,
            args=(assunto, corpo, remetente),
            daemon=True
        ).start()

    def _em_append_log(self, msg):
        self.em_preview_text.configure(state="normal")
        self.em_preview_text.insert("end", msg + "\n")
        self.em_preview_text.see("end")
        self.em_preview_text.configure(state="disabled")

    def _run_em_emails(self, assunto, corpo, remetente):
        try:
            import win32com.client as win32
            import pythoncom
            pythoncom.CoInitialize()

            outlook = win32.Dispatch("Outlook.Application")

            # Carregar base de emails
            self.after(0, self._em_append_log, "Carregando base de emails...")
            emails_map = {}
            if os.path.exists(BASE_FILE):
                wb = openpyxl.load_workbook(BASE_FILE, read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
                    if row[0]:
                        code = str(row[0]).strip()
                        code_a = f"A{code}" if not code.startswith("A") else code
                        code_num = code[1:] if code.startswith("A") else code
                        info = {
                            "nome": sanitize_text(row[1]) if row[1] else "",
                            "email": str(row[2]).strip() if row[2] else "",
                        }
                        emails_map[code] = info
                        emails_map[code_a] = info
                        emails_map[code_num] = info
                wb.close()
                self.after(0, self._em_append_log, f"  {len(emails_map)} registros de email carregados.")
            else:
                self.after(0, self._em_append_log, f"  BASE EMAILS nao encontrado em {BASE_FILE}")

            # CC fixos
            cc_fixos = "artur.brito@somuscapital.com.br;leonardo.dellatorre@somuscapital.com.br"

            hoje = datetime.now().strftime("%d/%m/%Y")
            logo_b64 = ""
            if os.path.exists(LOGO_PATH):
                with open(LOGO_PATH, "rb") as f:
                    logo_b64 = base64.b64encode(f.read()).decode()
            logo_tag = (
                f'<img src="data:image/png;base64,{logo_b64}" width="155" height="38"'
                f' style="vertical-align:middle;margin-right:16px;" alt="Somus Capital">'
            ) if logo_b64 else ""

            corpo_html = corpo.replace("\n", "<br>") if corpo else ""

            # Headers para tabela (excluir col 0=código e col 1=nome assessor)
            table_headers = self._em_headers[2:] if len(self._em_headers) > 2 else self._em_headers

            criados = 0
            sem_email = 0
            erros = 0

            self.after(0, self._em_append_log, f"\nProcessando {len(self._em_assessor_groups)} assessores...")
            self.after(0, self._em_append_log, "-" * 60)

            for code, group in sorted(self._em_assessor_groups.items()):
                nome = group["nome"]
                rows = group["rows"]

                # Buscar email
                email_info = emails_map.get(code, {})
                if not email_info:
                    code_num = code[1:] if code.startswith("A") else code
                    email_info = emails_map.get(code_num, {})
                email_to = email_info.get("email", "")

                if not email_to or "@" not in email_to:
                    self.after(0, self._em_append_log, f"  {code}  {nome}:  Sem email cadastrado")
                    sem_email += 1
                    continue

                primeiro_nome = nome.split()[0].capitalize() if nome else "Prezado(a)"

                # Montar tabela HTML (esconder colunas vazias)
                visible_cols = []
                for ci_idx, h in enumerate(table_headers):
                    has_data = any(
                        r[ci_idx + 2] is not None and str(r[ci_idx + 2]).strip() != ""
                        for r in rows if len(r) > ci_idx + 2
                    )
                    if has_data:
                        visible_cols.append((ci_idx + 2, h))

                table_html = '<table cellpadding="0" cellspacing="0" border="0" '
                table_html += 'style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:9pt;margin-bottom:6px;">\n'

                # Header
                table_html += '<tr>'
                for col_idx, label in visible_cols:
                    table_html += (
                        f'<td style="padding:3px 8px 4px 8px;font-weight:bold;color:#00785a;'
                        f'font-size:8.5pt;border-bottom:1.5px solid #00785a;'
                        f'white-space:nowrap;">{sanitize_text(str(label))}</td>'
                    )
                table_html += '</tr>\n'

                # Detectar tipo de cada coluna visivel pelo header
                _MONEY_KW = ("valor", "posic", "posi\u00e7", "financ", "pago", "total", "preco", "pre\u00e7o", "saldo", "montante", "aporte", "resgate")
                _RATE_KW = ("taxa", "rate", "spread", "cupom", "juros", "yield")
                _DATE_KW = ("data ", "data_", "vencimento", "venc", "emissao", "emiss\u00e3o", "aplicacao", "aplica\u00e7")
                _QTY_KW = ("qtd", "quantidade", "quant", "lote")

                def _col_type(lbl):
                    """Detecta tipo da coluna: 'money', 'rate', 'date', 'qty', 'text'."""
                    lo = str(lbl).lower().strip()
                    if any(k in lo for k in _DATE_KW):
                        return "date"
                    if any(k in lo for k in _RATE_KW):
                        return "rate"
                    if any(k in lo for k in _MONEY_KW):
                        return "money"
                    if any(k in lo for k in _QTY_KW):
                        return "qty"
                    return "text"

                col_types = {ci: _col_type(lb) for ci, lb in visible_cols}

                def _fmt_cell(val, ctype):
                    """Formata valor de celula de acordo com o tipo da coluna."""
                    if val is None or (isinstance(val, str) and val.strip() == ""):
                        return ("-", "#1a1a2e", "left")
                    # datetime → DD/MM/YYYY
                    if isinstance(val, datetime):
                        return (val.strftime("%d/%m/%Y"), "#1a1a2e", "center")
                    # date sem hora
                    if hasattr(val, "strftime") and not isinstance(val, (int, float)):
                        return (val.strftime("%d/%m/%Y"), "#1a1a2e", "center")
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        if ctype == "rate":
                            formatted = f"{val:.2f}".replace(".", ",") + "%"
                            return (formatted, "#1a1a2e", "right")
                        if ctype == "money":
                            color = "#c0392b" if val < 0 else "#1a1a2e"
                            neg = val < 0
                            v = abs(round(val, 2))
                            inteiro = int(v)
                            cents = round((v - inteiro) * 100)
                            if cents >= 100:
                                inteiro += 1
                                cents = 0
                            s_int = f"{inteiro:,}".replace(",", ".")
                            formatted = f"R$ {s_int},{cents:02d}"
                            if neg:
                                formatted = f"-{formatted}"
                            return (formatted, color, "right")
                        if ctype == "qty":
                            if val == int(val):
                                return (f"{int(val):,}".replace(",", "."), "#1a1a2e", "right")
                            return (f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), "#1a1a2e", "right")
                        # Tipo 'text' ou 'date' com numero — detectar automaticamente
                        if ctype == "date":
                            return (str(val), "#1a1a2e", "center")
                        # Numero generico grande → BRL, pequeno → decimal
                        if abs(val) >= 100:
                            color = "#c0392b" if val < 0 else "#1a1a2e"
                            neg = val < 0
                            v = abs(round(val, 2))
                            inteiro = int(v)
                            cents = round((v - inteiro) * 100)
                            if cents >= 100:
                                inteiro += 1
                                cents = 0
                            s_int = f"{inteiro:,}".replace(",", ".")
                            formatted = f"R$ {s_int},{cents:02d}"
                            if neg:
                                formatted = f"-{formatted}"
                            return (formatted, color, "right")
                        return (f"{val:,.4f}".replace(",", "X").replace(".", ",").replace("X", "."), "#1a1a2e", "right")
                    return (sanitize_text(str(val)), "#1a1a2e", "left")

                # Data rows
                for ri, row in enumerate(rows):
                    bg = "#f7faf9" if ri % 2 == 0 else "#ffffff"
                    table_html += f'<tr style="background:{bg};">'
                    for col_idx, label in visible_cols:
                        val = row[col_idx] if len(row) > col_idx else ""
                        ctype = col_types.get(col_idx, "text")
                        formatted, color, align = _fmt_cell(val, ctype)
                        table_html += (
                            f'<td style="padding:2px 8px;text-align:{align};color:{color};'
                            f'border-bottom:1px solid #eef1ef;font-size:8.5pt;'
                            f'white-space:nowrap;">{formatted}</td>'
                        )
                    table_html += '</tr>\n'
                table_html += '</table>\n'

                # Montar HTML do email
                email_html = f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:960px;margin:0 auto;">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td style="padding:16px 0 12px 0;vertical-align:middle;">
      {logo_tag}<!--
      --><span style="font-size:16pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:300;"> | </span><!--
      --><span style="font-size:16pt;color:#004d33;font-weight:bold;">Produtos</span>
    </td>
  </tr>
  <tr>
    <td style="padding:0 0 16px 0;">
      <hr style="border:none;border-top:2px solid #004d33;margin:0;">
    </td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:18px;">
  <tr><td style="padding:0 4px;">
    <p style="font-size:11pt;color:#1a1a2e;margin-bottom:4px;">
      Prezado(a) <b>{primeiro_nome}</b>, tudo bem?
    </p>
    {f'<p style="font-size:10.5pt;color:#4b5563;margin-top:8px;">{corpo_html}</p>' if corpo_html else ''}
  </td></tr>
  <tr><td style="padding:12px 4px 0 4px;">
    {table_html}
  </td></tr>
  <tr><td style="padding:16px 4px 0 4px;">
    <p style="font-size:10.5pt;color:#1a1a2e;">Qualquer d&uacute;vida estou &agrave; disposi&ccedil;&atilde;o!</p>
    <p style="font-size:10.5pt;color:#1a1a2e;">Atenciosamente,</p>
  </td></tr>
  <tr><td style="padding:28px 0 0 0;">
    <hr style="border:none;border-top:2px solid #004d33;margin:0 0 12px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="vertical-align:middle;">
          <span style="font-size:10pt;font-weight:bold;color:#004d33;">Somus Capital</span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:300;"> | </span><!--
          --><span style="font-size:10pt;color:#004d33;font-weight:bold;">Produtos</span>
        </td>
        <td style="text-align:right;">
          <span style="font-size:8.5pt;color:#6b7280;">Envio Mesa &middot; {hoje}</span><br>
          <span style="font-size:7.5pt;color:#9ca3af;">E-mail gerado automaticamente</span>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</div>
"""

                try:
                    mail = outlook.CreateItem(0)
                    if remetente:
                        mail.SentOnBehalfOfName = remetente
                    mail.To = email_to
                    mail.CC = cc_fixos
                    mail.Subject = assunto
                    mail.HTMLBody = email_html
                    mail.Save()
                    criados += 1
                    self.after(0, self._em_append_log,
                        f"  {code}  {nome}  ->  {email_to}  [{len(rows)} linhas]")
                except Exception as e:
                    erros += 1
                    self.after(0, self._em_append_log,
                        f"  {code}  {nome}:  ERRO - {e}")

            pythoncom.CoUninitialize()

            self.after(0, self._em_append_log, "-" * 60)
            self.after(0, self._em_append_log,
                f"\nFinalizado!  {criados} rascunhos, {sem_email} sem email, {erros} erros.")

            def _done():
                self.em_process_btn.configure(state="normal")
                self.em_status_dot.configure(text_color=ACCENT_GREEN)
                self.em_status_text.configure(
                    text=f"  Concluido - {criados} rascunho(s), {sem_email} sem email, {erros} erro(s)"
                )
                messagebox.showinfo("Rascunhos Criados",
                    f"{criados} rascunho(s) criado(s) no Outlook!\n"
                    f"{sem_email} assessor(es) sem email\n{erros} erro(s).")
            self.after(0, _done)

        except Exception as e:
            def _err(msg=str(e)):
                self.em_process_btn.configure(state="normal")
                self.em_status_dot.configure(text_color=ACCENT_RED)
                self.em_status_text.configure(text=f"  Erro: {msg}")
                messagebox.showerror("Erro", msg)
            self.after(0, _err)

    # -----------------------------------------------------------------
    #  PAGE: TAREFAS (Organizador de Tarefas)
    # -----------------------------------------------------------------
    _TAREFAS_FILE = os.path.join(BASE_DIR, "BASE", "tarefas.json")

    _PRIORIDADE_CORES = {
        "Alta": "#dc3545",
        "Media": "#e6832a",
        "Baixa": "#1863DC",
    }
    _STATUS_CORES = {
        "Pendente": "#6b7280",
        "Em Andamento": "#e6832a",
        "Concluida": "#00a86b",
    }

    def _load_tarefas(self):
        import json
        if os.path.exists(self._TAREFAS_FILE):
            try:
                with open(self._TAREFAS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return []

    def _save_tarefas(self, tarefas):
        import json
        os.makedirs(os.path.dirname(self._TAREFAS_FILE), exist_ok=True)
        with open(self._TAREFAS_FILE, "w", encoding="utf-8") as f:
            json.dump(tarefas, f, ensure_ascii=False, indent=2)

    # -----------------------------------------------------------------
    #  PAGE: ENVIO ANIVERSÁRIOS
    # -----------------------------------------------------------------
    def _build_envio_aniversarios_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Envio Aniversários", subtitle="Mesa de Produtos")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Placeholder card
        card = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=14,
            border_width=1, border_color=BORDER_CARD
        )
        card.pack(fill="x", pady=(0, 16))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=30, pady=36)

        ctk.CTkLabel(
            inner, text="\u2605",
            font=("Segoe UI", 44), text_color=ACCENT_GREEN
        ).pack()

        ctk.CTkLabel(
            inner, text="Envio de Aniversários",
            font=("Segoe UI", 20, "bold"), text_color=TEXT_PRIMARY
        ).pack(pady=(10, 4))

        ctk.CTkLabel(
            inner, text="Módulo em construção — funcionalidade será implementada em breve.",
            font=("Segoe UI", 13), text_color=TEXT_SECONDARY
        ).pack()

        return page

    def _build_tarefas_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Tarefas", subtitle="Organizador")

        # ---- Toolbar ----
        toolbar = ctk.CTkFrame(page, fg_color=BG_CARD, height=56, corner_radius=0,
                               border_width=0)
        toolbar.pack(fill="x", padx=0, pady=0)
        toolbar.pack_propagate(False)

        tb_inner = ctk.CTkFrame(toolbar, fg_color="transparent")
        tb_inner.pack(fill="x", padx=20, pady=10)

        ctk.CTkButton(
            tb_inner, text="+  Nova Tarefa",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            height=36, corner_radius=8, width=150,
            command=self._on_tf_nova,
        ).pack(side="left", padx=(0, 12))

        # Filtro status
        ctk.CTkLabel(
            tb_inner, text="Filtro:",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        ).pack(side="left", padx=(12, 4))

        self.tf_filtro_status = ctk.CTkSegmentedButton(
            tb_inner,
            values=["Todas", "Pendente", "Em Andamento", "Concluida"],
            font=("Segoe UI", 10, "bold"),
            fg_color=BG_INPUT,
            selected_color=ACCENT_GREEN,
            selected_hover_color=BG_SIDEBAR_HOVER,
            unselected_color=BG_INPUT,
            unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY,
            corner_radius=8,
            height=32,
            command=lambda v: self._tf_refresh(),
        )
        self.tf_filtro_status.pack(side="left", padx=(0, 12))
        self.tf_filtro_status.set("Todas")

        # Filtro prioridade
        self.tf_filtro_prio = ctk.CTkSegmentedButton(
            tb_inner,
            values=["Todas", "Alta", "Media", "Baixa"],
            font=("Segoe UI", 10, "bold"),
            fg_color=BG_INPUT,
            selected_color=ACCENT_BLUE,
            selected_hover_color="#1555bb",
            unselected_color=BG_INPUT,
            unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY,
            corner_radius=8,
            height=32,
            command=lambda v: self._tf_refresh(),
        )
        self.tf_filtro_prio.pack(side="left", padx=(0, 12))
        self.tf_filtro_prio.set("Todas")

        # Contador
        self.tf_counter = ctk.CTkLabel(
            tb_inner, text="0 tarefas",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY
        )
        self.tf_counter.pack(side="right")

        # ---- KPI cards ----
        kpi_frame = ctk.CTkFrame(page, fg_color="transparent")
        kpi_frame.pack(fill="x", padx=20, pady=(12, 6))
        kpi_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.tf_kpi_total = self._tf_make_kpi(kpi_frame, "Total", "0", TEXT_PRIMARY, 0)
        self.tf_kpi_pend = self._tf_make_kpi(kpi_frame, "Pendentes", "0", "#6b7280", 1)
        self.tf_kpi_andamento = self._tf_make_kpi(kpi_frame, "Em Andamento", "0", ACCENT_ORANGE, 2)
        self.tf_kpi_concl = self._tf_make_kpi(kpi_frame, "Concluidas", "0", ACCENT_GREEN, 3)

        # ---- Scroll area com tarefas ----
        self.tf_scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        self.tf_scroll.pack(fill="both", expand=True, padx=0, pady=0)

        self.tf_list_frame = ctk.CTkFrame(self.tf_scroll, fg_color="transparent")
        self.tf_list_frame.pack(fill="x", padx=20, pady=10)

        # Carregar e renderizar
        self._tf_tarefas = self._load_tarefas()
        self._tf_refresh()

        return page

    def _tf_make_kpi(self, parent, label, value, color, col):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=10,
                            border_width=1, border_color=BORDER_CARD, height=70)
        card.grid(row=0, column=col, sticky="nsew", padx=4, pady=0)
        card.grid_propagate(False)

        ctk.CTkLabel(
            card, text=label,
            font=("Segoe UI", 9, "bold"), text_color=TEXT_TERTIARY
        ).place(relx=0.5, rely=0.3, anchor="center")

        val_label = ctk.CTkLabel(
            card, text=value,
            font=("Segoe UI", 20, "bold"), text_color=color
        )
        val_label.place(relx=0.5, rely=0.68, anchor="center")

        return val_label

    def _tf_refresh(self):
        """Reconstroi a lista de tarefas na UI."""
        # Limpar
        for w in self.tf_list_frame.winfo_children():
            w.destroy()

        filtro_st = self.tf_filtro_status.get()
        filtro_pr = self.tf_filtro_prio.get()

        tarefas = self._tf_tarefas

        # KPIs
        total = len(tarefas)
        pend = sum(1 for t in tarefas if t.get("status") == "Pendente")
        andamento = sum(1 for t in tarefas if t.get("status") == "Em Andamento")
        concl = sum(1 for t in tarefas if t.get("status") == "Concluida")
        self.tf_kpi_total.configure(text=str(total))
        self.tf_kpi_pend.configure(text=str(pend))
        self.tf_kpi_andamento.configure(text=str(andamento))
        self.tf_kpi_concl.configure(text=str(concl))

        # Filtrar
        filtered = tarefas
        if filtro_st != "Todas":
            filtered = [t for t in filtered if t.get("status") == filtro_st]
        if filtro_pr != "Todas":
            filtered = [t for t in filtered if t.get("prioridade") == filtro_pr]

        # Ordenar: Pendente > Em Andamento > Concluida, depois Alta > Media > Baixa
        _st_order = {"Pendente": 0, "Em Andamento": 1, "Concluida": 2}
        _pr_order = {"Alta": 0, "Media": 1, "Baixa": 2}
        filtered.sort(key=lambda t: (
            _st_order.get(t.get("status", ""), 9),
            _pr_order.get(t.get("prioridade", ""), 9),
        ))

        self.tf_counter.configure(text=f"{len(filtered)} tarefa(s)")

        if not filtered:
            ctk.CTkLabel(
                self.tf_list_frame,
                text="Nenhuma tarefa encontrada.\nClique em '+ Nova Tarefa' para criar.",
                font=("Segoe UI", 13), text_color=TEXT_TERTIARY,
                justify="center"
            ).pack(pady=40)
            return

        for idx, tarefa in enumerate(filtered):
            self._tf_render_card(tarefa, idx)

    def _tf_render_card(self, tarefa, idx):
        """Renderiza um card de tarefa."""
        tid = tarefa.get("id", "")
        status = tarefa.get("status", "Pendente")
        prio = tarefa.get("prioridade", "Media")
        titulo = tarefa.get("titulo", "Sem titulo")
        descricao = tarefa.get("descricao", "")
        responsavel = tarefa.get("responsavel", "")
        prazo = tarefa.get("prazo", "")
        categoria = tarefa.get("categoria", "")

        is_done = status == "Concluida"

        # Card com borda esquerda colorida pela prioridade
        prio_color = self._PRIORIDADE_CORES.get(prio, ACCENT_BLUE)
        st_color = self._STATUS_CORES.get(status, "#6b7280")

        card = ctk.CTkFrame(
            self.tf_list_frame, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD if not is_done else "#d1fae5"
        )
        card.pack(fill="x", pady=(0, 6))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)

        # Row 1: Prioridade pill + Título + Status pill + Ações
        row1 = ctk.CTkFrame(inner, fg_color="transparent")
        row1.pack(fill="x")

        # Barra prioridade
        prio_bar = ctk.CTkFrame(row1, fg_color=prio_color, width=4, corner_radius=2)
        prio_bar.pack(side="left", fill="y", padx=(0, 10), pady=2)

        # Prioridade pill
        ctk.CTkLabel(
            row1, text=f" {prio} ",
            font=("Segoe UI", 8, "bold"), text_color="white",
            fg_color=prio_color, corner_radius=4, height=18
        ).pack(side="left", padx=(0, 8))

        # Titulo
        title_font = ("Segoe UI", 13, "bold") if not is_done else ("Segoe UI", 13)
        title_color = TEXT_PRIMARY if not is_done else TEXT_TERTIARY
        ctk.CTkLabel(
            row1, text=titulo,
            font=title_font, text_color=title_color, anchor="w"
        ).pack(side="left", fill="x", expand=True)

        # Status pill
        ctk.CTkLabel(
            row1, text=f" {status} ",
            font=("Segoe UI", 8, "bold"), text_color="white",
            fg_color=st_color, corner_radius=4, height=18
        ).pack(side="right", padx=(8, 0))

        # Botões de ação
        btn_frame = ctk.CTkFrame(row1, fg_color="transparent")
        btn_frame.pack(side="right")

        if status != "Concluida":
            next_status = "Em Andamento" if status == "Pendente" else "Concluida"
            next_icon = "\u25b6" if status == "Pendente" else "\u2713"
            ctk.CTkButton(
                btn_frame, text=next_icon, width=28, height=28,
                font=("Segoe UI", 12), fg_color=BG_INPUT,
                hover_color=BORDER_CARD, text_color=st_color,
                corner_radius=6,
                command=lambda t=tarefa, ns=next_status: self._on_tf_change_status(t, ns),
            ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_frame, text="\u270e", width=28, height=28,
            font=("Segoe UI", 12), fg_color=BG_INPUT,
            hover_color=BORDER_CARD, text_color=ACCENT_BLUE,
            corner_radius=6,
            command=lambda t=tarefa: self._on_tf_editar(t),
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            btn_frame, text="\u2715", width=28, height=28,
            font=("Segoe UI", 12), fg_color=BG_INPUT,
            hover_color="#fde8e8", text_color=ACCENT_RED,
            corner_radius=6,
            command=lambda t=tarefa: self._on_tf_excluir(t),
        ).pack(side="left", padx=2)

        # Row 2: Meta info
        if descricao or responsavel or prazo or categoria:
            row2 = ctk.CTkFrame(inner, fg_color="transparent")
            row2.pack(fill="x", pady=(6, 0), padx=(14, 0))

            if descricao:
                desc_text = descricao if len(descricao) <= 120 else descricao[:120] + "..."
                ctk.CTkLabel(
                    row2, text=desc_text,
                    font=("Segoe UI", 10), text_color=TEXT_SECONDARY, anchor="w",
                    wraplength=600
                ).pack(fill="x", pady=(0, 4))

            meta_row = ctk.CTkFrame(row2, fg_color="transparent")
            meta_row.pack(fill="x")

            if responsavel:
                ctk.CTkLabel(
                    meta_row, text=f"\u2022 {responsavel}",
                    font=("Segoe UI", 9), text_color=TEXT_TERTIARY
                ).pack(side="left", padx=(0, 12))

            if categoria:
                ctk.CTkLabel(
                    meta_row, text=f"\u2022 {categoria}",
                    font=("Segoe UI", 9), text_color=TEXT_TERTIARY
                ).pack(side="left", padx=(0, 12))

            if prazo:
                prazo_color = TEXT_TERTIARY
                # Highlight se vencida
                try:
                    from datetime import datetime as _dt
                    dt_prazo = _dt.strptime(prazo, "%d/%m/%Y")
                    if dt_prazo.date() < _dt.now().date() and not is_done:
                        prazo_color = ACCENT_RED
                except Exception:
                    pass
                ctk.CTkLabel(
                    meta_row, text=f"\u2022 Prazo: {prazo}",
                    font=("Segoe UI", 9, "bold"), text_color=prazo_color
                ).pack(side="left", padx=(0, 12))

    def _on_tf_nova(self):
        self._tf_open_dialog()

    def _on_tf_editar(self, tarefa):
        self._tf_open_dialog(tarefa)

    def _tf_open_dialog(self, tarefa=None):
        """Abre dialog para criar/editar tarefa."""
        is_edit = tarefa is not None

        dialog = ctk.CTkToplevel(self)
        dialog.title("Editar Tarefa" if is_edit else "Nova Tarefa")
        dialog.geometry("520x560")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 520) // 2
        y = (dialog.winfo_screenheight() - 560) // 2
        dialog.geometry(f"520x560+{x}+{y}")

        dialog.configure(fg_color=BG_PRIMARY)

        # Header
        hdr = ctk.CTkFrame(dialog, fg_color=ACCENT_GREEN, height=50, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="  Editar Tarefa" if is_edit else "  Nova Tarefa",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_WHITE, anchor="w"
        ).pack(side="left", padx=20)

        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=24, pady=16)

        # Titulo
        ctk.CTkLabel(content, text="Titulo *", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        titulo_entry = ctk.CTkEntry(content, font=("Segoe UI", 12), height=38,
                                    corner_radius=8, border_color=BORDER_CARD)
        titulo_entry.pack(fill="x", pady=(0, 12))
        if is_edit:
            titulo_entry.insert(0, tarefa.get("titulo", ""))

        # Descricao
        ctk.CTkLabel(content, text="Descricao", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        desc_text = ctk.CTkTextbox(content, font=("Segoe UI", 11), height=80,
                                   corner_radius=8, border_width=1, border_color=BORDER_CARD,
                                   fg_color=BG_CARD, text_color=TEXT_PRIMARY, wrap="word")
        desc_text.pack(fill="x", pady=(0, 12))
        if is_edit and tarefa.get("descricao"):
            desc_text.insert("1.0", tarefa["descricao"])

        # Row: Prioridade + Status
        row1 = ctk.CTkFrame(content, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 12))
        row1.columnconfigure((0, 1), weight=1)

        prio_frame = ctk.CTkFrame(row1, fg_color="transparent")
        prio_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(prio_frame, text="Prioridade", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        prio_seg = ctk.CTkSegmentedButton(
            prio_frame, values=["Alta", "Media", "Baixa"],
            font=("Segoe UI", 10, "bold"), fg_color=BG_INPUT,
            selected_color=ACCENT_BLUE, selected_hover_color="#1555bb",
            unselected_color=BG_INPUT, unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY, corner_radius=8, height=34,
        )
        prio_seg.pack(fill="x")
        prio_seg.set(tarefa.get("prioridade", "Media") if is_edit else "Media")

        status_frame = ctk.CTkFrame(row1, fg_color="transparent")
        status_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(status_frame, text="Status", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        status_seg = ctk.CTkSegmentedButton(
            status_frame, values=["Pendente", "Em Andamento", "Concluida"],
            font=("Segoe UI", 10, "bold"), fg_color=BG_INPUT,
            selected_color=ACCENT_GREEN, selected_hover_color=BG_SIDEBAR_HOVER,
            unselected_color=BG_INPUT, unselected_hover_color=BORDER_CARD,
            text_color=TEXT_PRIMARY, corner_radius=8, height=34,
        )
        status_seg.pack(fill="x")
        status_seg.set(tarefa.get("status", "Pendente") if is_edit else "Pendente")

        # Row: Responsavel + Categoria
        row2 = ctk.CTkFrame(content, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 12))
        row2.columnconfigure((0, 1), weight=1)

        resp_frame = ctk.CTkFrame(row2, fg_color="transparent")
        resp_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(resp_frame, text="Responsavel", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        resp_entry = ctk.CTkEntry(resp_frame, font=("Segoe UI", 12), height=38,
                                  corner_radius=8, border_color=BORDER_CARD,
                                  placeholder_text="Nome")
        resp_entry.pack(fill="x")
        if is_edit and tarefa.get("responsavel"):
            resp_entry.insert(0, tarefa["responsavel"])

        cat_frame = ctk.CTkFrame(row2, fg_color="transparent")
        cat_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(cat_frame, text="Categoria", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        cat_entry = ctk.CTkEntry(cat_frame, font=("Segoe UI", 12), height=38,
                                 corner_radius=8, border_color=BORDER_CARD,
                                 placeholder_text="Ex: Produtos, RV, RF")
        cat_entry.pack(fill="x")
        if is_edit and tarefa.get("categoria"):
            cat_entry.insert(0, tarefa["categoria"])

        # Prazo
        ctk.CTkLabel(content, text="Prazo (DD/MM/AAAA)", font=("Segoe UI", 11, "bold"),
                      text_color=TEXT_SECONDARY, anchor="w").pack(fill="x", pady=(0, 4))
        prazo_entry = ctk.CTkEntry(content, font=("Segoe UI", 12), height=38,
                                   corner_radius=8, border_color=BORDER_CARD,
                                   placeholder_text="Ex: 15/03/2026")
        prazo_entry.pack(fill="x", pady=(0, 16))
        if is_edit and tarefa.get("prazo"):
            prazo_entry.insert(0, tarefa["prazo"])

        # Botões
        btn_row = ctk.CTkFrame(content, fg_color="transparent")
        btn_row.pack(fill="x")

        def _salvar():
            titulo = titulo_entry.get().strip()
            if not titulo:
                messagebox.showwarning("Campo obrigatorio", "Preencha o titulo da tarefa.")
                return

            import uuid
            data = {
                "id": tarefa.get("id", str(uuid.uuid4())[:8]) if is_edit else str(uuid.uuid4())[:8],
                "titulo": titulo,
                "descricao": desc_text.get("1.0", "end").strip(),
                "prioridade": prio_seg.get(),
                "status": status_seg.get(),
                "responsavel": resp_entry.get().strip(),
                "categoria": cat_entry.get().strip(),
                "prazo": prazo_entry.get().strip(),
                "criado_em": tarefa.get("criado_em", datetime.now().strftime("%d/%m/%Y %H:%M")) if is_edit else datetime.now().strftime("%d/%m/%Y %H:%M"),
            }

            if is_edit:
                for i, t in enumerate(self._tf_tarefas):
                    if t.get("id") == tarefa.get("id"):
                        self._tf_tarefas[i] = data
                        break
            else:
                self._tf_tarefas.append(data)

            self._save_tarefas(self._tf_tarefas)
            self._tf_refresh()
            dialog.destroy()

        ctk.CTkButton(
            btn_row, text="Salvar",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            height=40, corner_radius=8, width=140,
            command=_salvar,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row, text="Cancelar",
            font=("Segoe UI", 11),
            fg_color="#6b7280", hover_color="#555d6a",
            height=40, corner_radius=8, width=100,
            command=dialog.destroy,
        ).pack(side="left")

    def _on_tf_change_status(self, tarefa, new_status):
        for t in self._tf_tarefas:
            if t.get("id") == tarefa.get("id"):
                t["status"] = new_status
                break
        self._save_tarefas(self._tf_tarefas)
        self._tf_refresh()

    def _on_tf_excluir(self, tarefa):
        if not messagebox.askyesno("Excluir Tarefa",
            f"Excluir '{tarefa.get('titulo', '')}'?\n\nEssa acao nao pode ser desfeita."):
            return
        self._tf_tarefas = [t for t in self._tf_tarefas if t.get("id") != tarefa.get("id")]
        self._save_tarefas(self._tf_tarefas)
        self._tf_refresh()

    # -----------------------------------------------------------------
    #  PAGE: CORPORATE - DASHBOARD
    # -----------------------------------------------------------------
    def _build_corp_dashboard_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Dashboard", subtitle="Corporate")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # Banner
        banner = ctk.CTkFrame(content, fg_color=ACCENT_GREEN, corner_radius=10, height=44)
        banner.pack(fill="x", pady=(0, 18))
        banner.pack_propagate(False)

        ctk.CTkLabel(
            banner, text="  Corporate - Em breve",
            font=("Segoe UI", 12), text_color=TEXT_WHITE, anchor="w"
        ).pack(side="left", padx=18, pady=10)

        # KPI placeholders
        kpi_frame = ctk.CTkFrame(content, fg_color="transparent")
        kpi_frame.pack(fill="x", pady=(0, 18))
        kpi_frame.columnconfigure((0, 1, 2, 3), weight=1)

        placeholders = [
            ("Operações", "--", ACCENT_GREEN),
            ("Clientes", "--", ACCENT_BLUE),
            ("Volume Total", "--", ACCENT_ORANGE),
            ("Receita", "--", ACCENT_PURPLE),
        ]
        for col, (label, val, color) in enumerate(placeholders):
            self._make_kpi_card(kpi_frame, label, val, color, col)

        # Card vazio central
        empty_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                  border_width=1, border_color=BORDER_CARD)
        empty_card.pack(fill="x", pady=(0, 18))

        empty_inner = ctk.CTkFrame(empty_card, fg_color="transparent")
        empty_inner.pack(fill="both", padx=24, pady=40)

        ctk.CTkLabel(
            empty_inner, text="\u25a3",
            font=("Segoe UI", 48), text_color=BORDER_CARD
        ).pack()

        ctk.CTkLabel(
            empty_inner, text="Dashboard Corporate",
            font=("Segoe UI", 18, "bold"), text_color=TEXT_SECONDARY
        ).pack(pady=(8, 4))

        ctk.CTkLabel(
            empty_inner, text="Os dados serão exibidos aqui quando disponíveis.\nEm breve: operações, clientes, volume e receita.",
            font=("Segoe UI", 12), text_color=TEXT_TERTIARY,
            justify="center"
        ).pack()

        return page

    # -----------------------------------------------------------------
    #  PAGE: CORPORATE - SIMULADOR
    # -----------------------------------------------------------------
    def _build_consorcio_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        # Top bar profissional
        self._make_topbar(
            page, "Simulador de Consórcio", subtitle="Corporate"
        )

        # Scroll area
        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # ========== DADOS DO CLIENTE ==========
        ctk.CTkLabel(
            content, text="Dados do Cliente",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 8))

        client_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=BORDER_CARD)
        client_card.pack(fill="x", pady=(0, 16))

        ci = ctk.CTkFrame(client_card, fg_color="transparent")
        ci.pack(fill="x", padx=20, pady=16)
        ci.columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(ci, text="Nome do Cliente", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.cs_nome = ctk.CTkEntry(ci, placeholder_text="Nome completo do cliente",
                                    height=38, corner_radius=8, fg_color=BG_INPUT,
                                    border_width=1, border_color=BORDER_LIGHT)
        self.cs_nome.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(2, 0))

        ctk.CTkLabel(ci, text="Assessor Responsavel", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.cs_assessor = ctk.CTkEntry(ci, placeholder_text="Nome do assessor",
                                        height=38, corner_radius=8, fg_color=BG_INPUT,
                                        border_width=1, border_color=BORDER_LIGHT)
        self.cs_assessor.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(2, 0))

        # ========== PARAMETROS DO CONSÓRCIO ==========
        ctk.CTkLabel(
            content, text="Parâmetros do Consórcio",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 8))

        params_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                   border_width=1, border_color=BORDER_CARD)
        params_card.pack(fill="x", pady=(0, 16))

        pi = ctk.CTkFrame(params_card, fg_color="transparent")
        pi.pack(fill="x", padx=20, pady=16)
        pi.columnconfigure((0, 1, 2), weight=1)

        # Row 0-1: Tipo + Administradora + Prazo
        ctk.CTkLabel(pi, text="Tipo do Bem", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.cs_tipo_bem = ctk.CTkOptionMenu(
            pi, values=["Imovel", "Automovel", "Moto", "Servicos", "Outros"],
            height=38, corner_radius=8, fg_color=BG_INPUT,
            button_color=ACCENT_GREEN, button_hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY, font=("Segoe UI", 11),
        )
        self.cs_tipo_bem.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(2, 8))

        ctk.CTkLabel(pi, text="Administradora", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=1, sticky="w", padx=(10, 10))
        self.cs_administradora = ctk.CTkEntry(
            pi, placeholder_text="Ex: Porto Seguro, Embracon...",
            height=38, corner_radius=8, fg_color=BG_INPUT,
            border_width=1, border_color=BORDER_LIGHT)
        self.cs_administradora.grid(row=1, column=1, sticky="ew", padx=(10, 10), pady=(2, 8))

        ctk.CTkLabel(pi, text="Prazo (meses)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.cs_prazo = ctk.CTkOptionMenu(
            pi, values=["36", "48", "60", "72", "80", "100", "120", "150", "180", "200", "240"],
            height=38, corner_radius=8, fg_color=BG_INPUT,
            button_color=ACCENT_GREEN, button_hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY, font=("Segoe UI", 11),
        )
        self.cs_prazo.set("120")
        self.cs_prazo.grid(row=1, column=2, sticky="ew", padx=(10, 0), pady=(2, 8))

        # Row 2-3: Valor da Carta (full width)
        ctk.CTkLabel(pi, text="Valor da Carta de Crédito (R$)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.cs_valor_carta = ctk.CTkEntry(
            pi, placeholder_text="Ex: 300000", height=42, corner_radius=8,
            fg_color=BG_INPUT, border_width=1, border_color=BORDER_LIGHT,
            font=("Segoe UI", 14))
        self.cs_valor_carta.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(2, 8))

        # Row 4-5: Taxas
        ctk.CTkLabel(pi, text="Taxa de Administração (%)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=4, column=0, sticky="w", padx=(0, 10))
        self.cs_taxa_adm = ctk.CTkEntry(pi, placeholder_text="Ex: 18", height=38,
                                        corner_radius=8, fg_color=BG_INPUT,
                                        border_width=1, border_color=BORDER_LIGHT)
        self.cs_taxa_adm.grid(row=5, column=0, sticky="ew", padx=(0, 10), pady=(2, 0))
        self.cs_taxa_adm.insert(0, "18")

        ctk.CTkLabel(pi, text="Fundo de Reserva (%)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=4, column=1, sticky="w", padx=(10, 10))
        self.cs_fundo_reserva = ctk.CTkEntry(pi, placeholder_text="Ex: 2", height=38,
                                             corner_radius=8, fg_color=BG_INPUT,
                                             border_width=1, border_color=BORDER_LIGHT)
        self.cs_fundo_reserva.grid(row=5, column=1, sticky="ew", padx=(10, 10), pady=(2, 0))
        self.cs_fundo_reserva.insert(0, "2")

        ctk.CTkLabel(pi, text="Seguro Prestamista (%)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=4, column=2, sticky="w", padx=(10, 0))
        self.cs_seguro = ctk.CTkEntry(pi, placeholder_text="Ex: 0", height=38,
                                      corner_radius=8, fg_color=BG_INPUT,
                                      border_width=1, border_color=BORDER_LIGHT)
        self.cs_seguro.grid(row=5, column=2, sticky="ew", padx=(10, 0), pady=(2, 0))
        self.cs_seguro.insert(0, "0")

        # Row 6-7: Correção anual + Tipo
        ctk.CTkLabel(pi, text="Correção Anual (% a.a.)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=6, column=0, sticky="w", padx=(0, 10), pady=(6, 0))
        self.cs_correção_anual = ctk.CTkEntry(pi, placeholder_text="Ex: 5.5", height=38,
                                              corner_radius=8, fg_color=BG_INPUT,
                                              border_width=1, border_color=BORDER_LIGHT)
        self.cs_correção_anual.grid(row=7, column=0, sticky="ew", padx=(0, 10), pady=(2, 2))
        self.cs_correção_anual.insert(0, "0")

        ctk.CTkLabel(pi, text="Tipo de Correção", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=6, column=1, sticky="w", padx=(10, 10), pady=(6, 0))
        self.cs_tipo_correção = ctk.CTkSegmentedButton(
            pi, values=["Pre-fixado", "Pós-fixado"],
            font=("Segoe UI", 10, "bold"),
            fg_color=BG_INPUT,
            selected_color=ACCENT_BLUE,
            selected_hover_color="#1555bb",
            unselected_color="#e0e3e8",
            unselected_hover_color="#d0d3d8",
            text_color=TEXT_PRIMARY,
            corner_radius=8, height=38,
        )
        self.cs_tipo_correção.set("Pós-fixado")
        self.cs_tipo_correção.grid(row=7, column=1, sticky="ew", padx=(10, 10), pady=(2, 2))

        ctk.CTkLabel(pi, text="Indice (INCC, IPCA...)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=6, column=2, sticky="w", padx=(10, 0), pady=(6, 0))
        self.cs_índice_correção = ctk.CTkOptionMenu(
            pi, values=["INCC", "IPCA", "IGP-M", "Outro"],
            height=38, corner_radius=8, fg_color=BG_INPUT,
            button_color=ACCENT_GREEN, button_hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY, font=("Segoe UI", 11),
        )
        self.cs_índice_correção.grid(row=7, column=2, sticky="ew", padx=(10, 0), pady=(2, 2))

        ctk.CTkLabel(pi, text="INCC p/ imovel, IPCA p/ veiculo. 0 = sem correção",
                     font=("Segoe UI", 8), text_color=TEXT_TERTIARY
                     ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(0, 0))

        # ========== CONTEMPLAÇÃO E LANCES ==========
        ctk.CTkLabel(
            content, text="Contemplação e Lances",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(8, 8))

        cl_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                               border_width=1, border_color=BORDER_CARD)
        cl_card.pack(fill="x", pady=(0, 16))

        cli = ctk.CTkFrame(cl_card, fg_color="transparent")
        cli.pack(fill="x", padx=20, pady=16)
        cli.columnconfigure((0, 1), weight=1)

        # Row 0-1: Prazo contemplação + Parcela reduzida
        ctk.CTkLabel(cli, text="Prazo de Contemplação (mes)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.cs_prazo_contemplação = ctk.CTkEntry(
            cli, placeholder_text="Ex: 60", height=38, corner_radius=8,
            fg_color=BG_INPUT, border_width=1, border_color=BORDER_LIGHT)
        self.cs_prazo_contemplação.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(2, 2))
        self.cs_prazo_contemplação.insert(0, "60")

        ctk.CTkLabel(cli, text="Parcela Reduzida (% do Fundo Comum)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.cs_parcela_reduzida = ctk.CTkOptionMenu(
            cli, values=["100% (Integral)", "70% (Reducao 30%)", "50% (Plano 50)"],
            height=38, corner_radius=8, fg_color=BG_INPUT,
            button_color=ACCENT_GREEN, button_hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY, font=("Segoe UI", 11),
        )
        self.cs_parcela_reduzida.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(2, 2))

        ctk.CTkLabel(cli, text="Mês estimado para contemplação (sorteio ou lance)",
                     font=("Segoe UI", 8), text_color=TEXT_TERTIARY
                     ).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        ctk.CTkLabel(cli, text="% do fundo comum pago antes da contemplação",
                     font=("Segoe UI", 8), text_color=TEXT_TERTIARY
                     ).grid(row=2, column=1, sticky="w", padx=(10, 0), pady=(0, 8))

        # Row 3-4: Lance livre + Lance embutido
        ctk.CTkLabel(cli, text="Lance Livre Ofertado (% da carta)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=3, column=0, sticky="w", padx=(0, 10))
        self.cs_lance_livre = ctk.CTkEntry(
            cli, placeholder_text="Ex: 20", height=38, corner_radius=8,
            fg_color=BG_INPUT, border_width=1, border_color=BORDER_LIGHT)
        self.cs_lance_livre.grid(row=4, column=0, sticky="ew", padx=(0, 10), pady=(2, 2))
        self.cs_lance_livre.insert(0, "0")

        ctk.CTkLabel(cli, text="Lance Embutido Ofertado (% da carta)", font=("Segoe UI", 10, "bold"),
                     text_color=TEXT_SECONDARY).grid(row=3, column=1, sticky="w", padx=(10, 0))
        self.cs_lance_embutido = ctk.CTkEntry(
            cli, placeholder_text="Ex: 15 (max ~30%)", height=38, corner_radius=8,
            fg_color=BG_INPUT, border_width=1, border_color=BORDER_LIGHT)
        self.cs_lance_embutido.grid(row=4, column=1, sticky="ew", padx=(10, 0), pady=(2, 2))
        self.cs_lance_embutido.insert(0, "0")

        ctk.CTkLabel(cli, text="Recursos próprios para antecipar contemplação",
                     font=("Segoe UI", 8), text_color=TEXT_TERTIARY
                     ).grid(row=5, column=0, sticky="w", padx=(0, 10))
        ctk.CTkLabel(cli, text="Descontado do crédito recebido (max ~25-30%)",
                     font=("Segoe UI", 8), text_color=TEXT_TERTIARY
                     ).grid(row=5, column=1, sticky="w", padx=(10, 0))

        # ========== BOTOES DE ACAO ==========
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(4, 16))

        ctk.CTkButton(
            btn_frame, text="  Calcular Simulação",
            font=("Segoe UI", 13, "bold"), fg_color=ACCENT_GREEN,
            hover_color=self._darken(ACCENT_GREEN), height=44, corner_radius=10,
            command=self._on_calcular_consorcio,
        ).pack(side="left", padx=(0, 10))

        self.cs_btn_pdf = ctk.CTkButton(
            btn_frame, text="  Gerar PDF",
            font=("Segoe UI", 13, "bold"), fg_color=ACCENT_BLUE,
            hover_color=self._darken(ACCENT_BLUE), height=44, corner_radius=10,
            command=self._on_gerar_pdf_consorcio, state="disabled",
        )
        self.cs_btn_pdf.pack(side="left", padx=(10, 10))

        self.cs_btn_relatório = ctk.CTkButton(
            btn_frame, text="  Gerar Relatório",
            font=("Segoe UI", 13, "bold"), fg_color=ACCENT_PURPLE,
            hover_color="#6a4daa", height=44, corner_radius=10,
            command=self._on_gerar_relatorio_consorcio, state="disabled",
        )
        self.cs_btn_relatório.pack(side="left", padx=(0, 10))

        self.cs_btn_email = ctk.CTkButton(
            btn_frame, text="  Enviar Email",
            font=("Segoe UI", 13, "bold"), fg_color=ACCENT_ORANGE,
            hover_color="#c96f1f", height=44, corner_radius=10,
            command=self._on_abrir_email_consorcio, state="disabled",
        )
        self.cs_btn_email.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame, text="  Limpar",
            font=("Segoe UI", 12), fg_color="#6b7280",
            hover_color="#555d6a", height=44, corner_radius=10,
            command=self._on_limpar_consorcio,
        ).pack(side="left", padx=(10, 0))

        # ========== RESULTADO ==========
        ctk.CTkLabel(
            content, text="Resultado da Simulação",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 8))

        self.cs_result_frame = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12,
                                            border_width=1, border_color=BORDER_CARD)
        self.cs_result_frame.pack(fill="x", pady=(0, 16))

        self.cs_result_inner = ctk.CTkFrame(self.cs_result_frame, fg_color="transparent")
        self.cs_result_inner.pack(fill="x", padx=20, pady=16)

        self.cs_result_placeholder = ctk.CTkLabel(
            self.cs_result_inner,
            text="Preencha os campos acima e clique em 'Calcular Simulação'",
            font=("Segoe UI", 12), text_color=TEXT_TERTIARY
        )
        self.cs_result_placeholder.pack(pady=20)

        return page

    # -----------------------------------------------------------------
    #  PAGE: CONSOLIDADOR DE CARTEIRAS
    # -----------------------------------------------------------------
    def _build_consolidador_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Consolidador", subtitle="Carteiras Multi-Instituicao")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # ======== DROP ZONE ========
        self.cons_drop_card = ctk.CTkFrame(
            content, fg_color="#f0f4ff", corner_radius=14,
            border_width=2, border_color=ACCENT_BLUE
        )
        self.cons_drop_card.pack(fill="x", pady=(0, 16))
        self.cons_drop_card.bind("<Button-1>", lambda e: self._on_cons_browse())

        drop_inner = ctk.CTkFrame(self.cons_drop_card, fg_color="transparent")
        drop_inner.pack(fill="x", padx=30, pady=36)
        drop_inner.bind("<Button-1>", lambda e: self._on_cons_browse())

        self.cons_icon_label = ctk.CTkLabel(
            drop_inner, text="\u229a",
            font=("Segoe UI", 44), text_color=ACCENT_BLUE
        )
        self.cons_icon_label.pack()
        self.cons_icon_label.bind("<Button-1>", lambda e: self._on_cons_browse())

        self.cons_drop_title = ctk.CTkLabel(
            drop_inner, text="Clique para selecionar relatorios de carteira",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY
        )
        self.cons_drop_title.pack(pady=(8, 2))
        self.cons_drop_title.bind("<Button-1>", lambda e: self._on_cons_browse())

        self.cons_drop_sub = ctk.CTkLabel(
            drop_inner, text="Formatos aceitos:  PDF (XP, Itau, Safra)  |  XLSX (BTG)",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.cons_drop_sub.pack()
        self.cons_drop_sub.bind("<Button-1>", lambda e: self._on_cons_browse())

        self.cons_browse_btn = ctk.CTkButton(
            drop_inner, text="  Selecionar Arquivos",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=40, corner_radius=8, width=220,
            command=self._on_cons_browse,
        )
        self.cons_browse_btn.pack(pady=(14, 0))

        # ======== FILES LIST (hidden initially) ========
        self.cons_files_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.cons_files_frame.pack(fill="x", pady=(0, 12))
        self.cons_files_frame.pack_forget()

        self.cons_files_inner = ctk.CTkFrame(self.cons_files_frame, fg_color="transparent")
        self.cons_files_inner.pack(fill="x", padx=16, pady=10)

        self.cons_files_label = ctk.CTkLabel(
            self.cons_files_inner, text="",
            font=("Segoe UI", 11), text_color=TEXT_PRIMARY, anchor="w",
            justify="left"
        )
        self.cons_files_label.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            self.cons_files_inner, text="Trocar", font=("Segoe UI", 10),
            fg_color=BG_INPUT, hover_color=BORDER_CARD,
            text_color=TEXT_SECONDARY, height=28, corner_radius=6, width=70,
            command=self._on_cons_browse,
        ).pack(side="right")

        # ======== ACTION BUTTONS ========
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 12))

        self.cons_process_btn = ctk.CTkButton(
            btn_frame, text="  Consolidar e Gerar PDF",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=self._darken(ACCENT_GREEN),
            height=44, corner_radius=10, state="disabled",
            command=self._on_cons_process,
        )
        self.cons_process_btn.pack(side="left", padx=(0, 10))

        self.cons_open_btn = ctk.CTkButton(
            btn_frame, text="  Abrir PDF",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=44, corner_radius=10, state="disabled",
            command=self._on_cons_open_pdf,
        )
        self.cons_open_btn.pack(side="left", padx=(0, 10))

        self.cons_open_folder_btn = ctk.CTkButton(
            btn_frame, text="  Abrir Pasta",
            font=("Segoe UI", 12),
            fg_color="#6b7280", hover_color="#555d6a",
            height=44, corner_radius=10,
            command=self._on_cons_open_folder,
        )
        self.cons_open_folder_btn.pack(side="left")

        # ======== RESULTS AREA (hidden initially) ========
        self.cons_results_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=12,
            border_width=1, border_color=BORDER_CARD
        )
        self.cons_results_frame.pack(fill="x", pady=(0, 14))
        self.cons_results_frame.pack_forget()

        self.cons_results_inner = ctk.CTkFrame(self.cons_results_frame, fg_color="transparent")
        self.cons_results_inner.pack(fill="x", padx=16, pady=12)

        # ======== LOG ========
        self.cons_log_label = ctk.CTkLabel(
            content, text="Log de Processamento",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.cons_log_label.pack(fill="x", pady=(4, 8))

        self.cons_log_frame = ctk.CTkFrame(
            content, fg_color=BG_LOG, corner_radius=10,
        )
        self.cons_log_frame.pack(fill="x", pady=(0, 8))

        self.cons_log_text = ctk.CTkTextbox(
            self.cons_log_frame, font=("Consolas", 9),
            fg_color=BG_LOG, text_color=TEXT_LOG,
            corner_radius=10, height=180, wrap="word"
        )
        self.cons_log_text.pack(fill="x", padx=4, pady=4)
        self.cons_log_text.insert("1.0", "  Aguardando arquivos...\n")
        self.cons_log_text.configure(state="disabled")

        # ======== STATUS BAR ========
        self.cons_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.cons_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.cons_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.cons_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.cons_status_dot.pack(side="left")

        self.cons_status_text = ctk.CTkLabel(
            si, text="  Aguardando arquivos...",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.cons_status_text.pack(side="left")

        # Internal state
        self._cons_file_paths = []
        self._cons_pdf_path = None
        self._cons_portfolio = None

        return page

    # -----------------------------------------------------------------
    #  CONSOLIDADOR: Acoes
    # -----------------------------------------------------------------
    def _cons_log(self, msg):
        """Append message to consolidador log."""
        self.cons_log_text.configure(state="normal")
        self.cons_log_text.insert("end", f"  {msg}\n")
        self.cons_log_text.see("end")
        self.cons_log_text.configure(state="disabled")

    def _on_cons_browse(self):
        paths = filedialog.askopenfilenames(
            title="Selecionar Relatorios de Carteira",
            initialdir=os.path.join(BASE_DIR, "Mesa Produtos", "Consolidador"),
            filetypes=[
                ("Relatorios", "*.pdf *.xlsx *.xls"),
                ("PDF files", "*.pdf"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ]
        )
        if paths:
            self._cons_file_paths = list(paths)
            self._cons_pdf_path = None
            self._cons_portfolio = None

            # Update UI
            names = [os.path.basename(p) for p in self._cons_file_paths]
            label_text = f"{len(names)} arquivo(s):  " + ",  ".join(names)
            self.cons_files_label.configure(text=label_text)
            self.cons_files_frame.pack(fill="x", pady=(0, 12))

            self.cons_drop_title.configure(text=f"{len(names)} arquivo(s) selecionado(s)")
            self.cons_drop_sub.configure(text="Clique para trocar")

            self.cons_process_btn.configure(state="normal")
            self.cons_open_btn.configure(state="disabled")

            # Clear previous results
            for w in self.cons_results_inner.winfo_children():
                w.destroy()
            self.cons_results_frame.pack_forget()

            # Clear log
            self.cons_log_text.configure(state="normal")
            self.cons_log_text.delete("1.0", "end")
            self.cons_log_text.insert("1.0", "  Arquivos carregados. Clique em 'Consolidar e Gerar PDF'.\n")
            self.cons_log_text.configure(state="disabled")

            self.cons_status_dot.configure(text_color=ACCENT_BLUE)
            self.cons_status_text.configure(text=f"  {len(names)} arquivo(s) pronto(s) para consolidar")

    def _cons_detect_and_parse(self, file_path):
        """Detecta instituicao e parseia o arquivo."""
        import pdfplumber
        import warnings
        warnings.filterwarnings("ignore")

        ext = os.path.splitext(file_path)[1].lower()
        fname = os.path.basename(file_path)

        if ext == ".pdf":
            with pdfplumber.open(file_path) as pdf:
                pages_text = ""
                for p in pdf.pages[:3]:
                    pages_text += (p.extract_text() or "") + "\n"
                pages_lower = pages_text.lower()

                # Detecta relatorio consolidado gerado pelo Somus (tem dados uteis)
                if "somus capital" in pages_lower and "consolidado multi" in pages_lower:
                    from agente_investimentos.consolidador.somus_parser import parse_somus_pdf
                    self._cons_log(f"[{fname}] Detectado: Relatorio Consolidado Somus")
                    return parse_somus_pdf(file_path)

                # Detecta XP
                if ("xp investimentos" in pages_lower or "assessor" in pages_lower
                        or "posicao detalhada" in pages_lower.replace("\xe7", "c").replace("\xe3", "a")
                        or "patrimonio investimento" in pages_lower.replace("\xf4", "o").replace("\xea", "e")):
                    from agente_investimentos.consolidador.xp_parser import parse_xp_pdf
                    self._cons_log(f"[{fname}] Detectado: XP Investimentos")
                    return parse_xp_pdf(file_path)

                # Detecta Itau
                if "itau" in pages_lower or "personnalite" in pages_lower or "total investido" in pages_lower:
                    from agente_investimentos.consolidador.itau_parser import parse_itau_pdf
                    self._cons_log(f"[{fname}] Detectado: Itau")
                    return parse_itau_pdf(file_path)

                # Detecta Safra
                if "safra" in pages_lower or "safrabm" in pages_lower:
                    from agente_investimentos.consolidador.safra_parser import parse_safra_pdf
                    self._cons_log(f"[{fname}] Detectado: Safra")
                    return parse_safra_pdf(file_path)

            # Fallback Safra
            from agente_investimentos.consolidador.safra_parser import parse_safra_pdf
            self._cons_log(f"[{fname}] Fallback: tentando parser Safra")
            return parse_safra_pdf(file_path)

        elif ext in (".xlsx", ".xls"):
            from agente_investimentos.consolidador.btg_parser import parse_btg_xlsx
            self._cons_log(f"[{fname}] Detectado: BTG (Excel)")
            return parse_btg_xlsx(file_path)

        return None

    def _on_cons_process(self):
        if not self._cons_file_paths:
            return
        self.cons_process_btn.configure(state="disabled", text="  Processando...")
        self.cons_status_dot.configure(text_color=ACCENT_ORANGE)
        self.cons_status_text.configure(text="  Processando relatorios...")

        # Clear log
        self.cons_log_text.configure(state="normal")
        self.cons_log_text.delete("1.0", "end")
        self.cons_log_text.configure(state="disabled")

        threading.Thread(target=self._cons_process_thread, daemon=True).start()

    def _cons_process_thread(self):
        from agente_investimentos.consolidador.models import ConsolidatedPortfolio
        try:
            instituicoes = []
            errors = []

            for path in self._cons_file_paths:
                fname = os.path.basename(path)
                self._cons_log(f"Processando: {fname}")
                try:
                    result = self._cons_detect_and_parse(path)
                    if result and result.patrimonio_bruto > 0:
                        instituicoes.append(result)
                        self._cons_log(
                            f"  OK: {result.nome} - {result.num_ativos} ativos - "
                            f"{fmt_currency(result.patrimonio_bruto)}"
                        )
                    elif result:
                        self._cons_log(f"  AVISO: {fname} parseado mas sem patrimonio")
                    else:
                        errors.append(fname)
                        self._cons_log(f"  ERRO: {fname} - formato nao reconhecido")
                except Exception as e:
                    errors.append(fname)
                    self._cons_log(f"  ERRO: {fname} - {str(e)[:120]}")

            if not instituicoes:
                self._cons_log("\nNenhuma instituicao valida encontrada.")
                self.after(0, lambda: self._cons_finish(False, "Nenhum arquivo valido"))
                return

            # Consolida
            cp = ConsolidatedPortfolio(instituicoes=instituicoes)
            self._cons_portfolio = cp

            self._cons_log(f"\nConsolidado: {cp.num_instituicoes} instituicoes, "
                          f"{cp.num_ativos_total} ativos, {fmt_currency(cp.patrimonio_bruto_total)}")

            # Gera PDF
            self._cons_log("Gerando PDF consolidado...")
            os.makedirs(CONSOLIDADOR_OUTPUT_DIR, exist_ok=True)

            from agente_investimentos.consolidador.pdf_builder import ConsolidadorPDFBuilder
            builder = ConsolidadorPDFBuilder(cp)
            pdf_path = builder.build()

            # Copia para pasta local do app
            import shutil
            local_pdf = os.path.join(CONSOLIDADOR_OUTPUT_DIR, "Relatorio_Consolidado.pdf")
            shutil.copy2(str(pdf_path), local_pdf)
            self._cons_pdf_path = local_pdf

            self._cons_log(f"PDF gerado: {local_pdf}")
            self._cons_log("\nProcessamento concluido com sucesso!")

            self.after(0, lambda: self._cons_finish(True, ""))

        except Exception as e:
            self._cons_log(f"\nERRO GERAL: {str(e)[:200]}")
            self.after(0, lambda: self._cons_finish(False, str(e)[:100]))

    def _cons_finish(self, success, error_msg):
        self.cons_process_btn.configure(state="normal", text="  Consolidar e Gerar PDF")

        if success:
            self.cons_status_dot.configure(text_color=ACCENT_GREEN)
            self.cons_status_text.configure(text="  Consolidacao concluida com sucesso!")
            self.cons_open_btn.configure(state="normal")
            self._cons_show_results()
        else:
            self.cons_status_dot.configure(text_color=ACCENT_RED)
            self.cons_status_text.configure(text=f"  Erro: {error_msg}")

    def _cons_show_results(self):
        """Mostra KPIs e resumo da consolidacao."""
        cp = self._cons_portfolio
        if not cp:
            return

        # Clear previous results
        for w in self.cons_results_inner.winfo_children():
            w.destroy()

        self.cons_results_frame.pack(fill="x", pady=(0, 14))

        # Title
        ctk.CTkLabel(
            self.cons_results_inner, text="Resultado da Consolidacao",
            font=("Segoe UI", 15, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        ).pack(fill="x", pady=(0, 12))

        # KPI cards row
        kpi_row = ctk.CTkFrame(self.cons_results_inner, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 12))

        kpis = [
            ("Patrimonio Bruto", fmt_currency(cp.patrimonio_bruto_total), ACCENT_GREEN),
            ("Patrimonio Liquido", fmt_currency(cp.patrimonio_liquido_total), ACCENT_BLUE),
            ("Instituicoes", str(cp.num_instituicoes), ACCENT_PURPLE),
            ("Total Ativos", str(cp.num_ativos_total), ACCENT_ORANGE),
            ("Rent. Ano", fmt_pct(cp.rent_ano_ponderada()), ACCENT_GREEN if cp.rent_ano_ponderada() >= 0 else ACCENT_RED),
        ]

        for i, (label, value, color) in enumerate(kpis):
            card = ctk.CTkFrame(kpi_row, fg_color=BG_CARD, corner_radius=10,
                               border_width=1, border_color=BORDER_CARD)
            card.pack(side="left", fill="x", expand=True, padx=(0 if i == 0 else 6, 0))

            ctk.CTkLabel(
                card, text=label,
                font=("Segoe UI", 9), text_color=TEXT_TERTIARY, anchor="w"
            ).pack(fill="x", padx=12, pady=(10, 0))

            ctk.CTkLabel(
                card, text=value,
                font=("Segoe UI", 14, "bold"), text_color=color, anchor="w"
            ).pack(fill="x", padx=12, pady=(2, 10))

        # Per-institution summary
        for inst in cp.instituicoes:
            inst_frame = ctk.CTkFrame(self.cons_results_inner, fg_color=BG_INPUT, corner_radius=8)
            inst_frame.pack(fill="x", pady=(4, 4))

            row = ctk.CTkFrame(inst_frame, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=8)

            ctk.CTkLabel(
                row, text=f"\u25cf  {inst.nome}",
                font=("Segoe UI", 12, "bold"), text_color=TEXT_PRIMARY
            ).pack(side="left")

            ctk.CTkLabel(
                row, text=f"{inst.num_ativos} ativos",
                font=("Segoe UI", 10), text_color=TEXT_SECONDARY
            ).pack(side="left", padx=(16, 0))

            ctk.CTkLabel(
                row, text=fmt_currency(inst.patrimonio_bruto),
                font=("Segoe UI", 11, "bold"), text_color=ACCENT_GREEN
            ).pack(side="right", padx=(0, 8))

            if inst.rent_carteira_ano:
                ctk.CTkLabel(
                    row, text=fmt_pct(inst.rent_carteira_ano),
                    font=("Segoe UI", 10), text_color=TEXT_SECONDARY
                ).pack(side="right", padx=(0, 16))

    def _on_cons_open_pdf(self):
        if self._cons_pdf_path and os.path.exists(self._cons_pdf_path):
            os.startfile(self._cons_pdf_path)

    def _on_cons_open_folder(self):
        os.makedirs(CONSOLIDADOR_OUTPUT_DIR, exist_ok=True)
        os.startfile(CONSOLIDADOR_OUTPUT_DIR)

    # -----------------------------------------------------------------
    #  PAGE: ORGANIZADOR DE PLANILHAS
    # -----------------------------------------------------------------
    def _build_organizador_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Organizador", subtitle="Planilhas Excel")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # ======== DROP ZONE ========
        self.org_drop_card = ctk.CTkFrame(
            content, fg_color="#f0f4ff", corner_radius=14,
            border_width=2, border_color=ACCENT_BLUE
        )
        self.org_drop_card.pack(fill="x", pady=(0, 16))
        self.org_drop_card.bind("<Button-1>", lambda e: self._on_browse_excel())

        drop_inner = ctk.CTkFrame(self.org_drop_card, fg_color="transparent")
        drop_inner.pack(fill="x", padx=30, pady=36)
        drop_inner.bind("<Button-1>", lambda e: self._on_browse_excel())

        self.org_icon_label = ctk.CTkLabel(
            drop_inner, text="\u25a6",
            font=("Segoe UI", 44), text_color=ACCENT_BLUE
        )
        self.org_icon_label.pack()
        self.org_icon_label.bind("<Button-1>", lambda e: self._on_browse_excel())

        self.org_drop_title = ctk.CTkLabel(
            drop_inner, text="Clique para selecionar uma planilha Excel",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY
        )
        self.org_drop_title.pack(pady=(8, 2))
        self.org_drop_title.bind("<Button-1>", lambda e: self._on_browse_excel())

        self.org_drop_sub = ctk.CTkLabel(
            drop_inner, text="Formatos aceitos:  .xlsx  .xls",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.org_drop_sub.pack()
        self.org_drop_sub.bind("<Button-1>", lambda e: self._on_browse_excel())

        self.org_browse_btn = ctk.CTkButton(
            drop_inner, text="  Selecionar Arquivo",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=40, corner_radius=8, width=200,
            command=self._on_browse_excel,
        )
        self.org_browse_btn.pack(pady=(14, 0))

        # ======== FILE INFO BAR (hidden initially) ========
        self.org_file_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.org_file_frame.pack(fill="x", pady=(0, 12))
        self.org_file_frame.pack_forget()

        fi = ctk.CTkFrame(self.org_file_frame, fg_color="transparent")
        fi.pack(fill="x", padx=16, pady=10)

        self.org_file_icon = ctk.CTkLabel(
            fi, text="\u25cf", font=("Segoe UI", 12), text_color=ACCENT_GREEN
        )
        self.org_file_icon.pack(side="left")

        self.org_file_name = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 11, "bold"), text_color=TEXT_PRIMARY
        )
        self.org_file_name.pack(side="left", padx=(6, 0))

        self.org_file_info = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 10), text_color=TEXT_SECONDARY
        )
        self.org_file_info.pack(side="left", padx=(12, 0))

        ctk.CTkButton(
            fi, text="Trocar", font=("Segoe UI", 10),
            fg_color=BG_INPUT, hover_color=BORDER_CARD,
            text_color=TEXT_SECONDARY, height=28, corner_radius=6, width=70,
            command=self._on_browse_excel,
        ).pack(side="right")

        # ======== PREVIEW ========
        self.org_preview_label = ctk.CTkLabel(
            content, text="Preview dos Dados",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.org_preview_label.pack(fill="x", pady=(4, 8))

        self.org_preview_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=12,
            border_width=1, border_color=BORDER_CARD
        )
        self.org_preview_frame.pack(fill="x", pady=(0, 14))

        self.org_preview_text = ctk.CTkTextbox(
            self.org_preview_frame, font=("Consolas", 9),
            fg_color=BG_CARD, text_color=TEXT_PRIMARY,
            corner_radius=10, height=200, wrap="none"
        )
        self.org_preview_text.pack(fill="x", padx=4, pady=4)
        self.org_preview_text.insert("1.0", "  Nenhum arquivo carregado...")
        self.org_preview_text.configure(state="disabled")

        # ======== ACTION BUTTONS ========
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 12))

        self.org_process_btn = ctk.CTkButton(
            btn_frame, text="  Processar e Organizar",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=self._darken(ACCENT_GREEN),
            height=44, corner_radius=10, state="disabled",
            command=self._on_process_excel,
        )
        self.org_process_btn.pack(side="left", padx=(0, 10))

        self.org_open_btn = ctk.CTkButton(
            btn_frame, text="  Abrir Resultado",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=44, corner_radius=10, state="disabled",
            command=self._on_open_organized,
        )
        self.org_open_btn.pack(side="left", padx=(0, 10))

        self.org_open_folder_btn = ctk.CTkButton(
            btn_frame, text="  Abrir Pasta",
            font=("Segoe UI", 12),
            fg_color="#6b7280", hover_color="#555d6a",
            height=44, corner_radius=10,
            command=self._on_open_org_folder,
        )
        self.org_open_folder_btn.pack(side="left")

        # ======== STATUS BAR ========
        self.org_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.org_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.org_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.org_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.org_status_dot.pack(side="left")

        self.org_status_text = ctk.CTkLabel(
            si, text="  Aguardando arquivo...",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.org_status_text.pack(side="left")

        # Internal state
        self._org_input_path = None
        self._org_output_path = None

        return page

    # -----------------------------------------------------------------
    #  ORGANIZADOR: Ações
    # -----------------------------------------------------------------
    def _on_browse_excel(self):
        path = filedialog.askopenfilename(
            title="Selecionar Planilha Excel",
            initialdir=os.path.join(BASE_DIR, "Mesa Produtos", "Organizador"),
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._org_load_file(path)

    def _org_load_file(self, path):
        self._org_input_path = path
        self._org_output_path = None

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            total_rows = 0
            total_cols = 0
            for ws in wb:
                total_rows += ws.max_row or 0
                total_cols = max(total_cols, ws.max_column or 0)
            wb.close()

            fname = os.path.basename(path)
            self.org_file_name.configure(text=fname)
            info = f"{len(sheets)} aba(s)  |  ~{total_rows} linhas  |  {total_cols} colunas"
            self.org_file_info.configure(text=info)
            self.org_file_frame.pack(fill="x", pady=(0, 12))

            # Update drop zone visual
            self.org_drop_title.configure(text=fname)
            self.org_drop_sub.configure(text="Arquivo carregado - clique para trocar")
            self.org_icon_label.configure(text_color=ACCENT_GREEN, text="\u2713")
            self.org_drop_card.configure(border_color=ACCENT_GREEN, fg_color="#f0fff4")

            self._org_show_preview(path)

            self.org_process_btn.configure(state="normal")
            self.org_open_btn.configure(state="disabled")

            self.org_status_dot.configure(text_color=ACCENT_BLUE)
            self.org_status_text.configure(
                text=f"  Arquivo carregado. Clique em 'Processar e Organizar'."
            )

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo:\n{e}")

    def _org_show_preview(self, path, records=None):
        try:
            self.org_preview_text.configure(state="normal")
            self.org_preview_text.delete("1.0", "end")

            if records:
                # Detect which format by checking keys
                if "produto" in records[0]:
                    # 4-row format
                    headers = ["CONTA", "ATIVO", "PRODUTO", "TIPO", "DT INICIO", "DT VENC", "QTD", "VALOR", "STATUS"]
                    col_widths = [10, 8, 16, 18, 12, 12, 8, 14, 18]
                    header_line = "".join(h.ljust(w) for h, w in zip(headers, col_widths))
                    self.org_preview_text.insert("end", header_line + "\n")
                    self.org_preview_text.insert("end", "-" * sum(col_widths) + "\n")

                    for rec in records[:20]:
                        vals = [
                            rec["conta"][:9], rec["ativo"][:7],
                            rec["produto"][:15], rec["tipo"][:17],
                            rec["data_inicio"][:11], rec["data_vencimento"][:11],
                            rec["qtd_str"][:7], rec["valor_str"][:13],
                            rec["status"][:17],
                        ]
                        line = "".join(v.ljust(w) for v, w in zip(vals, col_widths))
                        self.org_preview_text.insert("end", line + "\n")
                else:
                    # Condensed XP format
                    headers = [
                        "CLIENTE", "CONTA", "ATIVO", "DATA", "HORA",
                        "TIPO", "ORIGEM", "QTD", "VALOR",
                        "QTD EXEC", "VALOR EXEC", "STATUS"
                    ]
                    col_widths = [25, 10, 35, 12, 10, 8, 16, 12, 16, 12, 16, 12]
                    header_line = "".join(h.ljust(w) for h, w in zip(headers, col_widths))
                    self.org_preview_text.insert("end", header_line + "\n")
                    self.org_preview_text.insert("end", "-" * sum(col_widths) + "\n")

                    for rec in records[:20]:
                        vals = [
                            rec["cliente"][:24], rec["conta"][:9],
                            rec["ativo"][:34], rec["data_str"][:11], rec["hora"][:9],
                            rec["tipo"][:7], rec["origem"][:15],
                            rec["qtd_str"][:11], rec["valor_str"][:15],
                            rec["qtd_exec_str"][:11], rec["valor_exec_str"][:15],
                            rec["status"][:11],
                        ]
                        line = "".join(v.ljust(w) for v, w in zip(vals, col_widths))
                        self.org_preview_text.insert("end", line + "\n")

                if len(records) > 20:
                    self.org_preview_text.insert("end", f"\n  ... e mais {len(records) - 20} registros")
            else:
                # Show raw file preview
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                for i, row in enumerate(ws.iter_rows(max_row=min(18, ws.max_row or 18), values_only=True)):
                    cells = [str(c)[:50] if c is not None else "" for c in row]
                    self.org_preview_text.insert("end", "  |  ".join(cells) + "\n")
                wb.close()

            self.org_preview_text.configure(state="disabled")
        except Exception:
            pass

    def _detect_4row_format(self, ws):
        """Detect the 4-row grouped format: conta / ativo+produto / detalhes / status."""
        if (ws.max_column or 0) > 2:
            return False
        max_row = ws.max_row or 0
        if max_row < 4:
            return False
        # Row 1 = numeric (conta), Row 3 has date pattern DD/MM/YYYY, Row 4 = status text
        r1 = ws.cell(row=1, column=1).value
        r3 = sanitize_text(str(ws.cell(row=3, column=1).value or ""))
        r4 = sanitize_text(str(ws.cell(row=4, column=1).value or ""))
        if r1 is not None and re.search(r'\d{2}/\d{2}/\d{4}', r3):
            status_keywords = ["Pendente", "Cancelado", "Executad", "Total", "Parcial", "Registro"]
            if any(kw in r4 for kw in status_keywords):
                return True
        return False

    def _parse_4row_records(self, ws):
        """Parse the 4-row grouped format into structured records."""
        records = []
        max_row = ws.max_row or 0
        # Only iterate complete groups of 4
        usable_rows = (max_row // 4) * 4

        for i in range(0, usable_rows, 4):
            r1 = str(ws.cell(row=i + 1, column=1).value or "").strip()
            r2 = str(ws.cell(row=i + 2, column=1).value or "").strip()
            r3 = sanitize_text(str(ws.cell(row=i + 3, column=1).value or "")).strip()
            r4 = sanitize_text(str(ws.cell(row=i + 4, column=1).value or "")).strip()

            # Row 1: Conta
            conta = r1

            # Row 2: Ativo + Produto (e.g. "AURA33Smart Coupon")
            m_ativo = re.match(r'^([A-Z]{2,6}\d{1,4})(.*)', r2)
            if m_ativo:
                ativo = m_ativo.group(1)
                produto = m_ativo.group(2).strip()
            else:
                ativo = r2
                produto = ""

            # Row 3: Tipo + DataInicio + DataVencimento + Quantidade + Valor
            # e.g. "Nova Contratacao04/03/202605/08/2026310R$ 895,03"
            tipo = ""
            data_inicio = ""
            data_vencimento = ""
            qtd_str = ""
            qtd_num = 0.0
            valor_str = ""
            valor_num = 0.0

            dates = list(re.finditer(r'\d{2}/\d{2}/\d{4}', r3))
            if len(dates) >= 2:
                tipo = r3[:dates[0].start()].strip()
                data_inicio = dates[0].group()
                data_vencimento = dates[1].group()
                after_dates = r3[dates[1].end():]
                # after_dates e.g. "310R$ 895,03" or "2310R$ 6.669,43"
                m_val = re.match(r'^(\d[\d.]*)R\$\s*([\d.,]+)', after_dates)
                if m_val:
                    qtd_str = m_val.group(1)
                    qtd_num = float(qtd_str.replace('.', '').replace(',', '.'))
                    valor_str = m_val.group(2)
                    valor_num = float(valor_str.replace('.', '').replace(',', '.'))
            elif len(dates) == 1:
                tipo = r3[:dates[0].start()].strip()
                data_inicio = dates[0].group()
                after_date = r3[dates[0].end():]
                m_val = re.match(r'^(\d[\d.]*)R\$\s*([\d.,]+)', after_date)
                if m_val:
                    qtd_str = m_val.group(1)
                    qtd_num = float(qtd_str.replace('.', '').replace(',', '.'))
                    valor_str = m_val.group(2)
                    valor_num = float(valor_str.replace('.', '').replace(',', '.'))

            # Row 4: Status
            status = r4

            records.append({
                "conta": conta,
                "ativo": ativo,
                "produto": produto,
                "tipo": tipo,
                "data_inicio": data_inicio,
                "data_vencimento": data_vencimento,
                "qtd_str": qtd_str,
                "qtd_num": qtd_num,
                "valor_str": f"R$ {valor_str}" if valor_str else "",
                "valor_num": valor_num,
                "status": status,
            })

        return records

    def _detect_condensed_format(self, ws):
        """Detect if the sheet uses the condensed XP single-column format."""
        if (ws.max_column or 0) > 2:
            return False
        max_row = ws.max_row or 0
        if max_row < 6 or max_row % 6 != 0:
            return False
        # Check row 6 pattern = "Ver Nota"
        r6 = ws.cell(row=6, column=1).value
        if r6 and "Ver Nota" in str(r6):
            return True
        # Check row 1 has "Conta: "
        r1 = str(ws.cell(row=1, column=1).value or "")
        if "Conta:" in r1:
            return True
        return False

    def _parse_condensed_records(self, ws):
        """Parse the condensed XP format into structured records.
        Uses 'Conta:' markers to find each block instead of fixed 6-row stride,
        to handle injected header rows in the middle of the data."""
        records = []
        max_row = ws.max_row or 0

        # Find all rows that contain "Conta:" — these are block starts
        conta_rows = []
        for r in range(1, max_row + 1):
            v = str(ws.cell(r, 1).value or "")
            if "Conta:" in v:
                conta_rows.append(r)

        for start in conta_rows:
            # Each block: row+0=client, row+1=asset, row+2=date, row+3=data, row+4=status
            if start + 4 > max_row:
                break
            r1 = str(ws.cell(row=start, column=1).value or "")
            r2 = str(ws.cell(row=start + 1, column=1).value or "")
            r3 = str(ws.cell(row=start + 2, column=1).value or "")
            r4 = str(ws.cell(row=start + 3, column=1).value or "")
            r5 = str(ws.cell(row=start + 4, column=1).value or "")

            # --- Row 1: Cliente + Conta ---
            if "Conta:" in r1:
                parts = r1.split("Conta:")
                cliente = sanitize_text(parts[0]).strip()
                conta = parts[1].strip()
            else:
                cliente = sanitize_text(r1).strip()
                conta = ""

            # --- Row 2: Ativo ---
            ativo = sanitize_text(r2).strip()

            # --- Row 3: Data + Hora (M/D/YYYYhh:mm:ss) ---
            m_dt = re.match(r'^(\d{1,2}/\d{1,2}/\d{4})(\d{2}:\d{2}:\d{2})$', r3.strip())
            if m_dt:
                # Convert M/D/YYYY to DD/MM/YYYY
                try:
                    dt = datetime.strptime(m_dt.group(1), "%m/%d/%Y")
                    data_str = dt.strftime("%d/%m/%Y")
                    data_dt = dt
                except ValueError:
                    data_str = m_dt.group(1)
                    data_dt = None
                hora_str = m_dt.group(2)
            else:
                data_str = r3.strip()
                data_dt = None
                hora_str = ""

            # --- Row 4: Tipo + Origem + Qtd + Valor + QtdExec + ValorExec ---
            r4_clean = r4.replace('\xa0', ' ')
            parts_r4 = r4_clean.split('R$ ')

            tipo = ""
            origem = ""
            qtd_str = ""
            valor_str = ""
            qtd_exec_str = ""
            valor_exec_str = ""
            qtd_num = 0.0
            valor_num = 0.0
            qtd_exec_num = 0.0
            valor_exec_num = 0.0

            if len(parts_r4) >= 3:
                # Part 0: "CompraEstoque XP300" or "VendaEstoque clientes2250"
                p0 = parts_r4[0]
                m_p0 = re.match(r'^(Compra|Venda)(.*?)(\d[\d.,]*)$', p0)
                if m_p0:
                    tipo = m_p0.group(1)
                    origem = m_p0.group(2).strip()
                    qtd_str = m_p0.group(3)
                    qtd_num = float(qtd_str.replace('.', '').replace(',', '.'))

                # Part 1: "30.690,99300" -> valor,XX + qtd_exec
                p1 = parts_r4[1]
                m_p1 = re.match(r'^([\d.]+,\d{2})(.+)$', p1)
                if m_p1:
                    valor_str = m_p1.group(1)
                    qtd_exec_str = m_p1.group(2)
                    valor_num = float(valor_str.replace('.', '').replace(',', '.'))
                    qtd_exec_num = float(qtd_exec_str.replace('.', '').replace(',', '.'))

                # Part 2: "30.690,99" -> valor_exec
                valor_exec_str = parts_r4[2].strip()
                valor_exec_num = float(valor_exec_str.replace('.', '').replace(',', '.'))
            elif len(parts_r4) == 2:
                p0 = parts_r4[0]
                m_p0 = re.match(r'^(Compra|Venda)(.*?)(\d[\d.,]*)$', p0)
                if m_p0:
                    tipo = m_p0.group(1)
                    origem = m_p0.group(2).strip()
                    qtd_str = m_p0.group(3)
                    qtd_num = float(qtd_str.replace('.', '').replace(',', '.'))
                valor_str = parts_r4[1].strip()
                valor_num = float(valor_str.replace('.', '').replace(',', '.'))

            # --- Row 5: Status ---
            status = sanitize_text(r5).strip()

            records.append({
                "cliente": cliente,
                "conta": conta,
                "ativo": ativo,
                "data_str": data_str,
                "data_dt": data_dt,
                "hora": hora_str,
                "tipo": tipo,
                "origem": origem,
                "qtd_str": qtd_str,
                "qtd_num": qtd_num,
                "valor_str": valor_str,
                "valor_num": valor_num,
                "qtd_exec_str": qtd_exec_str,
                "qtd_exec_num": qtd_exec_num,
                "valor_exec_str": valor_exec_str,
                "valor_exec_num": valor_exec_num,
                "status": status,
            })

        return records

    def _on_process_excel(self):
        if not self._org_input_path:
            return

        self.org_process_btn.configure(state="disabled")
        self.org_status_dot.configure(text_color=ACCENT_ORANGE)
        self.org_status_text.configure(text="  Processando...")

        threading.Thread(target=self._run_process_excel, daemon=True).start()

    def _run_process_excel(self):
        try:
            from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
            from openpyxl.utils import get_column_letter

            path = self._org_input_path
            wb_in = openpyxl.load_workbook(path, data_only=True)
            ws_in = wb_in[wb_in.sheetnames[0]]

            is_4row = self._detect_4row_format(ws_in)
            is_condensed = self._detect_condensed_format(ws_in) if not is_4row else False

            if is_4row:
                records_4row = self._parse_4row_records(ws_in)
                records = None
            elif is_condensed:
                records_4row = None
                records = self._parse_condensed_records(ws_in)
            else:
                records_4row = None
                records = None

            wb_out = openpyxl.Workbook()
            ws_out = wb_out.active
            ws_out.title = "Dados Organizados"

            # Styles - Somus Capital branding
            header_fill = PatternFill(start_color="004d33", end_color="004d33", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

            data_font = Font(name="Calibri", size=10)
            data_align_left = Alignment(horizontal="left", vertical="center")
            data_align_center = Alignment(horizontal="center", vertical="center")
            data_align_right = Alignment(horizontal="right", vertical="center")

            alt_fill = PatternFill(start_color="f0f7f4", end_color="f0f7f4", fill_type="solid")
            white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

            thin_border = Border(
                left=Side(style="thin", color="d0d0d0"),
                right=Side(style="thin", color="d0d0d0"),
                top=Side(style="thin", color="d0d0d0"),
                bottom=Side(style="thin", color="d0d0d0"),
            )

            if records_4row:
                # ===== 4-ROW GROUPED FORMAT (Smart Coupon etc.) =====
                headers = [
                    "CONTA", "ATIVO", "PRODUTO", "TIPO",
                    "DATA INICIO", "DATA VENCIMENTO",
                    "QUANTIDADE", "VALOR (R$)", "STATUS"
                ]
                col_widths = [14, 12, 20, 22, 16, 18, 14, 18, 20]

                for col_idx, (hdr, width) in enumerate(zip(headers, col_widths), 1):
                    cell = ws_out.cell(row=1, column=col_idx)
                    cell.value = hdr
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_align
                    cell.border = thin_border
                    ws_out.column_dimensions[get_column_letter(col_idx)].width = width

                for row_idx, rec in enumerate(records_4row):
                    out_row = row_idx + 2
                    fill = alt_fill if row_idx % 2 == 0 else white_fill

                    row_values = [
                        (int(rec["conta"]) if rec["conta"].isdigit() else rec["conta"], data_align_center, None),
                        (rec["ativo"], data_align_center, None),
                        (rec["produto"], data_align_left, None),
                        (rec["tipo"], data_align_left, None),
                        (rec["data_inicio"], data_align_center, None),
                        (rec["data_vencimento"], data_align_center, None),
                        (rec["qtd_num"], data_align_right, '#,##0'),
                        (rec["valor_num"], data_align_right, '#,##0.00'),
                        (rec["status"], data_align_center, None),
                    ]

                    for col_idx, (val, align, num_fmt) in enumerate(row_values, 1):
                        cell = ws_out.cell(row=out_row, column=col_idx)
                        cell.value = val if val else ""
                        cell.font = data_font
                        cell.fill = fill
                        cell.border = thin_border
                        cell.alignment = align
                        if num_fmt and val:
                            cell.number_format = num_fmt

                total_processed = len(records_4row)
                records = records_4row  # for preview

            elif records:
                # ===== CONDENSED XP FORMAT =====
                headers = [
                    "CLIENTE", "CONTA", "ATIVO", "DATA", "HORA",
                    "TIPO", "ORIGEM", "QUANTIDADE", "VALOR (R$)",
                    "QTD EXECUTADA", "VALOR EXECUTADO (R$)", "STATUS"
                ]
                col_widths = [35, 14, 42, 14, 12, 12, 20, 16, 20, 16, 22, 14]

                # Write headers
                for col_idx, (hdr, width) in enumerate(zip(headers, col_widths), 1):
                    cell = ws_out.cell(row=1, column=col_idx)
                    cell.value = hdr
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_align
                    cell.border = thin_border
                    ws_out.column_dimensions[get_column_letter(col_idx)].width = width

                # Write data
                for row_idx, rec in enumerate(records):
                    out_row = row_idx + 2
                    fill = alt_fill if row_idx % 2 == 0 else white_fill

                    row_values = [
                        (rec["cliente"], data_align_left, None),
                        (int(rec["conta"]) if rec["conta"].isdigit() else rec["conta"], data_align_center, None),
                        (rec["ativo"], data_align_left, None),
                        (rec["data_dt"], data_align_center, "DD/MM/YYYY"),
                        (rec["hora"], data_align_center, None),
                        (rec["tipo"], data_align_center, None),
                        (rec["origem"], data_align_left, None),
                        (rec["qtd_num"], data_align_right, '#,##0.000' if rec["qtd_num"] != int(rec["qtd_num"]) else '#,##0'),
                        (rec["valor_num"], data_align_right, '#,##0.00'),
                        (rec["qtd_exec_num"], data_align_right, '#,##0.000' if rec["qtd_exec_num"] != int(rec["qtd_exec_num"]) else '#,##0'),
                        (rec["valor_exec_num"], data_align_right, '#,##0.00'),
                        (rec["status"], data_align_center, None),
                    ]

                    for col_idx, (val, align, num_fmt) in enumerate(row_values, 1):
                        cell = ws_out.cell(row=out_row, column=col_idx)
                        cell.value = val if val else ""
                        cell.font = data_font
                        cell.fill = fill
                        cell.border = thin_border
                        cell.alignment = align
                        if num_fmt and val:
                            cell.number_format = num_fmt

                total_processed = len(records)

            else:
                # ===== GENERIC FORMAT =====
                data = []
                for row in ws_in.iter_rows(values_only=True):
                    data.append(list(row))

                if not data:
                    raise ValueError("Planilha vazia")

                header_idx = 0
                max_filled = 0
                for i, row in enumerate(data[:10]):
                    filled = sum(1 for c in row if c is not None and str(c).strip())
                    if filled > max_filled:
                        max_filled = filled
                        header_idx = i

                headers_row = data[header_idx]
                data_rows = data[header_idx + 1:]
                data_rows = [
                    r for r in data_rows
                    if any(c is not None and str(c).strip() for c in r)
                ]

                n_cols = len(headers_row)
                used_cols = []
                for c in range(n_cols):
                    hv = headers_row[c]
                    if hv is not None and str(hv).strip():
                        used_cols.append(c)
                    elif any(r[c] is not None and str(r[c]).strip() for r in data_rows if c < len(r)):
                        used_cols.append(c)

                if not used_cols:
                    raise ValueError("Nenhuma coluna com dados encontrada")

                for out_col, src_col in enumerate(used_cols, 1):
                    val = headers_row[src_col]
                    cell = ws_out.cell(row=1, column=out_col)
                    cell.value = sanitize_text(val).strip().upper() if val else f"COLUNA {out_col}"
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_align
                    cell.border = thin_border

                for row_idx, row_data in enumerate(data_rows):
                    out_row = row_idx + 2
                    fill = alt_fill if row_idx % 2 == 0 else white_fill
                    for out_col, src_col in enumerate(used_cols, 1):
                        val = row_data[src_col] if src_col < len(row_data) else None
                        cell = ws_out.cell(row=out_row, column=out_col)
                        if isinstance(val, str):
                            val = sanitize_text(val).strip()
                        cell.value = val
                        cell.font = data_font
                        cell.fill = fill
                        cell.border = thin_border
                        if isinstance(val, datetime):
                            cell.number_format = "DD/MM/YYYY"
                            cell.alignment = data_align_center
                        elif isinstance(val, (int, float)):
                            cell.alignment = data_align_right
                            cell.number_format = '#,##0.00'
                        else:
                            cell.alignment = data_align_left

                for out_col, src_col in enumerate(used_cols, 1):
                    max_len = len(str(headers_row[src_col] or ""))
                    for rd in data_rows[:200]:
                        val = rd[src_col] if src_col < len(rd) else None
                        if val:
                            max_len = max(max_len, min(len(str(val)), 50))
                    ws_out.column_dimensions[get_column_letter(out_col)].width = min(max_len + 4, 55)

                total_processed = len(data_rows)

            # Autofilter + freeze
            last_col_letter = get_column_letter(ws_out.max_column)
            ws_out.auto_filter.ref = f"A1:{last_col_letter}{ws_out.max_row}"
            ws_out.freeze_panes = "A2"

            # Save
            os.makedirs(ORGANIZADOR_OUTPUT_DIR, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(path))[0]
            output_name = f"{base_name}_ORGANIZADO.xlsx"
            output_path = os.path.join(ORGANIZADOR_OUTPUT_DIR, output_name)

            wb_out.save(output_path)
            wb_in.close()

            self._org_output_path = output_path

            def _update_ok():
                self.org_process_btn.configure(state="normal")
                self.org_open_btn.configure(state="normal")
                self.org_status_dot.configure(text_color=ACCENT_GREEN)
                self.org_status_text.configure(
                    text=f"  Concluído! {total_processed} registros organizados  -  {output_name}"
                )
                self.org_drop_card.configure(border_color=ACCENT_GREEN)
                self._org_show_preview(output_path, records=records)
                self.org_preview_label.configure(text="Preview do Resultado")

            self.after(0, _update_ok)

        except Exception as e:
            def _update_err():
                self.org_process_btn.configure(state="normal")
                self.org_status_dot.configure(text_color=ACCENT_RED)
                self.org_status_text.configure(text=f"  Erro: {e}")

            self.after(0, _update_err)

    def _on_open_organized(self):
        if self._org_output_path and os.path.exists(self._org_output_path):
            os.startfile(self._org_output_path)

    def _on_open_org_folder(self):
        os.makedirs(ORGANIZADOR_OUTPUT_DIR, exist_ok=True)
        os.startfile(ORGANIZADOR_OUTPUT_DIR)

    # -----------------------------------------------------------------
    #  PAGE: ENVIO SALDOS
    # -----------------------------------------------------------------
    def _build_saldos_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Envio Saldos", subtitle="Saldos Diários")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # ======== DROP ZONE ========
        self.sld_drop_card = ctk.CTkFrame(
            content, fg_color="#f0f4ff", corner_radius=14,
            border_width=2, border_color=ACCENT_BLUE
        )
        self.sld_drop_card.pack(fill="x", pady=(0, 16))
        self.sld_drop_card.bind("<Button-1>", lambda e: self._on_browse_saldos())

        drop_inner = ctk.CTkFrame(self.sld_drop_card, fg_color="transparent")
        drop_inner.pack(fill="x", padx=30, pady=36)
        drop_inner.bind("<Button-1>", lambda e: self._on_browse_saldos())

        self.sld_icon_label = ctk.CTkLabel(
            drop_inner, text="\u21c6",
            font=("Segoe UI", 44), text_color=ACCENT_BLUE
        )
        self.sld_icon_label.pack()
        self.sld_icon_label.bind("<Button-1>", lambda e: self._on_browse_saldos())

        self.sld_drop_title = ctk.CTkLabel(
            drop_inner, text="Clique para selecionar a planilha de saldos",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY
        )
        self.sld_drop_title.pack(pady=(8, 2))
        self.sld_drop_title.bind("<Button-1>", lambda e: self._on_browse_saldos())

        self.sld_drop_sub = ctk.CTkLabel(
            drop_inner, text="Formatos aceitos:  .xlsx  .xls",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.sld_drop_sub.pack()
        self.sld_drop_sub.bind("<Button-1>", lambda e: self._on_browse_saldos())

        self.sld_browse_btn = ctk.CTkButton(
            drop_inner, text="  Selecionar Arquivo",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=40, corner_radius=8, width=200,
            command=self._on_browse_saldos,
        )
        self.sld_browse_btn.pack(pady=(14, 0))

        # ======== FILE INFO BAR (hidden initially) ========
        self.sld_file_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.sld_file_frame.pack(fill="x", pady=(0, 12))
        self.sld_file_frame.pack_forget()

        fi = ctk.CTkFrame(self.sld_file_frame, fg_color="transparent")
        fi.pack(fill="x", padx=16, pady=10)

        self.sld_file_icon = ctk.CTkLabel(
            fi, text="\u25cf", font=("Segoe UI", 12), text_color=ACCENT_GREEN
        )
        self.sld_file_icon.pack(side="left")

        self.sld_file_name = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 11, "bold"), text_color=TEXT_PRIMARY
        )
        self.sld_file_name.pack(side="left", padx=(6, 0))

        self.sld_file_info = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 10), text_color=TEXT_SECONDARY
        )
        self.sld_file_info.pack(side="left", padx=(12, 0))

        ctk.CTkButton(
            fi, text="Trocar", font=("Segoe UI", 10),
            fg_color=BG_INPUT, hover_color=BORDER_CARD,
            text_color=TEXT_SECONDARY, height=28, corner_radius=6, width=70,
            command=self._on_browse_saldos,
        ).pack(side="right")

        # ======== PREVIEW ========
        self.sld_preview_label = ctk.CTkLabel(
            content, text="Preview dos Dados",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.sld_preview_label.pack(fill="x", pady=(4, 8))

        self.sld_preview_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=12,
            border_width=1, border_color=BORDER_CARD
        )
        self.sld_preview_frame.pack(fill="x", pady=(0, 14))

        self.sld_preview_text = ctk.CTkTextbox(
            self.sld_preview_frame, font=("Consolas", 9),
            fg_color=BG_CARD, text_color=TEXT_PRIMARY,
            corner_radius=10, height=200, wrap="none"
        )
        self.sld_preview_text.pack(fill="x", padx=4, pady=4)
        self.sld_preview_text.insert("1.0", "  Nenhum arquivo carregado...")
        self.sld_preview_text.configure(state="disabled")

        # ======== ACTION BUTTONS ========
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 12))

        self.sld_process_btn = ctk.CTkButton(
            btn_frame, text="  Processar Saldos",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=self._darken(ACCENT_GREEN),
            height=44, corner_radius=10, state="disabled",
            command=self._on_process_saldos,
        )
        self.sld_process_btn.pack(side="left", padx=(0, 10))

        self.sld_open_btn = ctk.CTkButton(
            btn_frame, text="  Abrir Resultado",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=44, corner_radius=10, state="disabled",
            command=self._on_open_saldos_result,
        )
        self.sld_open_btn.pack(side="left", padx=(0, 10))

        # ======== STATUS BAR ========
        self.sld_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.sld_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.sld_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.sld_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.sld_status_dot.pack(side="left")

        self.sld_status_text = ctk.CTkLabel(
            si, text="  Aguardando arquivo...",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.sld_status_text.pack(side="left")

        # Botao Atualizar
        self._make_atualizar_btn(content, os.path.join(SALDOS_DIR, "BASE", "BASE ENVIAR.xlsx"))

        # Internal state
        self._sld_input_path = None
        self._sld_output_path = None

        return page

    # -----------------------------------------------------------------
    #  ENVIO SALDOS: Ações
    # -----------------------------------------------------------------
    def _on_browse_saldos(self):
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        path = filedialog.askopenfilename(
            title="Selecionar Planilha de Saldos",
            initialdir=downloads_dir if os.path.isdir(downloads_dir) else os.path.join(SALDOS_DIR, "BASE"),
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._sld_load_file(path)

    def _sld_load_file(self, path):
        self._sld_input_path = path
        self._sld_output_path = None

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            total_rows = 0
            total_cols = 0
            for ws in wb:
                total_rows += ws.max_row or 0
                total_cols = max(total_cols, ws.max_column or 0)
            wb.close()

            fname = os.path.basename(path)
            self.sld_file_name.configure(text=fname)

            # Verifica se o arquivo foi modificado hoje
            from datetime import datetime
            file_date = datetime.fromtimestamp(os.path.getmtime(path)).date()
            hoje = datetime.now().date()
            date_warning = ""
            if file_date < hoje:
                date_warning = f"  |  ATENCAO: arquivo de {file_date.strftime('%d/%m/%Y')}"
                messagebox.showwarning(
                    "Arquivo Desatualizado",
                    f"Este arquivo foi modificado em {file_date.strftime('%d/%m/%Y')}.\n"
                    f"Hoje e {hoje.strftime('%d/%m/%Y')}.\n\n"
                    f"Certifique-se de que esta usando o relatorio de saldos mais recente."
                )

            info = f"{len(sheets)} aba(s)  |  ~{total_rows} linhas  |  {total_cols} colunas{date_warning}"
            self.sld_file_info.configure(text=info)
            self.sld_file_frame.pack(fill="x", pady=(0, 12))

            # Update drop zone visual
            if file_date < hoje:
                self.sld_drop_title.configure(text=fname)
                self.sld_drop_sub.configure(text=f"Arquivo de {file_date.strftime('%d/%m/%Y')} - clique para trocar")
                self.sld_icon_label.configure(text_color=ACCENT_ORANGE, text="\u26a0")
                self.sld_drop_card.configure(border_color=ACCENT_ORANGE, fg_color="#fffbeb")
            else:
                self.sld_drop_title.configure(text=fname)
                self.sld_drop_sub.configure(text="Arquivo carregado - clique para trocar")
                self.sld_icon_label.configure(text_color=ACCENT_GREEN, text="\u2713")
                self.sld_drop_card.configure(border_color=ACCENT_GREEN, fg_color="#f0fff4")

            self._sld_show_preview(path)

            self.sld_process_btn.configure(state="normal")
            self.sld_open_btn.configure(state="disabled")

            self.sld_status_dot.configure(text_color=ACCENT_BLUE)
            self.sld_status_text.configure(
                text=f"  Arquivo carregado. Clique em 'Processar Saldos'."
            )

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo:\n{e}")

    def _sld_show_preview(self, path):
        try:
            self.sld_preview_text.configure(state="normal")
            self.sld_preview_text.delete("1.0", "end")

            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            for i, row in enumerate(ws.iter_rows(max_row=min(18, ws.max_row or 18), values_only=True)):
                cells = [str(c)[:50] if c is not None else "" for c in row]
                self.sld_preview_text.insert("end", "  |  ".join(cells) + "\n")
            wb.close()

            self.sld_preview_text.configure(state="disabled")
        except Exception:
            pass

    def _on_process_saldos(self):
        if not self._sld_input_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo primeiro.")
            return

        confirm = messagebox.askyesno(
            "Gerar Rascunhos no Outlook",
            "Isso vai criar rascunhos de email no Outlook para cada assessor "
            "com os saldos dos seus clientes.\n\n"
            "Os emails NÃO serão enviados, apenas salvos como rascunho.\n"
            "O Outlook precisa estar aberto. Deseja continuar?"
        )
        if not confirm:
            return

        self.sld_process_btn.configure(state="disabled")
        self.sld_status_dot.configure(text_color=ACCENT_ORANGE)
        self.sld_status_text.configure(text="  Processando saldos...")

        # Atualiza preview com log
        self.sld_preview_text.configure(state="normal")
        self.sld_preview_text.delete("1.0", "end")
        self.sld_preview_text.configure(state="disabled")

        threading.Thread(target=self._run_process_saldos, daemon=True).start()

    def _run_process_saldos(self):
        try:
            from processar_saldos import processar_saldos

            def log_callback(msg, tipo="info"):
                self.after(0, self._sld_append_log, msg, tipo)

            stats = processar_saldos(self._sld_input_path, callback=log_callback)

            if stats["criados"] > 0:
                self.after(0, self.sld_status_dot.configure, {"text_color": ACCENT_GREEN})
                self.after(0, self.sld_status_text.configure, {
                    "text": f"  Concluído - {stats['criados']} rascunhos criados no Outlook"
                })
                self.after(0, lambda: messagebox.showinfo(
                    "Rascunhos Criados",
                    f"{stats['criados']} rascunhos criados nos Rascunhos do Outlook!\n\n"
                    f"Erros: {stats['erros']}  |  Sem email: {stats['sem_email']}"
                ))
            elif stats["erros"] > 0:
                self.after(0, self.sld_status_dot.configure, {"text_color": ACCENT_RED})
                self.after(0, self.sld_status_text.configure, {
                    "text": f"  Erros no processamento - verifique o log"
                })
            else:
                self.after(0, self.sld_status_dot.configure, {"text_color": ACCENT_ORANGE})
                self.after(0, self.sld_status_text.configure, {
                    "text": f"  Nenhum rascunho criado (sem assessores com email)"
                })

        except Exception as e:
            err_msg = str(e)
            self.after(0, self.sld_status_dot.configure, {"text_color": ACCENT_RED})
            self.after(0, self.sld_status_text.configure, {"text": f"  Erro: {err_msg}"})
            self.after(0, lambda m=err_msg: messagebox.showerror("Erro", m))
        finally:
            self.after(0, self.sld_process_btn.configure, {"state": "normal"})

    def _sld_append_log(self, msg, tipo="info"):
        color_map = {"ok": ACCENT_GREEN, "error": ACCENT_RED, "skip": ACCENT_ORANGE, "info": TEXT_PRIMARY}
        self.sld_preview_text.configure(state="normal")
        self.sld_preview_text.insert("end", msg + "\n")
        self.sld_preview_text.see("end")
        self.sld_preview_text.configure(state="disabled")

    def _on_open_saldos_result(self):
        if self._sld_output_path and os.path.exists(self._sld_output_path):
            os.startfile(self._sld_output_path)

    def _on_open_saldos_folder(self):
        os.makedirs(SALDOS_DIR, exist_ok=True)
        os.startfile(SALDOS_DIR)

    # -----------------------------------------------------------------
    #  SEGUROS: Renovações Anuais
    # -----------------------------------------------------------------
    def _build_renovacoes_page(self):
        page = ctk.CTkFrame(self, fg_color=BG_PRIMARY, corner_radius=0)

        self._make_topbar(page, "Renovações Anuais", subtitle="Seguros")

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_PRIMARY, corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        content = ctk.CTkFrame(scroll, fg_color="transparent")
        content.pack(fill="x", padx=28, pady=20)

        # ======== DROP ZONE ========
        self.ren_drop_card = ctk.CTkFrame(
            content, fg_color="#f0f4ff", corner_radius=14,
            border_width=2, border_color=ACCENT_BLUE
        )
        self.ren_drop_card.pack(fill="x", pady=(0, 16))
        self.ren_drop_card.bind("<Button-1>", lambda e: self._on_browse_renovacoes())

        drop_inner = ctk.CTkFrame(self.ren_drop_card, fg_color="transparent")
        drop_inner.pack(fill="x", padx=30, pady=36)
        drop_inner.bind("<Button-1>", lambda e: self._on_browse_renovacoes())

        self.ren_icon_label = ctk.CTkLabel(
            drop_inner, text="\u26e8",
            font=("Segoe UI", 44), text_color=ACCENT_BLUE
        )
        self.ren_icon_label.pack()
        self.ren_icon_label.bind("<Button-1>", lambda e: self._on_browse_renovacoes())

        self.ren_drop_title = ctk.CTkLabel(
            drop_inner, text="Clique para selecionar a planilha de renovações",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY
        )
        self.ren_drop_title.pack(pady=(8, 2))
        self.ren_drop_title.bind("<Button-1>", lambda e: self._on_browse_renovacoes())

        self.ren_drop_sub = ctk.CTkLabel(
            drop_inner, text="Formatos aceitos:  .xlsx  .xls",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.ren_drop_sub.pack()
        self.ren_drop_sub.bind("<Button-1>", lambda e: self._on_browse_renovacoes())

        ren_btn_row = ctk.CTkFrame(drop_inner, fg_color="transparent")
        ren_btn_row.pack(pady=(14, 0))

        self.ren_browse_btn = ctk.CTkButton(
            ren_btn_row, text="  Selecionar Arquivo",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_BLUE, hover_color="#1555bb",
            height=40, corner_radius=8, width=200,
            command=self._on_browse_renovacoes,
        )
        self.ren_browse_btn.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            ren_btn_row, text="  \u2913  Baixar do SharePoint",
            font=("Segoe UI", 12, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            height=40, corner_radius=8, width=220,
            command=self._on_baixar_renovacoes,
        ).pack(side="left")

        # ======== FILE INFO BAR (hidden initially) ========
        self.ren_file_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.ren_file_frame.pack(fill="x", pady=(0, 12))
        self.ren_file_frame.pack_forget()

        fi = ctk.CTkFrame(self.ren_file_frame, fg_color="transparent")
        fi.pack(fill="x", padx=16, pady=10)

        self.ren_file_icon = ctk.CTkLabel(
            fi, text="\u25cf", font=("Segoe UI", 12), text_color=ACCENT_GREEN
        )
        self.ren_file_icon.pack(side="left")

        self.ren_file_name = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 11, "bold"), text_color=TEXT_PRIMARY
        )
        self.ren_file_name.pack(side="left", padx=(6, 0))

        self.ren_file_info = ctk.CTkLabel(
            fi, text="", font=("Segoe UI", 10), text_color=TEXT_SECONDARY
        )
        self.ren_file_info.pack(side="left", padx=(12, 0))

        ctk.CTkButton(
            fi, text="Trocar", font=("Segoe UI", 10),
            fg_color=BG_INPUT, hover_color=BORDER_CARD,
            text_color=TEXT_SECONDARY, height=28, corner_radius=6, width=70,
            command=self._on_browse_renovacoes,
        ).pack(side="right")

        # ======== PREVIEW ========
        self.ren_preview_label = ctk.CTkLabel(
            content, text="Preview dos Dados",
            font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY, anchor="w"
        )
        self.ren_preview_label.pack(fill="x", pady=(4, 8))

        self.ren_preview_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=12,
            border_width=1, border_color=BORDER_CARD
        )
        self.ren_preview_frame.pack(fill="x", pady=(0, 14))

        self.ren_preview_text = ctk.CTkTextbox(
            self.ren_preview_frame, font=("Consolas", 9),
            fg_color=BG_CARD, text_color=TEXT_PRIMARY,
            corner_radius=10, height=200, wrap="none"
        )
        self.ren_preview_text.pack(fill="x", padx=4, pady=4)
        self.ren_preview_text.insert("1.0", "  Nenhum arquivo carregado...")
        self.ren_preview_text.configure(state="disabled")

        # ======== ACTION BUTTONS ========
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 12))

        self.ren_process_btn = ctk.CTkButton(
            btn_frame, text="  Processar Renovações",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=self._darken(ACCENT_GREEN),
            height=44, corner_radius=10, state="disabled",
            command=self._on_process_renovacoes,
        )
        self.ren_process_btn.pack(side="left", padx=(0, 10))

        # ======== STATUS BAR ========
        self.ren_status_frame = ctk.CTkFrame(
            content, fg_color=BG_CARD, corner_radius=10,
            border_width=1, border_color=BORDER_CARD
        )
        self.ren_status_frame.pack(fill="x", pady=(0, 8))

        si = ctk.CTkFrame(self.ren_status_frame, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=10)

        self.ren_status_dot = ctk.CTkLabel(
            si, text="\u25cf", font=("Segoe UI", 11), text_color=TEXT_TERTIARY
        )
        self.ren_status_dot.pack(side="left")

        self.ren_status_text = ctk.CTkLabel(
            si, text="  Aguardando arquivo...",
            font=("Segoe UI", 11), text_color=TEXT_SECONDARY
        )
        self.ren_status_text.pack(side="left")

        # Botao Atualizar — abre pasta Seguros no SharePoint
        atualizar_frame = ctk.CTkFrame(content, fg_color="transparent")
        atualizar_frame.pack(fill="x", pady=(24, 8))
        ctk.CTkButton(
            atualizar_frame, text="\u21bb  Atualizar",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_WHITE, height=42, corner_radius=10,
            command=lambda: self._safe_open_folder(SEGUROS_DIR)
        ).pack(fill="x")

        # Internal state
        self._ren_input_path = None
        self._ren_output_path = None

        return page

    def _safe_open_folder(self, path):
        try:
            os.makedirs(path, exist_ok=True)
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}")

    def _on_browse_renovacoes(self):
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        path = filedialog.askopenfilename(
            title="Selecionar Planilha de Renovações",
            initialdir=downloads_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._ren_load_file(path)

    def _on_baixar_renovacoes(self):
        """Copia arquivos de renovações do SharePoint para Downloads e carrega."""
        import shutil
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        if not os.path.isdir(SEGUROS_DIR):
            messagebox.showwarning("Aviso", f"Pasta de Seguros não encontrada:\n{SEGUROS_DIR}")
            return

        found = []
        try:
            for f in os.listdir(SEGUROS_DIR):
                if f.lower().endswith(('.xlsx', '.xls')) and 'renov' in f.lower():
                    found.append(f)
        except (PermissionError, OSError) as e:
            messagebox.showerror("Erro", f"Sem acesso à pasta do SharePoint:\n{e}")
            return

        if not found:
            messagebox.showinfo("Info", "Nenhuma planilha de renovações encontrada na pasta do SharePoint.")
            return

        copied = []
        for fname in found:
            src = os.path.join(SEGUROS_DIR, fname)
            dst = os.path.join(downloads_dir, fname)
            try:
                shutil.copy2(src, dst)
                copied.append(dst)
            except (PermissionError, OSError) as e:
                messagebox.showwarning("Aviso", f"Não foi possível copiar {fname}:\n{e}\n\nTente abrir o arquivo no SharePoint primeiro para baixá-lo.")

        if not copied:
            return

        if len(copied) == 1:
            self._ren_load_file(copied[0])
        else:
            messagebox.showinfo(
                "Arquivos Copiados",
                f"{len(copied)} planilhas copiadas para Downloads.\nSelecione qual deseja processar."
            )
            path = filedialog.askopenfilename(
                title="Selecionar Planilha de Renovações",
                initialdir=downloads_dir,
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if path:
                self._ren_load_file(path)

    def _ren_load_file(self, path):
        self._ren_input_path = path
        self._ren_output_path = None

        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            total_rows = 0
            total_cols = 0
            for ws in wb:
                total_rows += ws.max_row or 0
                total_cols = max(total_cols, ws.max_column or 0)
            wb.close()

            fname = os.path.basename(path)
            self.ren_file_name.configure(text=fname)

            info = f"{len(sheets)} aba(s)  |  ~{total_rows} linhas  |  {total_cols} colunas"
            self.ren_file_info.configure(text=info)
            self.ren_file_frame.pack(fill="x", pady=(0, 12))

            self.ren_drop_title.configure(text=fname)
            self.ren_drop_sub.configure(text="Arquivo carregado - clique para trocar")
            self.ren_icon_label.configure(text_color=ACCENT_GREEN, text="\u2713")
            self.ren_drop_card.configure(border_color=ACCENT_GREEN, fg_color="#f0fff4")

            self._ren_show_preview(path)

            self.ren_process_btn.configure(state="normal")

            self.ren_status_dot.configure(text_color=ACCENT_BLUE)
            self.ren_status_text.configure(
                text=f"  Arquivo carregado. Clique em 'Processar Renovações'."
            )

        except PermissionError:
            messagebox.showerror(
                "Permissão Negada",
                f"Sem permissão para ler o arquivo:\n{path}\n\n"
                "Se o arquivo está no SharePoint/OneDrive, use o botão "
                "'Baixar do SharePoint' para copiá-lo para Downloads primeiro."
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo:\n{e}")

    def _ren_show_preview(self, path):
        try:
            self.ren_preview_text.configure(state="normal")
            self.ren_preview_text.delete("1.0", "end")

            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            for i, row in enumerate(ws.iter_rows(max_row=min(18, ws.max_row or 18), values_only=True)):
                cells = [str(c)[:50] if c is not None else "" for c in row]
                self.ren_preview_text.insert("end", "  |  ".join(cells) + "\n")
            wb.close()

            self.ren_preview_text.configure(state="disabled")
        except Exception:
            pass

    def _on_process_renovacoes(self):
        if not self._ren_input_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo primeiro.")
            return

        confirm = messagebox.askyesno(
            "Gerar Rascunhos no Outlook",
            "Isso vai criar rascunhos de email no Outlook para cada assessor "
            "com as renovações anuais de seguros dos seus clientes.\n\n"
            "Os emails NÃO serão enviados, apenas salvos como rascunho.\n"
            "O Outlook precisa estar aberto. Deseja continuar?"
        )
        if not confirm:
            return

        self.ren_process_btn.configure(state="disabled")
        self.ren_status_dot.configure(text_color=ACCENT_ORANGE)
        self.ren_status_text.configure(text="  Processando renovações...")

        self.ren_preview_text.configure(state="normal")
        self.ren_preview_text.delete("1.0", "end")
        self.ren_preview_text.configure(state="disabled")

        threading.Thread(target=self._run_process_renovacoes, daemon=True).start()

    def _run_process_renovacoes(self):
        try:
            from processar_renovacoes import processar_renovacoes

            def log_callback(msg, tipo="info"):
                self.after(0, self._ren_append_log, msg, tipo)

            stats = processar_renovacoes(self._ren_input_path, callback=log_callback)

            if stats["criados"] > 0:
                self.after(0, self.ren_status_dot.configure, {"text_color": ACCENT_GREEN})
                self.after(0, self.ren_status_text.configure, {
                    "text": f"  Concluído - {stats['criados']} rascunhos criados no Outlook"
                })
                self.after(0, lambda: messagebox.showinfo(
                    "Rascunhos Criados",
                    f"{stats['criados']} rascunhos criados nos Rascunhos do Outlook!\n\n"
                    f"Erros: {stats['erros']}  |  Sem email: {stats['sem_email']}"
                ))
            elif stats["erros"] > 0:
                self.after(0, self.ren_status_dot.configure, {"text_color": ACCENT_RED})
                self.after(0, self.ren_status_text.configure, {
                    "text": f"  Erros no processamento - verifique o log"
                })
            else:
                self.after(0, self.ren_status_dot.configure, {"text_color": ACCENT_ORANGE})
                self.after(0, self.ren_status_text.configure, {
                    "text": f"  Nenhum rascunho criado (sem assessores com email)"
                })

        except Exception as e:
            err_msg = str(e)
            self.after(0, self.ren_status_dot.configure, {"text_color": ACCENT_RED})
            self.after(0, self.ren_status_text.configure, {"text": f"  Erro: {err_msg}"})
            self.after(0, lambda m=err_msg: messagebox.showerror("Erro", m))
        finally:
            self.after(0, self.ren_process_btn.configure, {"state": "normal"})

    def _ren_append_log(self, msg, tipo="info"):
        self.ren_preview_text.configure(state="normal")
        self.ren_preview_text.insert("end", msg + "\n")
        self.ren_preview_text.see("end")
        self.ren_preview_text.configure(state="disabled")

    # -----------------------------------------------------------------
    #  CONSÓRCIO: Ações
    # -----------------------------------------------------------------
    def _parse_number(self, text):
        s = text.strip().replace("%", "").replace("R$", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)

    def _on_calcular_consorcio(self):
        try:
            nome = self.cs_nome.get().strip()
            assessor = self.cs_assessor.get().strip()
            tipo_bem = self.cs_tipo_bem.get()
            administradora = self.cs_administradora.get().strip()
            valor_carta = self._parse_number(self.cs_valor_carta.get())
            prazo = int(self.cs_prazo.get())
            taxa_adm = self._parse_number(self.cs_taxa_adm.get())
            fundo_reserva = self._parse_number(self.cs_fundo_reserva.get())
            seguro = self._parse_number(self.cs_seguro.get())
            prazo_contemp = int(self._parse_number(self.cs_prazo_contemplação.get()))
            lance_livre_pct = self._parse_number(self.cs_lance_livre.get())
            lance_embutido_pct = self._parse_number(self.cs_lance_embutido.get())
            correção_anual = self._parse_number(self.cs_correção_anual.get())
            tipo_correção = self.cs_tipo_correção.get()
            índice_correção = self.cs_índice_correção.get()
        except (ValueError, AttributeError):
            messagebox.showerror("Erro", "Preencha todos os campos numericos corretamente.")
            return

        # Parcela reduzida
        pr_text = self.cs_parcela_reduzida.get()
        if "50%" in pr_text:
            parcela_red_pct = 50
        elif "70%" in pr_text:
            parcela_red_pct = 70
        else:
            parcela_red_pct = 100

        # Validações
        if valor_carta <= 0 or prazo <= 0:
            messagebox.showerror("Erro", "Valor da carta e prazo devem ser maiores que zero.")
            return
        if prazo_contemp <= 0 or prazo_contemp > prazo:
            messagebox.showerror("Erro", f"Prazo de contemplação deve ser entre 1 e {prazo}.")
            return
        if lance_embutido_pct > 30:
            messagebox.showwarning("Aviso",
                "Lance embutido acima de 30% e incomum.\nVerifique com a administradora.")

        # === CÁLCULOS BASE (mensais fixos sobre prazo total) ===
        fc_integral = valor_carta / prazo
        ta_mensal = (valor_carta * taxa_adm / 100) / prazo
        fr_mensal = (valor_carta * fundo_reserva / 100) / prazo
        sg_mensal = (valor_carta * seguro / 100) / prazo
        taxas_mensais = ta_mensal + fr_mensal + sg_mensal

        # === FASE 1: Antes da contemplação ===
        fc_reduzido = fc_integral * (parcela_red_pct / 100)
        parcela_fase1 = fc_reduzido + taxas_mensais
        fundo_pago_fase1 = fc_reduzido * prazo_contemp

        # === LANCES (no momento da contemplação) ===
        lance_livre_valor = valor_carta * (lance_livre_pct / 100)
        lance_embutido_valor = valor_carta * (lance_embutido_pct / 100)
        lance_total = lance_livre_valor + lance_embutido_valor
        carta_líquida = valor_carta - lance_embutido_valor

        # === FASE 2: Após contemplação ===
        meses_restantes = prazo - prazo_contemp
        fundo_restante = max(0, valor_carta - fundo_pago_fase1 - lance_total)
        fc_fase2 = fundo_restante / meses_restantes if meses_restantes > 0 else 0
        parcela_fase2 = fc_fase2 + taxas_mensais

        # === CORREÇÃO ANUAL ===
        taxa_corr = correção_anual / 100  # ex: 5.5% -> 0.055

        # Calcular parcelas corrigidas mês a mês (fator composto anual)
        desemb_fase1_corr = 0.0
        desemb_fase2_corr = 0.0
        parcela_f1_final = parcela_fase1  # última parcela fase 1 (corrigida)
        parcela_f2_final = parcela_fase2  # última parcela fase 2 (corrigida)

        for mes in range(1, prazo + 1):
            fator = (1 + taxa_corr) ** ((mes - 1) // 12)
            if mes <= prazo_contemp:
                p_corr = parcela_fase1 * fator
                desemb_fase1_corr += p_corr
                parcela_f1_final = p_corr
            else:
                p_corr = parcela_fase2 * fator
                desemb_fase2_corr += p_corr
                parcela_f2_final = p_corr

        # === TOTAIS ===
        desemb_fase1 = parcela_fase1 * prazo_contemp
        desemb_fase2 = parcela_fase2 * meses_restantes if meses_restantes > 0 else 0
        total_desembolsado = desemb_fase1 + lance_livre_valor + desemb_fase2
        total_taxas = taxas_mensais * prazo

        # Totais com correção
        total_desemb_corr = desemb_fase1_corr + lance_livre_valor + desemb_fase2_corr
        tem_correção = correção_anual > 0

        # === CUSTO EFETIVO ===
        from gerar_pdf_consorcio import _calc_custo_efetivo
        total_ref = total_desemb_corr if tem_correção else total_desembolsado
        ce_mensal, ce_anual, ce_pct = _calc_custo_efetivo(carta_líquida, total_ref, prazo)
        custo_total = total_ref - carta_líquida

        # Salvar dados para o PDF
        self._consórcio_dados = {
            "cliente_nome": nome or "Cliente",
            "assessor": assessor,
            "tipo_bem": tipo_bem,
            "administradora": administradora,
            "valor_carta": valor_carta,
            "prazo_meses": prazo,
            "taxa_adm": taxa_adm,
            "fundo_reserva": fundo_reserva,
            "seguro": seguro,
            "prazo_contemplação": prazo_contemp,
            "parcela_reduzida_pct": parcela_red_pct,
            "lance_livre_pct": lance_livre_pct,
            "lance_embutido_pct": lance_embutido_pct,
            "lance_livre_valor": lance_livre_valor,
            "lance_embutido_valor": lance_embutido_valor,
            "carta_líquida": carta_líquida,
            "parcela_fase1": parcela_fase1,
            "parcela_fase2": parcela_fase2,
            "fc_integral": fc_integral,
            "fc_reduzido": fc_reduzido,
            "fc_fase2": fc_fase2,
            "ta_mensal": ta_mensal,
            "fr_mensal": fr_mensal,
            "sg_mensal": sg_mensal,
            "desemb_fase1": desemb_fase1,
            "desemb_fase2": desemb_fase2,
            "total_desembolsado": total_desembolsado,
            "total_taxas": total_taxas,
            "correção_anual": correção_anual,
            "tipo_correção": tipo_correção,
            "índice_correção": índice_correção,
        }

        # ======== ATUALIZAR UI DE RESULTADO ========
        for w in self.cs_result_inner.winfo_children():
            w.destroy()

        def _result_row(parent, label, value, color=TEXT_PRIMARY, bold=False):
            r = ctk.CTkFrame(parent, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text=label, font=("Segoe UI", 10),
                         text_color=TEXT_SECONDARY).pack(side="left")
            f = ("Segoe UI", 11, "bold") if bold else ("Segoe UI", 11)
            ctk.CTkLabel(r, text=value, font=f,
                         text_color=color).pack(side="right")

        def _dot_row(parent, label, value, color):
            r = ctk.CTkFrame(parent, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text="\u25cf", font=("Segoe UI", 9),
                         text_color=color).pack(side="left")
            ctk.CTkLabel(r, text=f"  {label}", font=("Segoe UI", 10),
                         text_color=TEXT_SECONDARY).pack(side="left")
            ctk.CTkLabel(r, text=value, font=("Segoe UI", 11, "bold"),
                         text_color=TEXT_PRIMARY).pack(side="right")

        def _separator(parent):
            ctk.CTkFrame(parent, fg_color=BORDER_LIGHT, height=1).pack(fill="x", pady=8)

        # ------ FASE 1 ------
        red_label = f"  (reduzida {parcela_red_pct}%)" if parcela_red_pct < 100 else ""
        f1 = ctk.CTkFrame(self.cs_result_inner, fg_color=ACCENT_GREEN, corner_radius=10)
        f1.pack(fill="x", pady=(0, 6))
        f1i = ctk.CTkFrame(f1, fg_color="transparent")
        f1i.pack(fill="x", padx=20, pady=12)
        ctk.CTkLabel(f1i, text=f"FASE 1  -  Antes da Contemplação{red_label}",
                     font=("Segoe UI", 9), text_color="#80c0a0").pack(anchor="w")
        ctk.CTkLabel(f1i, text=f"{fmt_currency(parcela_fase1)} /mes   x   {prazo_contemp} meses",
                     font=("Segoe UI", 22, "bold"), text_color=TEXT_WHITE).pack(anchor="w", pady=(2, 0))
        if tem_correção:
            ctk.CTkLabel(f1i, text=f"Parcela final estimada (corrigida): {fmt_currency(parcela_f1_final)}",
                         font=("Segoe UI", 9), text_color="#b0e0c8").pack(anchor="w")

        fc_lbl = "Fundo Comum" + (f" ({parcela_red_pct}%)" if parcela_red_pct < 100 else "")
        _dot_row(self.cs_result_inner, fc_lbl, fmt_currency(fc_reduzido), ACCENT_BLUE)
        _dot_row(self.cs_result_inner, "Taxa de Administração", fmt_currency(ta_mensal), ACCENT_ORANGE)
        _dot_row(self.cs_result_inner, "Fundo de Reserva", fmt_currency(fr_mensal), ACCENT_TEAL)
        _dot_row(self.cs_result_inner, "Seguro Prestamista", fmt_currency(sg_mensal), ACCENT_PURPLE)

        _separator(self.cs_result_inner)

        # ------ CONTEMPLAÇÃO / LANCES ------
        if lance_livre_valor > 0 or lance_embutido_valor > 0:
            cf = ctk.CTkFrame(self.cs_result_inner, fg_color=ACCENT_BLUE, corner_radius=10)
            cf.pack(fill="x", pady=(0, 6))
            ci = ctk.CTkFrame(cf, fg_color="transparent")
            ci.pack(fill="x", padx=20, pady=12)
            ctk.CTkLabel(ci, text=f"CONTEMPLAÇÃO  -  Mês {prazo_contemp}",
                         font=("Segoe UI", 9), text_color="#a0c0f0").pack(anchor="w")

            if lance_livre_valor > 0:
                r = ctk.CTkFrame(ci, fg_color="transparent")
                r.pack(fill="x", pady=1)
                ctk.CTkLabel(r, text="Lance Livre (recursos próprios)",
                             font=("Segoe UI", 10), text_color="#c0d8f8").pack(side="left")
                ctk.CTkLabel(r, text=fmt_currency(lance_livre_valor),
                             font=("Segoe UI", 12, "bold"), text_color=TEXT_WHITE).pack(side="right")

            if lance_embutido_valor > 0:
                r = ctk.CTkFrame(ci, fg_color="transparent")
                r.pack(fill="x", pady=1)
                ctk.CTkLabel(r, text="Lance Embutido (descontado da carta)",
                             font=("Segoe UI", 10), text_color="#c0d8f8").pack(side="left")
                ctk.CTkLabel(r, text=fmt_currency(lance_embutido_valor),
                             font=("Segoe UI", 12, "bold"), text_color=TEXT_WHITE).pack(side="right")

            r = ctk.CTkFrame(ci, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text="Carta de Crédito Líquida",
                         font=("Segoe UI", 10, "bold"), text_color="#c0d8f8").pack(side="left")
            ctk.CTkLabel(r, text=fmt_currency(carta_líquida),
                         font=("Segoe UI", 14, "bold"), text_color=TEXT_WHITE).pack(side="right")

            _separator(self.cs_result_inner)

        # ------ FASE 2 ------
        if meses_restantes > 0:
            f2 = ctk.CTkFrame(self.cs_result_inner, fg_color="#005c3d", corner_radius=10)
            f2.pack(fill="x", pady=(0, 6))
            f2i = ctk.CTkFrame(f2, fg_color="transparent")
            f2i.pack(fill="x", padx=20, pady=12)
            ctk.CTkLabel(f2i, text="FASE 2  -  Após Contemplação  (parcela reajustada)",
                         font=("Segoe UI", 9), text_color="#80c0a0").pack(anchor="w")
            ctk.CTkLabel(f2i, text=f"{fmt_currency(parcela_fase2)} /mes   x   {meses_restantes} meses",
                         font=("Segoe UI", 22, "bold"), text_color=TEXT_WHITE).pack(anchor="w", pady=(2, 0))
            if tem_correção:
                ctk.CTkLabel(f2i, text=f"Parcela final estimada (corrigida): {fmt_currency(parcela_f2_final)}",
                             font=("Segoe UI", 9), text_color="#b0e0c8").pack(anchor="w")

            _dot_row(self.cs_result_inner, "Fundo Comum (reajustado)", fmt_currency(fc_fase2), ACCENT_BLUE)
            _dot_row(self.cs_result_inner, "Taxa de Administração", fmt_currency(ta_mensal), ACCENT_ORANGE)
            _dot_row(self.cs_result_inner, "Fundo de Reserva", fmt_currency(fr_mensal), ACCENT_TEAL)
            _dot_row(self.cs_result_inner, "Seguro Prestamista", fmt_currency(sg_mensal), ACCENT_PURPLE)

        _separator(self.cs_result_inner)

        # ------ TOTAIS ------
        _result_row(self.cs_result_inner, "Valor da Carta de Crédito", fmt_currency(valor_carta), bold=True)
        if lance_embutido_valor > 0:
            _result_row(self.cs_result_inner, "Carta Líquida (apos lance embutido)", fmt_currency(carta_líquida), ACCENT_RED, True)
        _result_row(self.cs_result_inner, f"Desembolso Fase 1 ({prazo_contemp} meses)", fmt_currency(desemb_fase1))
        if lance_livre_valor > 0:
            _result_row(self.cs_result_inner, "Lance Livre (desembolso próprio)", fmt_currency(lance_livre_valor))
        if meses_restantes > 0:
            _result_row(self.cs_result_inner, f"Desembolso Fase 2 ({meses_restantes} meses)", fmt_currency(desemb_fase2))
        _result_row(self.cs_result_inner, "Total Desembolsado (sem correção)", fmt_currency(total_desembolsado), ACCENT_GREEN, True)
        if tem_correção:
            corr_label = f"{tipo_correção} {fmt_pct(correção_anual)} a.a."
            if tipo_correção == "Pós-fixado":
                corr_label += f" ({índice_correção})"
            _result_row(self.cs_result_inner, f"Total Corrigido ({corr_label})", fmt_currency(total_desemb_corr), ACCENT_RED, True)
        _result_row(self.cs_result_inner, "Total em Taxas e Encargos", fmt_currency(total_taxas))

        _separator(self.cs_result_inner)

        # ------ RESUMO DA OPERAÇÃO ------
        ro = ctk.CTkFrame(self.cs_result_inner, fg_color=ACCENT_GREEN, corner_radius=10)
        ro.pack(fill="x", pady=(0, 6))
        roi = ctk.CTkFrame(ro, fg_color="transparent")
        roi.pack(fill="x", padx=20, pady=12)
        ctk.CTkLabel(roi, text="CUSTO EFETIVO DO CONSÓRCIO",
                     font=("Segoe UI", 9), text_color="#80c0a0").pack(anchor="w")
        ctk.CTkLabel(roi, text=f"{fmt_pct(ce_pct)} sobre o crédito",
                     font=("Segoe UI", 20, "bold"), text_color=TEXT_WHITE).pack(anchor="w", pady=(2, 0))
        ctk.CTkLabel(roi, text=f"{fmt_pct(ce_anual * 100)} a.a.  |  {fmt_pct(ce_mensal * 100)} a.m.",
                     font=("Segoe UI", 12, "bold"), text_color="#b0e0c8").pack(anchor="w")

        _separator(self.cs_result_inner)

        _result_row(self.cs_result_inner, "Crédito Recebido (carta líquida)", fmt_currency(carta_líquida), ACCENT_BLUE, True)
        _result_row(self.cs_result_inner, "Custo Total da Operação", fmt_currency(custo_total), ACCENT_RED, True)
        relação = total_ref / carta_líquida if carta_líquida > 0 else 0
        _result_row(self.cs_result_inner, "Relação Custo / Crédito", f"{relação:.2f}x".replace(".", ","))
        _result_row(self.cs_result_inner, "Parcela Media", fmt_currency(total_ref / prazo if prazo > 0 else 0))

        # Habilitar PDF, Relatório e Email
        self.cs_btn_pdf.configure(state="normal")
        self.cs_btn_relatório.configure(state="normal")
        self.cs_btn_email.configure(state="normal")

    def _on_gerar_pdf_consorcio(self):
        if not hasattr(self, '_consórcio_dados') or not self._consórcio_dados:
            messagebox.showwarning("Aviso", "Calcule a simulação primeiro.")
            return
        try:
            filepath = generate_consorcio_pdf(self._consórcio_dados, CONSÓRCIO_OUTPUT_DIR)
            messagebox.showinfo("PDF Gerado",
                                f"PDF salvo com sucesso!\n\n{os.path.basename(filepath)}")
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{e}")

    def _on_gerar_relatorio_consorcio(self):
        if not hasattr(self, '_consórcio_dados') or not self._consórcio_dados:
            messagebox.showwarning("Aviso", "Calcule a simulação primeiro.")
            return
        try:
            filepath = generate_relatorio_consorcio(self._consórcio_dados, CONSÓRCIO_OUTPUT_DIR)
            messagebox.showinfo("Relatório Gerado",
                                f"Relatório explicativo salvo!\n\n{os.path.basename(filepath)}")
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{e}")

    def _on_abrir_email_consorcio(self):
        if not hasattr(self, '_consórcio_dados') or not self._consórcio_dados:
            messagebox.showwarning("Aviso", "Calcule a simulação primeiro.")
            return

        d = self._consórcio_dados
        from gerar_pdf_consorcio import _calc, fmt_currency as fc2, fmt_pct as fp2
        c = _calc(d)

        # Janela de composição de email
        win = ctk.CTkToplevel(self)
        win.title("Enviar Email - Consórcio")
        win.geometry("680x720")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        win.configure(fg_color=BG_PRIMARY)

        # Header
        hdr = ctk.CTkFrame(win, fg_color=ACCENT_GREEN, height=50, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  Enviar Simulação por Email",
                     font=("Segoe UI", 16, "bold"), text_color=TEXT_WHITE).pack(side="left", padx=20)

        body = ctk.CTkFrame(win, fg_color=BG_PRIMARY)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # -- Destinatario --
        ctk.CTkLabel(body, text="Para (email do cliente)", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SECONDARY).pack(anchor="w")
        email_to = ctk.CTkEntry(body, placeholder_text="cliente@email.com", height=38,
                                corner_radius=8, fg_color=BG_INPUT,
                                border_width=1, border_color=BORDER_LIGHT,
                                font=("Segoe UI", 11))
        email_to.pack(fill="x", pady=(2, 10))

        # -- CC --
        ctk.CTkLabel(body, text="CC (opcional)", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SECONDARY).pack(anchor="w")
        email_cc = ctk.CTkEntry(body, placeholder_text="copia@email.com", height=38,
                                corner_radius=8, fg_color=BG_INPUT,
                                border_width=1, border_color=BORDER_LIGHT,
                                font=("Segoe UI", 11))
        email_cc.pack(fill="x", pady=(2, 10))

        # -- Assunto --
        nome_cli = d.get("cliente_nome", "Cliente")
        tipo = d.get("tipo_bem", "")
        assunto_default = f"Simulação de Consórcio - {nome_cli} - {tipo}"

        ctk.CTkLabel(body, text="Assunto", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SECONDARY).pack(anchor="w")
        email_assunto = ctk.CTkEntry(body, height=38, corner_radius=8, fg_color=BG_INPUT,
                                     border_width=1, border_color=BORDER_LIGHT,
                                     font=("Segoe UI", 11))
        email_assunto.pack(fill="x", pady=(2, 10))
        email_assunto.insert(0, assunto_default)

        # -- Anexos --
        ctk.CTkLabel(body, text="Anexos", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SECONDARY).pack(anchor="w", pady=(4, 4))

        anexo_frame = ctk.CTkFrame(body, fg_color=BG_CARD, corner_radius=10,
                                   border_width=1, border_color=BORDER_CARD)
        anexo_frame.pack(fill="x", pady=(0, 10))
        af = ctk.CTkFrame(anexo_frame, fg_color="transparent")
        af.pack(fill="x", padx=16, pady=12)

        chk_proposta_var = ctk.BooleanVar(value=True)
        chk_proposta = ctk.CTkCheckBox(
            af, text="  Proposta de Simulação (PDF)",
            variable=chk_proposta_var, font=("Segoe UI", 11),
            fg_color=ACCENT_GREEN, hover_color=BG_SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY, corner_radius=4,
        )
        chk_proposta.pack(anchor="w", pady=2)

        chk_relatório_var = ctk.BooleanVar(value=True)
        chk_relatório = ctk.CTkCheckBox(
            af, text="  Relatório Explicativo (PDF)",
            variable=chk_relatório_var, font=("Segoe UI", 11),
            fg_color=ACCENT_PURPLE, hover_color="#6a4daa",
            text_color=TEXT_PRIMARY, corner_radius=4,
        )
        chk_relatório.pack(anchor="w", pady=2)

        ctk.CTkLabel(af, text="Os PDFs serão gerados automaticamente ao enviar",
                     font=("Segoe UI", 9), text_color=TEXT_TERTIARY).pack(anchor="w", pady=(4, 0))

        # -- Preview resumo --
        ctk.CTkLabel(body, text="Resumo no corpo do email", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SECONDARY).pack(anchor="w", pady=(4, 4))

        prev = ctk.CTkFrame(body, fg_color=BG_CARD, corner_radius=10,
                            border_width=1, border_color=BORDER_CARD)
        prev.pack(fill="x", pady=(0, 10))
        pi = ctk.CTkFrame(prev, fg_color="transparent")
        pi.pack(fill="x", padx=16, pady=12)

        preview_items = [
            ("Carta de Crédito", fc2(d.get("valor_carta", 0))),
            ("Administradora", d.get("administradora", "-")),
            ("Prazo", f"{d.get('prazo_meses', 0)} meses"),
            ("Parcela Fase 1", f"{fc2(d.get('parcela_fase1', 0))}/mes"),
            ("Custo Efetivo a.a.", fp2(c["ce_anual"] * 100)),
        ]
        for label, val in preview_items:
            r = ctk.CTkFrame(pi, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text=label, font=("Segoe UI", 10),
                         text_color=TEXT_SECONDARY).pack(side="left")
            ctk.CTkLabel(r, text=val, font=("Segoe UI", 10, "bold"),
                         text_color=TEXT_PRIMARY).pack(side="right")

        # -- Botoes --
        btn_f = ctk.CTkFrame(body, fg_color="transparent")
        btn_f.pack(fill="x", pady=(8, 0))

        def _do_enviar():
            to = email_to.get().strip()
            if not to:
                messagebox.showwarning("Aviso", "Preencha o email do destinatario.", parent=win)
                return
            cc = email_cc.get().strip()
            assunto = email_assunto.get().strip()
            anexar_prop = chk_proposta_var.get()
            anexar_rel = chk_relatório_var.get()

            if not anexar_prop and not anexar_rel:
                messagebox.showwarning("Aviso", "Selecione ao menos um anexo.", parent=win)
                return

            try:
                import win32com.client as win32
                import pythoncom
                pythoncom.CoInitialize()

                outlook = win32.Dispatch("Outlook.Application")
            except Exception as e:
                messagebox.showerror("Erro",
                    f"Não foi possível conectar ao Outlook.\nVerifique se está aberto.\n\n{e}",
                    parent=win)
                return

            try:
                # Gerar PDFs necessários
                anexos = []
                if anexar_prop:
                    prop_path = generate_consorcio_pdf(d, CONSÓRCIO_OUTPUT_DIR)
                    anexos.append(prop_path)
                if anexar_rel:
                    rel_path = generate_relatorio_consorcio(d, CONSÓRCIO_OUTPUT_DIR)
                    anexos.append(rel_path)

                # Montar email
                dados_resumo = {
                    "tipo_bem": d.get("tipo_bem", ""),
                    "valor_carta_fmt": fc2(d.get("valor_carta", 0)),
                    "administradora": d.get("administradora", "-"),
                    "prazo": str(d.get("prazo_meses", 0)),
                    "parcela_f1": fc2(d.get("parcela_fase1", 0)),
                    "ce_anual": fp2(c["ce_anual"] * 100),
                }
                html_body = build_email_consorcio(nome_cli, dados_resumo)

                mail = outlook.CreateItem(0)
                mail.To = to
                if cc:
                    mail.CC = cc
                mail.Subject = assunto
                mail.HTMLBody = html_body

                for a in anexos:
                    mail.Attachments.Add(os.path.abspath(a))

                mail.Save()

                n_anexos = len(anexos)
                pythoncom.CoUninitialize()
                win.destroy()
                messagebox.showinfo("Email Criado",
                    f"Rascunho salvo no Outlook!\n\n"
                    f"Para: {to}\n"
                    f"{n_anexos} anexo(s) incluido(s).\n\n"
                    f"Abra o Outlook e revise na pasta Rascunhos.")

            except Exception as e:
                pythoncom.CoUninitialize()
                messagebox.showerror("Erro", f"Erro ao criar email:\n{e}", parent=win)

        ctk.CTkButton(
            btn_f, text="  Enviar",
            font=("Segoe UI", 13, "bold"), fg_color=ACCENT_GREEN,
            hover_color=BG_SIDEBAR_HOVER, height=44, corner_radius=10,
            command=_do_enviar,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_f, text="  Cancelar",
            font=("Segoe UI", 12), fg_color="#6b7280",
            hover_color="#555d6a", height=44, corner_radius=10,
            command=win.destroy,
        ).pack(side="left")

    def _on_limpar_consorcio(self):
        self.cs_nome.delete(0, "end")
        self.cs_assessor.delete(0, "end")
        self.cs_tipo_bem.set("Imovel")
        self.cs_administradora.delete(0, "end")
        self.cs_valor_carta.delete(0, "end")
        self.cs_prazo.set("120")
        self.cs_taxa_adm.delete(0, "end")
        self.cs_taxa_adm.insert(0, "18")
        self.cs_fundo_reserva.delete(0, "end")
        self.cs_fundo_reserva.insert(0, "2")
        self.cs_seguro.delete(0, "end")
        self.cs_seguro.insert(0, "0")
        self.cs_prazo_contemplação.delete(0, "end")
        self.cs_prazo_contemplação.insert(0, "60")
        self.cs_parcela_reduzida.set("100% (Integral)")
        self.cs_lance_livre.delete(0, "end")
        self.cs_lance_livre.insert(0, "0")
        self.cs_lance_embutido.delete(0, "end")
        self.cs_lance_embutido.insert(0, "0")
        self.cs_correção_anual.delete(0, "end")
        self.cs_correção_anual.insert(0, "0")
        self.cs_tipo_correção.set("Pós-fixado")
        self.cs_índice_correção.set("INCC")

        for w in self.cs_result_inner.winfo_children():
            w.destroy()
        self.cs_result_placeholder = ctk.CTkLabel(
            self.cs_result_inner,
            text="Preencha os campos acima e clique em 'Calcular Simulação'",
            font=("Segoe UI", 12), text_color=TEXT_TERTIARY
        )
        self.cs_result_placeholder.pack(pady=20)
        self.cs_btn_pdf.configure(state="disabled")
        self.cs_btn_relatório.configure(state="disabled")
        self.cs_btn_email.configure(state="disabled")
        if hasattr(self, '_consórcio_dados'):
            self._consórcio_dados = None

    def _on_abrir_pasta_consorcio(self):
        os.makedirs(CONSÓRCIO_OUTPUT_DIR, exist_ok=True)
        os.startfile(CONSÓRCIO_OUTPUT_DIR)

    # =================================================================
    #  DATA LOADING
    # =================================================================
    def _load_data_async(self):
        self.after(500, lambda: threading.Thread(target=self._do_load_data, daemon=True).start())

    def _do_load_data(self):
        try:
            assessores = load_base_emails()
            summary = load_eventos_summary()

            def update_ui():
                self.kpi_assessores.configure(text=str(len(assessores)))
                self.kpi_eventos.configure(text=f"{summary['total_eventos']:,}".replace(",", "."))
                self.kpi_ativos.configure(text=str(summary["ativos"]))
                self.kpi_datas.configure(text=str(summary["datas_unicas"]))
                self.total_value_label.configure(text=fmt_currency(summary["valor_total"]))

                if summary["data_min"] and summary["data_max"]:
                    d1 = summary["data_min"].strftime("%d/%m/%Y")
                    d2 = summary["data_max"].strftime("%d/%m/%Y")
                    self.período_text.configure(
                        text=f"  Periodo:  {d1}   \u2192   {d2}   |   {summary['datas_unicas']} datas com eventos"
                    )
                    self.total_sub_label.configure(
                        text=f"Somatorio de {summary['total_eventos']:,} eventos de {len(assessores)} assessores".replace(",", ".")
                    )

                self._update_pdf_status()

                # Tipo breakdown cards
                tipo_colors = {
                    "PAGAMENTO DE JUROS": ACCENT_BLUE,
                    "AMORTIZACAO": ACCENT_GREEN,
                    "INCORPORACAO": ACCENT_ORANGE,
                    "PREMIO": ACCENT_PURPLE,
                }
                tipo_names = {
                    "PAGAMENTO DE JUROS": "Pagamento de Juros",
                    "AMORTIZACAO": "Amortizacao",
                    "INCORPORACAO": "Incorporacao",
                    "PREMIO": "Premio",
                }
                tipos = summary.get("tipos", {})
                idx = 0
                for tipo_key in ["PAGAMENTO DE JUROS", "AMORTIZACAO", "INCORPORACAO", "PREMIO"]:
                    if tipo_key in tipos:
                        r = idx // 2
                        c = idx % 2
                        self._make_tipo_card(
                            self.tipo_frame,
                            tipo_names.get(tipo_key, tipo_key),
                            tipos[tipo_key]["count"],
                            tipos[tipo_key]["total"],
                            tipo_colors.get(tipo_key, TEXT_SECONDARY),
                            r, c
                        )
                        idx += 1

            self.after(0, update_ui)
        except Exception as e:
            import traceback; traceback.print_exc()
            err_msg = f"  Erro ao carregar: {e}"
            self.after(0, lambda m=err_msg: self.período_text.configure(text=m))

    def _update_pdf_status(self):
        if os.path.exists(OUTPUT_DIR):
            pdfs = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf")]
            if pdfs:
                self.período_pdfs.configure(text=f"\u2713 {len(pdfs)} PDFs gerados  ")
                return
        self.período_pdfs.configure(text="Nenhum PDF gerado  ")

    # =================================================================
    #  LOGGING (Operations page)
    # =================================================================
    def _log(self, msg, tag=None):
        def _do():
            self.ops_log.configure(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            prefix = {"ok": "\u2713", "error": "\u2717", "skip": "\u2013"}.get(tag, " ")
            self.ops_log.insert("end", f"  [{ts}]  {prefix}  {msg}\n")
            self.ops_log.see("end")
            self.ops_log.configure(state="disabled")
        self.after(0, _do)

    def _clear_ops_log(self):
        self.ops_log.configure(state="normal")
        self.ops_log.delete("1.0", "end")
        self.ops_log.configure(state="disabled")

    def _set_ops_status(self, text, color=TEXT_PRIMARY):
        self.after(0, lambda: self.ops_status_text.configure(text=f"  {text}"))
        self.after(0, lambda: self.ops_status_dot.configure(text_color=color))

    def _set_ops_progress(self, value, text=""):
        self.after(0, lambda: self.ops_progress.set(value))
        self.after(0, lambda: self.ops_progress_text.configure(text=text))

    def _set_ops_counter(self, text):
        self.after(0, lambda: self.ops_counter.configure(text=text))

    def _set_sidebar_buttons_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for btn in self.sidebar_buttons.values():
            self.after(0, lambda b=btn, s=state: b.configure(state=s))

    # =================================================================
    #  ACTIONS
    # =================================================================
    def _go_gerar_pdfs(self):
        self._show_page("operations")
        self._action_gerar_pdfs()

    def _go_gerar_emails(self):
        self._show_page("operations")
        self._action_gerar_emails()

    def _action_gerar_pdfs(self):
        self._clear_ops_log()
        self.ops_title.configure(text="Gerando PDFs")
        self._update_sidebar_active("fluxo_rf")
        self._set_sidebar_buttons_enabled(False)
        self._set_ops_status("Iniciando...", ACCENT_ORANGE)
        self._set_ops_progress(0, "Preparando...")
        self._set_ops_counter("")
        threading.Thread(target=self._run_gerar_pdfs, daemon=True).start()

    def _run_gerar_pdfs(self):
        try:
            self._log("Importando gerador de PDFs...")
            import gerar_pdfs
            import importlib
            importlib.reload(gerar_pdfs)

            self._log("Carregando base de assessores...")
            assessores = gerar_pdfs.load_base_emails(BASE_FILE)
            self._log(f"{len(assessores)} assessores na base.")

            self._log("Carregando eventos...")
            eventos = gerar_pdfs.load_eventos(ENTRADA_FILE)
            total = sum(len(v) for v in eventos.values())
            self._log(f"{total} eventos carregados.")

            os.makedirs(OUTPUT_DIR, exist_ok=True)
            self._log("Gerando PDFs...\n")
            self._set_ops_status("Gerando PDFs...", ACCENT_ORANGE)

            ok = 0
            erros = 0
            codes = sorted(assessores.keys())
            n_total = len(codes)

            for i, código in enumerate(codes):
                info = assessores[código]
                evts = eventos.get(código, [])
                progress = (i + 1) / n_total

                try:
                    gerar_pdfs.generate_pdf(código, info, evts, OUTPUT_DIR)
                    self._log(f"{código}  {info['nome']}  ({len(evts)} eventos)", "ok")
                    ok += 1
                except Exception as e:
                    self._log(f"{código}  {info['nome']}:  {e}", "error")
                    erros += 1

                self._set_ops_progress(progress, f"{i + 1} de {n_total}")
                self._set_ops_counter(f"{ok} / {n_total}")

            self._log(f"\nFinalizado!  {ok} PDFs gerados,  {erros} erros.")
            self._set_ops_progress(1.0, "Concluído")
            self._set_ops_status(f"Concluído  -  {ok} PDFs gerados", "#00a86b" if erros == 0 else ACCENT_ORANGE)
            self._set_ops_counter(f"{ok} PDFs")
            self.after(0, self._update_pdf_status)

            if erros == 0:
                messagebox.showinfo("Sucesso", f"{ok} PDFs gerados com sucesso!\n\nPasta: {OUTPUT_DIR}")
            else:
                messagebox.showwarning("Atencao", f"{ok} PDFs gerados, {erros} com erro.")

        except Exception as e:
            self._log(f"ERRO: {e}", "error")
            self._set_ops_status(f"Erro:  {e}", ACCENT_RED)
            messagebox.showerror("Erro", str(e))
        finally:
            self._set_sidebar_buttons_enabled(True)

    def _action_gerar_emails(self):
        if not os.path.exists(OUTPUT_DIR) or not any(f.endswith(".pdf") for f in os.listdir(OUTPUT_DIR)):
            messagebox.showwarning("PDFs não encontrados", "Gere os PDFs primeiro antes de criar os emails.")
            self._show_page("fluxo_rf")
            return

        resp = messagebox.askyesno(
            "Gerar Emails no Outlook",
            "Isso vai criar um rascunho de email no Outlook para cada assessor.\n\n"
            "Os emails NÃO serão enviados, apenas criados como rascunho.\n\n"
            "O Outlook precisa estar aberto. Deseja continuar?"
        )
        if not resp:
            self._show_page("fluxo_rf")
            return

        self._clear_ops_log()
        self.ops_title.configure(text="Gerando Emails")
        self._update_sidebar_active("fluxo_rf")
        self._set_sidebar_buttons_enabled(False)
        self._set_ops_status("Conectando ao Outlook...", ACCENT_ORANGE)
        self._set_ops_progress(0, "Preparando...")
        self._set_ops_counter("")
        threading.Thread(target=self._run_gerar_emails, daemon=True).start()

    def _run_gerar_emails(self):
        try:
            import win32com.client as win32
            import pythoncom
            pythoncom.CoInitialize()

            self._log("Conectando ao Outlook...")
            try:
                outlook = win32.Dispatch("Outlook.Application")
            except Exception as e:
                self._log(f"Outlook não encontrado: {e}", "error")
                self._set_ops_status("Outlook não encontrado", ACCENT_RED)
                messagebox.showerror("Erro", "Não foi possível conectar ao Outlook.\nVerifique se está aberto.")
                return

            self._log("Carregando base de assessores...")
            assessores = load_base_emails()
            self._log(f"{len(assessores)} assessores na base.\n")

            criados = 0
            erros = 0
            sem_pdf = 0
            codes = sorted(assessores.keys())
            n_total = len(codes)
            self._set_ops_status("Criando rascunhos...", ACCENT_BLUE)

            for i, código in enumerate(codes):
                info = assessores[código]
                nome = info["nome"]
                email_assessor = info["email"]
                email_assistente = info["email_assistente"]
                progress = (i + 1) / n_total

                pdf_path = find_pdf_for_assessor(código, nome)
                if not pdf_path:
                    self._log(f"{código}  {nome}:  PDF não encontrado", "skip")
                    sem_pdf += 1
                    self._set_ops_progress(progress, f"{i + 1} de {n_total}")
                    continue

                if not email_assessor or email_assessor == "-":
                    self._log(f"{código}  {nome}:  Sem email", "skip")
                    self._set_ops_progress(progress, f"{i + 1} de {n_total}")
                    continue

                try:
                    mail = outlook.CreateItem(0)
                    mail.To = email_assessor
                    if email_assistente and email_assistente != "-":
                        mail.CC = email_assistente
                    mail.Subject = f"Fluxo de Renda Fixa - {nome} ({código})"
                    mail.HTMLBody = build_email_body(nome)
                    _attach_logo_cid(mail)
                    mail.Attachments.Add(os.path.abspath(pdf_path))
                    mail.Save()

                    cc = f"  CC: {email_assistente}" if email_assistente and email_assistente != "-" else ""
                    self._log(f"{código}  {nome}  \u2192  {email_assessor}{cc}", "ok")
                    criados += 1
                except Exception as e:
                    self._log(f"{código}  {nome}:  {e}", "error")
                    erros += 1

                self._set_ops_progress(progress, f"{i + 1} de {n_total}")
                self._set_ops_counter(f"{criados} emails")

            pythoncom.CoUninitialize()

            self._log(f"\nFinalizado!  {criados} emails criados,  {erros} erros,  {sem_pdf} sem PDF.")
            self._log("Os emails estao na pasta RASCUNHOS do Outlook.")
            self._set_ops_progress(1.0, "Concluído")
            self._set_ops_status(f"Concluído  -  {criados} emails nos Rascunhos", "#00a86b" if erros == 0 else ACCENT_ORANGE)

            if erros == 0:
                messagebox.showinfo("Emails Criados", f"{criados} emails criados nos Rascunhos do Outlook!")
            else:
                messagebox.showwarning("Atencao", f"{criados} emails criados, {erros} com erro.")

        except Exception as e:
            self._log(f"ERRO: {e}", "error")
            self._set_ops_status(f"Erro:  {e}", ACCENT_RED)
            messagebox.showerror("Erro", str(e))
        finally:
            self._set_sidebar_buttons_enabled(True)

    def _on_abrir_pasta(self):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.startfile(OUTPUT_DIR)

    # -----------------------------------------------------------------
    #  REPORTAR ERRO
    # -----------------------------------------------------------------
    def _on_reportar_erro(self):
        win = ctk.CTkToplevel(self)
        win.title("Reportar Erro - SomusApp")
        win.geometry("520x480")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        # Centralizar
        win.update_idletasks()
        x = (win.winfo_screenwidth() - 520) // 2
        y = (win.winfo_screenheight() - 480) // 2
        win.geometry(f"520x480+{x}+{y}")

        try:
            ico_path = os.path.join(BASE_DIR, "assets", "icon_somus.ico")
            if os.path.exists(ico_path):
                win.iconbitmap(ico_path)
        except Exception:
            pass

        # Header
        header = ctk.CTkFrame(win, fg_color=ACCENT_RED, height=52, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text="\u26a0  Reportar Erro",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_WHITE
        ).pack(side="left", padx=20)

        # Body
        body = ctk.CTkFrame(win, fg_color=BG_PRIMARY, corner_radius=0)
        body.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(body, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=24, pady=20)

        # Página onde ocorreu
        ctk.CTkLabel(
            inner, text="Página onde ocorreu o erro",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        erro_página = ctk.CTkComboBox(
            inner, font=("Segoe UI", 12), height=38,
            corner_radius=8, border_color=BORDER_CARD,
            values=["Dashboard", "FLUXO - RF", "Informativo", "Envio de Ordens",
                    "Ctrl Receita", "Organizador", "Envio Saldos",
                    "Corporate - Dashboard", "Corporate - Simulador", "Outro"],
            state="readonly"
        )
        erro_página.pack(fill="x", pady=(0, 14))
        erro_página.set("Outro")

        # Descrição do erro
        ctk.CTkLabel(
            inner, text="Descreva o erro encontrado",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, anchor="w"
        ).pack(fill="x", pady=(0, 4))

        erro_texto = ctk.CTkTextbox(
            inner, font=("Segoe UI", 11), height=180,
            corner_radius=8, border_width=1, border_color=BORDER_CARD,
            fg_color=BG_CARD, text_color=TEXT_PRIMARY, wrap="word"
        )
        erro_texto.pack(fill="x", pady=(0, 14))

        # Status
        status_label = ctk.CTkLabel(
            inner, text="",
            font=("Segoe UI", 10), text_color=TEXT_TERTIARY, anchor="w"
        )
        status_label.pack(fill="x", pady=(0, 8))

        # Botoes
        btn_frame = ctk.CTkFrame(inner, fg_color="transparent")
        btn_frame.pack(fill="x")

        def _enviar():
            página = erro_página.get()
            descrição = erro_texto.get("1.0", "end").strip()

            if not descrição:
                messagebox.showwarning("Campo obrigatorio", "Descreva o erro antes de enviar.", parent=win)
                return

            enviar_btn.configure(state="disabled")
            status_label.configure(text="Enviando...", text_color=ACCENT_ORANGE)
            win.update()

            try:
                import win32com.client as win32
                import pythoncom
                pythoncom.CoInitialize()

                outlook = win32.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = "artur.brito@somuscapital.com.br"
                mail.Subject = f"[SomusApp - Bug Report] {página}"

                hoje = datetime.now().strftime("%d/%m/%Y %H:%M")
                mail.HTMLBody = f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:700px;margin:0 auto;">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr><td style="background:#dc3545;padding:14px 20px;">
    <span style="font-size:15pt;font-weight:bold;color:#fff;">\u26a0 Bug Report - SomusApp</span>
  </td></tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;">
  <tr><td style="padding:0 4px;">
    <table cellpadding="0" cellspacing="0" border="0"
      style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10pt;width:100%;margin-bottom:16px;">
      <tr style="background:#f7faf9;">
        <td style="padding:8px 14px;font-weight:bold;color:#004d33;border-bottom:1.5px solid #004d33;width:140px;">Página</td>
        <td style="padding:8px 14px;color:#1a1a2e;border-bottom:1.5px solid #004d33;">{página}</td>
      </tr>
      <tr>
        <td style="padding:8px 14px;font-weight:bold;color:#004d33;border-bottom:1px solid #eee;">Data/Hora</td>
        <td style="padding:8px 14px;color:#1a1a2e;border-bottom:1px solid #eee;">{hoje}</td>
      </tr>
      <tr style="background:#f7faf9;">
        <td style="padding:8px 14px;font-weight:bold;color:#004d33;border-bottom:1px solid #eee;">Modulo</td>
        <td style="padding:8px 14px;color:#1a1a2e;border-bottom:1px solid #eee;">{self.current_module}</td>
      </tr>
    </table>
  </td></tr>
  <tr><td style="padding:22px 0 8px 0;">
    <table cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background-color:#dc3545;width:4px;border-radius:4px;">&nbsp;</td>
      <td style="padding-left:12px;">
        <span style="font-size:12.5pt;color:#dc3545;font-weight:bold;">Descrição do Erro</span>
      </td>
    </tr></table>
  </td></tr>
  <tr><td style="padding:8px 4px;">
    <div style="background:#fdf2f2;border:1px solid #f5c6cb;border-radius:6px;padding:14px 18px;font-size:10.5pt;color:#1a1a2e;">
      {descrição.replace(chr(10), '<br>')}
    </div>
  </td></tr>
  <tr><td style="padding:24px 0 0 0;">
    <hr style="border:none;border-top:1px solid #e0e0e0;margin:0 0 8px 0;">
    <span style="font-size:8pt;color:#9ca3af;">Enviado automaticamente pelo SomusApp - BETA</span>
  </td></tr>
</table>
</div>
"""
                mail.Send()
                pythoncom.CoUninitialize()

                status_label.configure(text="Erro reportado com sucesso!", text_color="#00a86b")
                messagebox.showinfo("Enviado", "Bug report enviado com sucesso!", parent=win)
                win.after(500, win.destroy)

            except Exception as e:
                enviar_btn.configure(state="normal")
                status_label.configure(text=f"Falha: {e}", text_color=ACCENT_RED)
                messagebox.showerror("Erro", f"Não foi possível enviar:\n{e}", parent=win)

        enviar_btn = ctk.CTkButton(
            btn_frame, text="\u2709  Enviar",
            font=("Segoe UI", 13, "bold"),
            fg_color=ACCENT_RED, hover_color="#b02a37",
            height=42, corner_radius=10,
            command=_enviar,
        )
        enviar_btn.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame, text="Cancelar",
            font=("Segoe UI", 11),
            fg_color="#6b7280", hover_color="#555d6a",
            height=42, corner_radius=10, width=100,
            command=win.destroy,
        ).pack(side="left")

    @staticmethod
    def _darken(hex_color):
        r = max(0, int(hex_color[1:3], 16) - 30)
        g = max(0, int(hex_color[3:5], 16) - 30)
        b = max(0, int(hex_color[5:7], 16) - 30)
        return f"#{r:02x}{g:02x}{b:02x}"


# =====================================================================
#  LOGIN
# =====================================================================

_LOGIN_BG = "#004d33"
_LOGIN_BG_LIGHT = "#005c3d"
_LOGIN_CARD = "#ffffff"
_LOGIN_BTN = "#f0e0c0"
_LOGIN_BTN_HOVER = "#e6d1a8"
_LOGIN_BTN_TEXT = "#004d33"
_LOGIN_ACCENT = "#00b876"


class LoginWindow(ctk.CTkToplevel):
    """Tela de login com senha fixa - visual premium Somus Capital."""

    APP_PASSWORD = "@Produtos1"

    def __init__(self, master, on_success):
        super().__init__(master)
        self.on_success = on_success
        self.title("Somus Capital")
        self.geometry("480x600")
        self.resizable(False, False)
        self.configure(fg_color=_LOGIN_BG)
        self.transient(master)
        self.grab_set()

        # Centralizar na tela
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 480) // 2
        y = (self.winfo_screenheight() - 600) // 2
        self.geometry(f"480x600+{x}+{y}")

        # Icone
        try:
            ico_path = os.path.join(BASE_DIR, "assets", "icon_somus.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
        except Exception:
            pass

        # Fechar janela = fechar app
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()

    def _build_ui(self):
        # ── Topo verde com logo ──
        top_frame = ctk.CTkFrame(self, fg_color="transparent", height=180)
        top_frame.pack(fill="x", pady=(0, 0))
        top_frame.pack_propagate(False)

        # Logo (branca sobre fundo verde = visivel)
        try:
            logo_img = Image.open(LOGO_PATH)
            self._logo_ctk = ctk.CTkImage(
                light_image=logo_img, dark_image=logo_img, size=(200, 49)
            )
            ctk.CTkLabel(top_frame, image=self._logo_ctk, text="").place(
                relx=0.5, rely=0.38, anchor="center"
            )
        except Exception:
            ctk.CTkLabel(
                top_frame, text="SOMUS CAPITAL",
                font=("Segoe UI", 24, "bold"), text_color=TEXT_WHITE
            ).place(relx=0.5, rely=0.38, anchor="center")

        # Subtítulos
        ctk.CTkLabel(
            top_frame, text="Mesa de Produtos  |  Agente de Investimentos",
            font=("Segoe UI", 11), text_color="#8fbfaa"
        ).place(relx=0.5, rely=0.68, anchor="center")

        # Linha decorativa verde claro
        line = ctk.CTkFrame(self, fg_color=_LOGIN_ACCENT, height=3, corner_radius=2)
        line.pack(fill="x", padx=60, pady=(0, 0))

        # ── Card branco central ──
        card = ctk.CTkFrame(
            self, fg_color=_LOGIN_CARD, corner_radius=16,
            border_width=0, width=360, height=280
        )
        card.pack(pady=(30, 0))
        card.pack_propagate(False)

        # Icone Somus pequeno no card
        try:
            icon_path = os.path.join(BASE_DIR, "assets", "icon_somus.png")
            icon_img = Image.open(icon_path)
            self._icon_ctk = ctk.CTkImage(
                light_image=icon_img, dark_image=icon_img, size=(48, 48)
            )
            ctk.CTkLabel(card, image=self._icon_ctk, text="").pack(pady=(24, 4))
        except Exception:
            ctk.CTkLabel(
                card, text="S", font=("Segoe UI", 28, "bold"),
                text_color=_LOGIN_BG, width=48, height=48,
                fg_color="#e8f5e9", corner_radius=24
            ).pack(pady=(24, 4))

        # Título do card
        ctk.CTkLabel(
            card, text="Acesso Restrito",
            font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY
        ).pack(pady=(0, 16))

        # Input senha
        self.password_entry = ctk.CTkEntry(
            card, placeholder_text="Digite a senha...",
            show="\u2022", width=280, height=44,
            font=("Segoe UI", 13),
            border_color="#d1d5db", border_width=1,
            fg_color="#f9fafb", text_color=TEXT_PRIMARY,
            corner_radius=10
        )
        self.password_entry.pack(pady=(0, 6))
        self.password_entry.bind("<Return>", lambda e: self._try_login())
        self.password_entry.focus_set()

        # Mensagem de erro
        self.error_label = ctk.CTkLabel(
            card, text="", font=("Segoe UI", 10),
            text_color="#dc3545", height=18
        )
        self.error_label.pack(pady=(0, 8))

        # Botao entrar (bege/dourado)
        self.login_btn = ctk.CTkButton(
            card, text="Entrar", width=280, height=44,
            font=("Segoe UI", 14, "bold"),
            fg_color=_LOGIN_BTN, hover_color=_LOGIN_BTN_HOVER,
            text_color=_LOGIN_BTN_TEXT,
            corner_radius=10, command=self._try_login
        )
        self.login_btn.pack(pady=(0, 24))

        # ── Footer ──
        ctk.CTkLabel(
            self, text="SomusApp - BETA",
            font=("Segoe UI", 9), text_color="#5a9a7d"
        ).pack(side="bottom", pady=(0, 16))

    def _try_login(self):
        senha = self.password_entry.get()
        if senha == self.APP_PASSWORD:
            self.grab_release()
            self.destroy()
            self.on_success()
        else:
            self.error_label.configure(text="Senha incorreta. Tente novamente.")
            self.password_entry.delete(0, "end")
            self.password_entry.focus_set()
            # Shake animation
            orig_x = self.winfo_x()
            orig_y = self.winfo_y()
            import time
            for dx in [10, -10, 8, -8, 4, -4, 0]:
                self.geometry(f"+{orig_x + dx}+{orig_y}")
                self.update_idletasks()
                time.sleep(0.03)

    def _on_close(self):
        self.grab_release()
        self.destroy()
        self.master.quit()


def main():
    # Definir AppUserModelID para o Windows usar o icone do app na barra de tarefas
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("somus.capital.fluxorf")
    except Exception:
        pass

    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("dark-blue")

    # Janela root oculta (host para o login)
    root = ctk.CTk()
    root.withdraw()

    try:
        ico_path = os.path.join(BASE_DIR, "assets", "icon_somus.ico")
        if os.path.exists(ico_path):
            root.iconbitmap(ico_path)
    except Exception:
        pass

    login_ok = [False]

    def on_login_success():
        login_ok[0] = True
        root.quit()

    login = LoginWindow(root, on_login_success)
    root.mainloop()

    # Apos login bem-sucedido, destruir root e criar o app em novo mainloop
    if login_ok[0]:
        root.destroy()
        app = App()
        app.mainloop()


if __name__ == "__main__":
    main()
