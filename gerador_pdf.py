"""
Gerador de PDF — v4.0
Compatível com Windows e Linux.
Cabeçalho com brasão, checklists por tipo, ✓/✗ coloridos.
"""
import io, os, re, sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, Image)
from reportlab.platypus.flowables import HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# LOCALIZAÇÃO DE FONTES (Windows / Linux / Mac)
# ─────────────────────────────────────────────────────────────────────────────
def _find_font(candidates: list) -> str | None:
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None

_WIN_FONTS  = r"C:\Windows\Fonts"
_LIN_LIB    = "/usr/share/fonts/truetype/liberation"
_LIN_DEJAVU = "/usr/share/fonts/truetype/dejavu"
_LIN_FIRA   = "/usr/share/fonts/truetype/fira"

# Fonte principal (texto)
_FONT_REG = _find_font([
    os.path.join(_WIN_FONTS, "arial.ttf"),
    os.path.join(_WIN_FONTS, "Arial.ttf"),
    os.path.join(_LIN_LIB,  "LiberationSans-Regular.ttf"),
    os.path.join(_LIN_DEJAVU,"DejaVuSans.ttf"),
])
_FONT_BOL = _find_font([
    os.path.join(_WIN_FONTS, "arialbd.ttf"),
    os.path.join(_WIN_FONTS, "ArialBD.ttf"),
    os.path.join(_LIN_LIB,  "LiberationSans-Bold.ttf"),
    os.path.join(_LIN_DEJAVU,"DejaVuSans-Bold.ttf"),
])

# Fonte para símbolos unicode ✓ ✗ (DejaVu ou Segoe)
_FONT_SYM = _find_font([
    os.path.join(_WIN_FONTS, "seguisym.ttf"),   # Segoe UI Symbol (Win)
    os.path.join(_WIN_FONTS, "arial.ttf"),       # Arial suporta ✓ no Win
    os.path.join(_LIN_DEJAVU,"DejaVuSans.ttf"),
])

try:
    if _FONT_REG:
        pdfmetrics.registerFont(TTFont("_MainR", _FONT_REG))
        _F = "_MainR"
    else:
        _F = "Helvetica"

    if _FONT_BOL:
        pdfmetrics.registerFont(TTFont("_MainB", _FONT_BOL))
        _FB = "_MainB"
    else:
        _FB = "Helvetica-Bold"

    if _FONT_SYM:
        pdfmetrics.registerFont(TTFont("_Sym", _FONT_SYM))
        _FSYM = "_Sym"
    else:
        _FSYM = _FB
except Exception:
    _F, _FB, _FSYM = "Helvetica", "Helvetica-Bold", "Helvetica-Bold"

# Símbolos de checklist
CHK_OK  = "\u2713"  # ✓
CHK_ERR = "\u2717"  # ✗
COR_OK  = colors.HexColor("#15803D")
COR_ERR = colors.HexColor("#DC2626")

# Pasta base (onde está este arquivo)
_BASE       = os.path.dirname(os.path.abspath(__file__))
_BRASAO_PNG = os.path.join(_BASE, "brasao.png")

# Dimensões de página
PW, PH = A4
ML = MR = 1.9 * cm;  MT = 2.54 * cm;  MB = 2.0 * cm
TW = PW - ML - MR
ML_C = MR_C = 2.25 * cm
TW_C = PW - ML_C - MR_C

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _S(name, size=12, bold=False, align=TA_LEFT, indent=0, sb=0, sa=0, color=None):
    kw = dict(fontName=_FB if bold else _F, fontSize=size,
              leading=size * 1.35, alignment=align,
              leftIndent=indent, spaceBefore=sb, spaceAfter=sa, wordWrap="LTR")
    if color:
        kw["textColor"] = color
    return ParagraphStyle(name, **kw)

def _P(text, style):
    return Paragraph(str(text) if text else "", style)

def _LN(pts=6):
    return Spacer(1, pts)

def _tbs():
    return [
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME",      (0, 0), (-1, -1), _F),
        ("FONTSIZE",      (0, 0), (-1, -1), 12),
        ("LEADING",       (0, 0), (-1, -1), 14),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]

# ─────────────────────────────────────────────────────────────────────────────
# CABEÇALHO COM BRASÃO
# ─────────────────────────────────────────────────────────────────────────────
def _build_header(tw):
    s = _S("h_c", size=11, bold=True, align=TA_CENTER)
    brasao = (Image(_BRASAO_PNG, width=2.646 * cm, height=1.535 * cm)
              if os.path.exists(_BRASAO_PNG) else Spacer(1, 1.5 * cm))
    t = Table([
        [brasao],
        [_P("<b>ESTADO DO MARANHÃO</b>", s)],
        [_P("<b>PREFEITURA MUNICIPAL DE GOVERNADOR EDISON LOBÃO</b>", s)],
        [_P("<b>CONTROLADORIA DO MUNICÍPIO</b>", s)],
    ], colWidths=[tw])
    t.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))
    return t

def _sep(tw):
    return HRFlowable(width=tw, thickness=1, color=colors.black,
                      spaceAfter=4, spaceBefore=2)

# ─────────────────────────────────────────────────────────────────────────────
# UTILITÁRIO: PRÓXIMO NÚMERO
# ─────────────────────────────────────────────────────────────────────────────
def proximo_numero(excel_path: str) -> int:
    try:
        df = pd.read_excel(excel_path, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        col = "NÚMERO DO DOCUMENTO"
        if col not in df.columns:
            return 1
        nums = [int(m.group(1))
                for v in df[col].dropna()
                if (m := re.match(r"^(\d+)$", str(v).strip()))]
        return max(nums) + 1 if nums else 1
    except Exception:
        return 1

# ─────────────────────────────────────────────────────────────────────────────
# CAPA DO PROCESSO
# ─────────────────────────────────────────────────────────────────────────────
def gerar_pdf_capa(d: dict) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=ML_C, rightMargin=MR_C, topMargin=MT, bottomMargin=MB)
    C0, C1 = 4.75 * cm, TW_C - 4.75 * cm
    sH = _S("cH", size=16, bold=True, align=TA_CENTER)
    sL = _S("cL", size=16, bold=True)
    sV = _S("cV", size=16)
    sO = _S("cO", size=12)

    def td(t, b=False): return _P(f"<b>{t}</b>" if b else str(t), sL if b else sV)

    rows = [
        [_P("<b>PROCESSO DE PAGAMENTO</b>", sH), ""],
        [td("Órgão:", True),           td(d.get("orgao", ""))],
        [td("Processo:", True),         td(d.get("processo", ""))],
        [td("Fornecedor:", True),       td(d.get("fornecedor", ""))],
        [td("CNPJ:", True),             td(d.get("cnpj", ""))],
        [td("NF/Fatura:", True),        td(d.get("nf", ""))],
        [td("Contrato:", True),         td(d.get("contrato", ""))],
        [td("Modalidade:", True),       td(d.get("modalidade", ""))],
        [td("Período de ref.:", True),  td(d.get("periodo_ref", ""))],
        [td("N° Ordem de C.:", True),   td(d.get("ordem_compra", ""))],
        [td("Data da NF.:", True),      td(d.get("data_nf", ""))],
        [td("Secretário(a):", True),    td(d.get("secretario", ""))],
        [td("Data do ateste:", True),   td(d.get("data_ateste", ""))],
    ]
    st2 = _tbs() + [
        ("SPAN",          (0, 0), (1, 0)),
        ("ALIGN",         (0, 0), (1, 0), "CENTER"),
        ("FONTNAME",      (0, 0), (1, 0), _FB),
        ("FONTSIZE",      (0, 0), (-1, -1), 16),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    t0 = Table(rows, colWidths=[C0, C1])
    t0.setStyle(TableStyle(st2))

    obs = d.get("obs_capa", "")
    t1 = Table([[_P(f"<b>Obs.:</b>  {obs}", sO)]], colWidths=[TW_C])
    t1.setStyle(TableStyle(_tbs() + [
        ("FONTSIZE",     (0, 0), (-1, -1), 12),
        ("MINROWHEIGHT", (0, 0), (-1, -1), 2.5 * cm),
    ]))
    doc.build([_build_header(TW_C), _sep(TW_C), _LN(6), t0, _LN(4), t1])
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# CHECKLIST BUILDER  (compartilhado pelos tipos padrão/eng/tdf)
# ─────────────────────────────────────────────────────────────────────────────
def _build_checklist_padrao(checklist: list, situacoes: list, TW: float):
    """Retorna Table de checklist com coluna Situação (✓/✗)."""
    CK_I = 1.3 * cm
    CK_S = 2.0 * cm
    CK_D = TW - CK_I - CK_S
    s12c = _S("_c12c", size=12, align=TA_CENTER)
    s12  = _S("_c12",  size=12)
    s_h  = _S("_cHdr", size=10, bold=True, align=TA_CENTER)
    rows = [[_P("<b>Item</b>", s12c),
             _P("<b>Descrição: Documentos – Ato</b>", s12c),
             _P("<b>Situação</b>", s_h)]]
    for i, item in enumerate(checklist):
        ok  = situacoes[i]
        sym = CHK_OK if ok else CHK_ERR
        cor = COR_OK  if ok else COR_ERR
        s_sym = ParagraphStyle(f"_sym{i}", fontName=_FSYM, fontSize=16,
                               alignment=TA_CENTER, textColor=cor, leading=20)
        rows.append([_P(str(i + 1), s12c), _P(item, s12), Paragraph(sym, s_sym)])
    st = _tbs() + [
        ("FONTNAME",      (0, 0), (-1, 0), _FB),
        ("ALIGN",         (0, 0), (0, -1), "CENTER"),
        ("ALIGN",         (2, 0), (2, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, 0), 5),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING",    (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
    ]
    t = Table(rows, colWidths=[CK_I, CK_D, CK_S])
    t.setStyle(TableStyle(st))
    return t

def _build_checklist_passagem(checklist: list, situacoes: list, TW: float):
    """Retorna Table de checklist com colunas Sim / Não."""
    CK_I = 1.7 * cm; CK_S = 1.8 * cm; CK_N = 1.8 * cm
    CK_D = TW - CK_I - CK_S - CK_N
    s12c = _S("_pc12c", size=12, align=TA_CENTER)
    s12  = _S("_pc12",  size=12)
    rows = [[_P("<b>Item</b>", s12c),
             _P("<b>Descrição: Documentos - Ato</b>", s12c),
             _P("<b>Sim</b>", s12c), _P("<b>Não</b>", s12c)]]
    for i, item in enumerate(checklist):
        ok = situacoes[i]
        s_ok  = ParagraphStyle(f"_psim{i}", fontName=_FSYM, fontSize=16,
                               alignment=TA_CENTER, textColor=COR_OK,  leading=20)
        s_nok = ParagraphStyle(f"_pnao{i}", fontName=_FSYM, fontSize=16,
                               alignment=TA_CENTER, textColor=COR_ERR, leading=20)
        rows.append([
            _P(str(i + 1), s12c),
            _P(item, s12),
            Paragraph(CHK_OK,  s_ok)  if ok  else Paragraph("", s12c),
            Paragraph(CHK_ERR, s_nok) if not ok else Paragraph("", s12c),
        ])
    st = _tbs() + [
        ("FONTNAME",      (0, 0), (-1, 0), _FB),
        ("ALIGN",         (0, 0), (0, -1), "CENTER"),
        ("ALIGN",         (2, 0), (3, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
    ]
    t = Table(rows, colWidths=[CK_I, CK_D, CK_S, CK_N])
    t.setStyle(TableStyle(st))
    return t

# ─────────────────────────────────────────────────────────────────────────────
# PARECER PADRÃO / ENGENHARIA / TDF  (estrutura unificada)
# ─────────────────────────────────────────────────────────────────────────────
def gerar_pdf_parecer_padrao(d: dict, tipo: str, deferir: bool,
                              checklist: list, situacoes: list | None = None) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB)
    n = len(checklist)
    if situacoes is None or len(situacoes) != n:
        situacoes = [True] * n

    CW6 = [2.49 * cm, 2.25 * cm, 2.69 * cm, 3.26 * cm, 3.05 * cm, 3.44 * cm]
    s12   = _S("n12",  size=12)
    s12j  = _S("n12j", size=12, align=TA_JUSTIFY)
    s12bj = _S("n12bj",size=12, bold=True, align=TA_JUSTIFY)
    s12r  = _S("n12r", size=12, align=TA_RIGHT)
    s12c  = _S("n12c", size=12, align=TA_CENTER)
    s12i  = _S("n12i", size=12, indent=8.6 * cm)
    dec = "DEFERIMOS O PAGAMENTO:" if deferir else "INDEFERIMOS O PAGAMENTO:"

    def td(t, b=False): return _P(f"<b>{t}</b>" if b else str(t), s12)

    t0_data = [
        [td("OBJETO:", True),             td(d.get("objeto", "")),  "", "", "", ""],
        [td("Secretaria/Programa:", True), "", td(d.get("orgao", "")),    "", "", ""],
        [td("Fornecedor/Credor:", True),   "", td(d.get("fornecedor", "")), "", "", ""],
        [td("Modalidade", True),           "", td(d.get("modalidade", "")), "", "", ""],
        [td("Contrato", True),             "", td(d.get("contrato", "")),   "", "", ""],
        [td("CNPJ/CPF Nº", True),          "", td(d.get("cnpj", "")),       "", "", ""],
        [td("Documento Fiscal", True), "",
         td(d.get("tipo_doc", "")),
         td("Nº " + d.get("nf", "")),
         td("Tipo " + d.get("tipo_nf", "")),
         td(d.get("valor", ""))],
    ]
    t0_st = _tbs() + [
        ("SPAN", (1, 0), (5, 0)),
        ("SPAN", (0, 1), (1, 1)), ("SPAN", (2, 1), (5, 1)),
        ("SPAN", (0, 2), (1, 2)), ("SPAN", (2, 2), (5, 2)),
        ("SPAN", (0, 3), (1, 3)), ("SPAN", (2, 3), (5, 3)),
        ("SPAN", (0, 4), (1, 4)), ("SPAN", (2, 4), (5, 4)),
        ("SPAN", (0, 5), (1, 5)), ("SPAN", (2, 5), (5, 5)),
        ("SPAN", (0, 6), (1, 6)),
        ("FONTNAME", (0, 0), (0, -1), _FB),
    ]
    t0 = Table(t0_data, colWidths=CW6)
    t0.setStyle(TableStyle(t0_st))
    t1 = _build_checklist_padrao(checklist, situacoes, TW)

    obs = d.get("obs", "").strip()
    obs_b = [_P(l.strip(), s12j) for l in obs.split("\n") if l.strip()] if obs else [_LN(12)]

    story = [
        _build_header(TW), _sep(TW), _LN(4),
        _P(f"<b>PARECER DE VERIFICAÇÃO E ANÁLISE DOCUMENTAL Nº "
           f"{d.get('processo', '')} (LIBERAÇÃO PARA PAGAMENTO)</b>", s12bj),
        _LN(4), _P("Ao ", s12),
        _P(f"Órgão / Departamento: {d.get('orgao', '')}", s12),
        _LN(4), _P("Ref. Processo de Pagamento de Despesa.", s12),
        t0, _LN(4),
        _P("Após análise e verificação da documentação constante no processo "
           "de pagamento acima citado, constatamos o seguinte:", s12),
        _LN(4), t1, _LN(4),
        _P("OBSERVAÇÃO:", s12), _LN(2),
    ] + obs_b + [
        _LN(4),
        _P(f"Governador Edison Lobão/MA, {d.get('data_ateste', '')}", s12r),
        _P("Nestes Termos:", s12i),
        _P(dec, s12i),
        _LN(20), _LN(20), _LN(20),
        _P("Thiago Soares Lima", s12c),
        _P("Controlador Geral",  s12c),
        _P("Portaria 002/2025",  s12c),
    ]
    doc.build(story)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# PARECER DE RESTITUIÇÃO DE PASSAGEM
# ─────────────────────────────────────────────────────────────────────────────
def gerar_pdf_parecer_passagem(d: dict, deferir: bool,
                                checklist: list, situacoes: list | None = None) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB)
    n = len(checklist)
    if situacoes is None or len(situacoes) != n:
        situacoes = [True] * n

    CW3 = [2.5 * cm, 2.3 * cm, 12.4 * cm]
    s12   = _S("pp12",  size=12)
    s12j  = _S("pp12j", size=12, align=TA_JUSTIFY)
    s12bj = _S("pp12bj",size=12, bold=True, align=TA_JUSTIFY)
    s12r  = _S("pp12r", size=12, align=TA_RIGHT)
    s12c  = _S("pp12c", size=12, align=TA_CENTER)
    s12i  = _S("pp12i", size=12, indent=8.6 * cm)
    dec = "DEFERIMOS O PAGAMENTO:" if deferir else "INDEFERIMOS O PAGAMENTO:"

    def td(t, b=False): return _P(f"<b>{t}</b>" if b else str(t), s12)

    t0_data = [
        [td("OBJETO:", True),             td(d.get("objeto", "")), ""],
        [td("Secretaria/Programa:", True), "", td(d.get("orgao", ""))],
        [td("Documento", True),            "", td(d.get("documento", ""))],
        [td("Solicitante:", True),         "", td(d.get("solicitante", ""))],
        [td("CPF Nº:", True),              "", td(d.get("cpf", ""))],
        [td("Valor:", True),               "", td(d.get("valor", ""))],
    ]
    t0_st = _tbs() + [
        ("SPAN", (1, 0), (2, 0)),
        ("SPAN", (0, 1), (1, 1)), ("SPAN", (0, 2), (1, 2)),
        ("SPAN", (0, 3), (1, 3)), ("SPAN", (0, 4), (1, 4)),
        ("SPAN", (0, 5), (1, 5)),
        ("FONTNAME", (0, 0), (0, -1), _FB),
    ]
    t0 = Table(t0_data, colWidths=CW3)
    t0.setStyle(TableStyle(t0_st))
    t1 = _build_checklist_passagem(checklist, situacoes, TW)

    obs = d.get("obs", "").strip()
    obs_b = [_P(l.strip(), s12j) for l in obs.split("\n") if l.strip()] if obs else [_LN(12)]

    story = [
        _build_header(TW), _sep(TW), _LN(4),
        _P(f"<b>PARECER DE VERIFICAÇÃO E ANÁLISE DOCUMENTAL Nº "
           f"{d.get('processo', '')} LIBERAÇÃO PARA PAGAMENTO</b>", s12bj),
        _LN(4), _P("Ao ", s12),
        _P(f"Órgão / Departamento: {d.get('orgao', '')}", s12),
        _LN(4), _P("Ref. Processo de Pagamento de Despesa.", s12),
        t0, _LN(4),
        _P("Após análise e verificação da documentação constante no processo "
           "de pagamento acima citado, constatamos o seguinte:", s12),
        _LN(4), t1, _LN(4),
        _P("OBSERVAÇÃO:", s12), _LN(2),
    ] + obs_b + [
        _LN(4), _P("Nestes Termos:", s12), _P(dec, s12), _LN(4),
        _P(f"Governador Edison Lobão/MA, {d.get('data_ateste', '')}.", s12r),
        _P("Nestes Termos:", s12i),
        _LN(20), _LN(20),
        _P("Thiago Soares Lima", s12c),
        _P("Controladora Geral", s12c),
        _P("Portaria 002/2025",  s12c),
    ]
    doc.build(story)
    return buf.getvalue()
