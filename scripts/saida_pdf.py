"""Relatório executivo em PDF (A4), pensado pra ser lido em 2 minutos."""

from reportlab import rl_config
from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from analise import brl, pct

# Sem data de criação nem ID aleatório no arquivo: mesmo dado, mesmo PDF.
rl_config.invariant = 1

PRIMARIA = colors.HexColor("#1F4E5A")
META = colors.HexColor("#C9D3D6")
TINTA = colors.HexColor("#1B1F23")
SUAVE = colors.HexColor("#5F6368")
FUNDO = colors.HexColor("#F2F5F6")
VERDE = colors.HexColor("#1E7F4F")
VERMELHO = colors.HexColor("#B3261E")

TITULO = ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=TINTA)
SUB = ParagraphStyle("sub", fontName="Helvetica", fontSize=9.5, leading=13, textColor=SUAVE)
SECAO = ParagraphStyle("secao", fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=PRIMARIA,
                       spaceBefore=10, spaceAfter=5)
CORPO = ParagraphStyle("corpo", fontName="Helvetica", fontSize=9.5, leading=13, textColor=TINTA)
PEQUENO = ParagraphStyle("pequeno", fontName="Helvetica", fontSize=8, leading=10.5, textColor=SUAVE)
KPI_LBL = ParagraphStyle("kpil", fontName="Helvetica", fontSize=7.5, leading=9, textColor=SUAVE)
KPI_VAL = ParagraphStyle("kpiv", fontName="Helvetica-Bold", fontSize=14, leading=17, textColor=TINTA)

LARGURA = A4[0] - 30 * mm


def cor_sinal(v, limite=0):
    return "#1E7F4F" if v >= limite else "#B3261E"


def esc(texto):
    """Texto que vai dentro de Paragraph (que interpreta marcação tipo HTML)."""
    return str(texto).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def kpis(r):
    t = r["total"]
    celulas = [
        ("FATURAMENTO LÍQUIDO", brl(t["liquido"]), None),
        (f"VS {r['mes_anterior_extenso'].upper()}",
         pct(t["variacao"], sinal=True) if t["variacao"] is not None else "—",
         cor_sinal(t["variacao"]) if t["variacao"] is not None else None),
        ("DA META CONSOLIDADA", pct(t["pct_meta"]) if t["pct_meta"] is not None else "—",
         cor_sinal(t["pct_meta"], 1) if t["pct_meta"] is not None else None),
        ("VENDAS · TICKET MÉDIO", f"{t['vendas']} · {brl(t['ticket_medio'])}", None),
    ]
    linha = []
    for rotulo, valor, cor in celulas:
        estilo = KPI_VAL if cor is None else ParagraphStyle("k", parent=KPI_VAL, textColor=colors.HexColor(cor))
        linha.append([Paragraph(rotulo, KPI_LBL), Paragraph(valor, estilo)])
    tabela = Table([linha], colWidths=[LARGURA / 4] * 4)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO),
        ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
    ]))
    return tabela


def tabela_lojas(r):
    num = ParagraphStyle("n", parent=CORPO, alignment=TA_RIGHT)
    cab = ["Loja", "Faturamento", "Meta", "% meta", f"vs {r['mes_anterior_extenso'].split('/')[0]}",
           "Vendas", "Ticket médio"]
    linhas = [cab]
    for l in r["lojas"] + [dict(r["total"], loja="Total", enviou=True)]:
        pm = l["pct_meta"]
        va = l["variacao"]
        nome = l["loja"] if l["enviou"] else f"{l['loja']} (sem dados)"
        linhas.append([
            Paragraph(f"<b>{esc(nome)}</b>" if l["loja"] == "Total" else esc(nome), CORPO),
            Paragraph(brl(l["liquido"]), num),
            Paragraph(brl(l["meta"]) if l["meta"] else "—", num),
            Paragraph(f'<font color="{cor_sinal(pm, 1)}"><b>{pct(pm)}</b></font>' if pm is not None else "—", num),
            Paragraph(f'<font color="{cor_sinal(va)}">{pct(va, sinal=True)}</font>' if va is not None else "—", num),
            Paragraph(str(l["vendas"]), num),
            Paragraph(brl(l["ticket_medio"]), num),
        ])
    t = Table(linhas, colWidths=[LARGURA * f for f in (0.19, 0.17, 0.17, 0.11, 0.11, 0.09, 0.16)])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("TEXTCOLOR", (0, 0), (-1, 0), SUAVE), ("ALIGN", (1, 0), (-1, 0), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, PRIMARIA),
        ("LINEABOVE", (0, -1), (-1, -1), 0.8, SUAVE),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, FUNDO]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def grafico_meta(r, largura, altura=150):
    d = Drawing(largura, altura)
    g = VerticalBarChart()
    g.x, g.y, g.width, g.height = 42, 22, largura - 52, altura - 58
    g.data = [[l["liquido"] for l in r["lojas"]], [l["meta"] or 0 for l in r["lojas"]]]
    g.categoryAxis.categoryNames = [l["loja"] for l in r["lojas"]]
    g.categoryAxis.labels.fontName = "Helvetica"
    g.categoryAxis.labels.fontSize = 8
    g.valueAxis.valueMin = 0
    g.valueAxis.labels.fontName = "Helvetica"
    g.valueAxis.labels.fontSize = 7
    g.valueAxis.labelTextFormat = lambda v: f"{v / 1000:.0f} mil"
    g.bars[0].fillColor = PRIMARIA
    g.bars[1].fillColor = META
    g.bars.strokeColor = None
    g.groupSpacing = 12
    d.add(g)
    d.add(String(42, altura - 12, "Faturamento líquido x meta", fontName="Helvetica-Bold", fontSize=9, fillColor=TINTA))
    # legenda desenhada ("■" não existe na fonte padrão do PDF), abaixo do título
    x = 42
    for cor, rotulo in [(PRIMARIA, "faturamento líquido"), (META, "meta")]:
        d.add(Rect(x, altura - 27, 7, 7, fillColor=cor, strokeColor=None))
        d.add(String(x + 10, altura - 26, rotulo, fontName="Helvetica", fontSize=7.5, fillColor=SUAVE))
        x += 95
    return d


def grafico_categorias(r, largura, altura=150):
    cats = list(reversed(r["categorias"]))
    d = Drawing(largura, altura)
    g = HorizontalBarChart()
    g.x, g.y, g.width, g.height = 118, 8, largura - 160, altura - 32
    g.data = [[c["liquido"] for c in cats]]
    g.categoryAxis.categoryNames = [c["categoria"] for c in cats]
    g.categoryAxis.labels.fontName = "Helvetica"
    g.categoryAxis.labels.fontSize = 7.5
    g.categoryAxis.labels.dx = -4
    g.categoryAxis.labels.boxAnchor = "e"
    g.valueAxis.visible = False
    g.valueAxis.valueMin = 0
    g.bars[0].fillColor = PRIMARIA
    g.bars.strokeColor = None
    g.barLabelFormat = lambda v: pct(v / r["total"]["liquido"]) if r["total"]["liquido"] else ""
    g.barLabels.fontName = "Helvetica"
    g.barLabels.fontSize = 7
    g.barLabels.boxAnchor = "w"
    g.barLabels.dx = 3
    d.add(g)
    d.add(String(0, altura - 12, "Participação por categoria", fontName="Helvetica-Bold", fontSize=9, fillColor=TINTA))
    return d


def grafico_diario(r, largura, altura=120):
    d = Drawing(largura, altura)
    g = VerticalBarChart()
    g.x, g.y, g.width, g.height = 42, 22, largura - 50, altura - 42
    g.data = [[x["liquido"] for x in r["diario"]]]
    g.categoryAxis.categoryNames = [x["data"][8:10] for x in r["diario"]]
    g.categoryAxis.labels.fontName = "Helvetica"
    g.categoryAxis.labels.fontSize = 6.5
    g.valueAxis.valueMin = 0
    g.valueAxis.labels.fontName = "Helvetica"
    g.valueAxis.labels.fontSize = 7
    g.valueAxis.labelTextFormat = lambda v: f"{v / 1000:.1f} mil".replace(".", ",")
    g.bars[0].fillColor = PRIMARIA
    g.bars.strokeColor = None
    g.barSpacing = 1
    d.add(g)
    d.add(String(42, altura - 12, "Faturamento líquido por dia (as 3 lojas)", fontName="Helvetica-Bold",
                 fontSize=9, fillColor=TINTA))
    return d


def tabela_qualidade(r):
    linhas = [["Loja", "Linhas usadas", "O que foi corrigido ou descartado"]]
    for q in r["qualidade"]:
        linhas.append([esc(q["loja"]), f"{q['linhas_validas']} de {q['linhas_lidas']}",
                       Paragraph(esc("; ".join(q["problemas"]) or "nenhum problema encontrado"), PEQUENO)])
    t = Table(linhas, colWidths=[LARGURA * 0.16, LARGURA * 0.16, LARGURA * 0.68])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("TEXTCOLOR", (0, 0), (-1, -1), SUAVE), ("LINEBELOW", (0, 0), (-1, 0), 0.6, SUAVE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(SUAVE)
    canvas.drawString(15 * mm, 10 * mm, "Farmácia Vitalis Manipulação é fictícia e os dados são simulados · "
                                        "relatório gerado automaticamente a partir das planilhas das lojas")
    canvas.drawRightString(A4[0] - 15 * mm, 10 * mm, f"página {doc.page}")
    canvas.restoreState()


def salvar(caminho, r):
    doc = SimpleDocTemplate(caminho, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=14 * mm, bottomMargin=16 * mm,
                            title=f"Relatório de vendas {r['mes_extenso']}", author="automacao-relatorio-vendas")
    metade = LARGURA / 2
    destaques = Table([[Paragraph("• " + esc(f), CORPO)] for f in r["destaques"]], colWidths=[LARGURA])
    destaques.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO), ("LINEBEFORE", (0, 0), (0, -1), 3, PRIMARIA),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
    ]))
    graficos = Table([[grafico_meta(r, metade - 6), grafico_categorias(r, metade - 6)]],
                     colWidths=[metade, metade])
    graficos.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))

    historia = [
        Paragraph(f"Relatório de vendas · {esc(r['mes_extenso'].capitalize())}", TITULO),
        Paragraph("Farmácia Vitalis Manipulação · consolidado das lojas Centro, Umarizal e Ananindeua", SUB),
        Spacer(1, 8),
        destaques,
        Spacer(1, 8),
        kpis(r),
        Paragraph("Resultado por loja", SECAO),
        tabela_lojas(r),
        Spacer(1, 10),
        *([graficos, grafico_diario(r, LARGURA)] if r["diario"] and r["total"]["bruto"] else
          [Paragraph("Nenhuma venda utilizável neste mês — veja abaixo o que aconteceu com cada planilha.",
                     ParagraphStyle("vazio", parent=CORPO, textColor=VERMELHO))]),
        KeepTogether([
            Paragraph("Qualidade dos dados recebidos", SECAO),
            Paragraph("Cada loja exporta de um sistema diferente. Antes de somar, o robô corrige e registra o que "
                      "encontrou — pra ninguém confiar num número sem saber de onde ele veio.", PEQUENO),
            Spacer(1, 4),
            tabela_qualidade(r),
        ]),
    ]
    doc.build(historia, onFirstPage=rodape, onLaterPages=rodape)
