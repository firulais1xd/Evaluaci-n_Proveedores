"""Generación del resumen general (.xlsx), el informe de gestión (.docx) y las fichas por proveedor (.docx)."""
from __future__ import annotations

import io
import textwrap

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Inches, Pt, RGBColor  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

from .analisis import (CAT_HEX, CAT_LABEL, CATS, PROC, action_rows, consolidated_paras, crit_rows, fnum,  # noqa: E402
                       history_for, pct, pp, provider_insights, rank_in, share, type_paras)

GREY = "#C5CCC8"
NAVY = "1F3B5C"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": "#BBBBBB",
                     "axes.spines.top": False, "axes.spines.right": False})


# ================================================================ gráficas
def _png(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return buf.getvalue()


def _bar_labels(ax, bars, fmt=lambda v: f"{v:g}", horizontal=False, size=8):
    for b in bars:
        v = b.get_width() if horizontal else b.get_height()
        if horizontal:
            ax.text(b.get_x() + v + 1, b.get_y() + b.get_height() / 2, fmt(v), va="center", fontsize=size,
                    fontweight="bold", color="#333")
        else:
            ax.text(b.get_x() + b.get_width() / 2, v, fmt(v), ha="center", va="bottom", fontsize=size,
                    fontweight="bold", color="#333")


def chart_categories(t, prev_t=None, prev_label="", cur_label="") -> bytes:
    """Barras de proveedores por categoría para un tipo (con el periodo anterior en gris si existe)."""
    fig, ax = plt.subplots(figsize=(4.2, 2.6))
    x = range(len(CATS))
    labels = [c.replace("MUY ", "MUY\n") for c in CATS]
    cur = [t.counts[c] for c in CATS]
    if prev_t is not None:
        w = 0.38
        b1 = ax.bar([i - w / 2 for i in x], [prev_t.counts[c] for c in CATS], w, color=GREY, label=prev_label)
        b2 = ax.bar([i + w / 2 for i in x], cur, w, color=[CAT_HEX[c] for c in CATS], label=cur_label)
        _bar_labels(ax, b1, size=7)
        _bar_labels(ax, b2, size=7)
        ax.legend(handles=[b1, b2[0]], labels=[prev_label, cur_label], fontsize=7, frameon=False, loc="upper right")
    else:
        b = ax.bar(list(x), cur, 0.6, color=[CAT_HEX[c] for c in CATS])
        _bar_labels(ax, b)
    ax.set_xticks(list(x), labels, fontsize=7)
    ax.yaxis.get_major_locator().set_params(integer=True)
    ax.tick_params(axis="y", labelsize=7)
    ax.set_ylim(0, max(1, *cur, *([prev_t.counts[c] for c in CATS] if prev_t else [0])) * 1.2)
    ax.grid(axis="y", color="#E6E9E7")
    ax.set_axisbelow(True)
    return _png(fig)


def chart_distribution(a, cmp=None) -> bytes:
    fig, ax = plt.subplots(figsize=(7, 3.3))
    x = list(range(len(CATS)))
    cur = [share(a["counts"][c], a["total"]) for c in CATS]
    fmt = lambda v: fnum(v, 1) + "%"
    if cmp:
        pa = cmp["prev"]
        w = 0.38
        b1 = ax.bar([i - w / 2 for i in x], [share(pa["counts"][c], pa["total"]) for c in CATS], w, color=GREY)
        b2 = ax.bar([i + w / 2 for i in x], cur, w, color=[CAT_HEX[c] for c in CATS])
        _bar_labels(ax, b1, fmt)
        _bar_labels(ax, b2, fmt)
        ax.legend([b1, b2[0]], [cmp["prevLabel"], a["label"]], frameon=False, fontsize=8)
    else:
        b = ax.bar(x, cur, 0.6, color=[CAT_HEX[c] for c in CATS])
        _bar_labels(ax, b, fmt)
    ax.set_xticks(x, CATS, fontsize=8)
    ax.set_ylim(0, 105)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.grid(axis="y", color="#E6E9E7")
    ax.set_axisbelow(True)
    ax.set_title("CONSOLIDADO POR CATEGORÍA (%)", fontsize=11, fontweight="bold", color="#333")
    return _png(fig)


def chart_areas(a) -> bytes:
    areas = sorted(a["procs"], key=lambda g: g["avg"])
    fig, ax = plt.subplots(figsize=(7, 0.42 * len(areas) + 1.4))
    y = list(range(len(areas)))
    left = [0.0] * len(areas)
    for c in CATS:
        vals = [share(g["counts"][c], g["n"]) for g in areas]
        ax.barh(y, vals, left=left, color=CAT_HEX[c], label=CAT_LABEL[c], height=0.62, edgecolor="white", linewidth=1)
        left = [l + v for l, v in zip(left, vals)]
    for i, g in enumerate(areas):
        ax.text(102, i, pct(g["avg"]), va="center", fontsize=8.5, fontweight="bold", color="#222")
    ax.set_yticks(y, [f"{g['name']} ({g['n']})" for g in areas], fontsize=8.5)
    ax.set_xlim(0, 100)
    ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.tick_params(axis="x", labelsize=8)
    ax.legend(ncol=5, fontsize=7.5, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.08))
    ax.set_title("CALIFICACIÓN POR ÁREA (distribución por categoría y promedio)", fontsize=11, fontweight="bold",
                 color="#333")
    return _png(fig)


def chart_process_compare(a, cmp) -> tuple[bytes, int] | None:
    pmap = {g["k"]: g for g in cmp["prev"]["procs"]}
    procs = [g for g in a["procs"] if g["k"] in pmap]
    if not procs:
        return None
    procs = sorted(procs, key=lambda g: g["avg"])
    fig, ax = plt.subplots(figsize=(7, 0.5 * len(procs) + 1.3))
    y = list(range(len(procs)))
    h = 0.38
    b1 = ax.barh([i + h / 2 for i in y], [pmap[g["k"]]["avg"] for g in procs], h, color=GREY, label=cmp["prevLabel"])
    b2 = ax.barh([i - h / 2 for i in y], [g["avg"] for g in procs], h, color="#1F4F7A", label=a["label"])
    lo = max(0, (min(min(g["avg"], pmap[g["k"]]["avg"]) for g in procs) // 10) * 10 - 10)
    for bars in (b1, b2):
        for b in bars:
            ax.text(b.get_width() + 0.3, b.get_y() + b.get_height() / 2, fnum(b.get_width()), va="center",
                    fontsize=7.5, color="#333")
    ax.set_yticks(y, [g["name"] for g in procs], fontsize=8.5)
    ax.set_xlim(lo, 104)
    ax.legend(fontsize=8, frameon=False, loc="lower right")
    ax.set_title("PROMEDIO POR PROCESO (%)", fontsize=11, fontweight="bold", color="#333")
    ax.grid(axis="x", color="#E6E9E7")
    ax.set_axisbelow(True)
    return _png(fig), len(procs)


def chart_criteria(t, r) -> bytes | None:
    cr = crit_rows(t, r)
    if not cr:
        return None
    cr = cr[::-1]
    fig, ax = plt.subplots(figsize=(7, 0.55 * len(cr) + 0.9))
    y = list(range(len(cr)))
    labels = ["\n".join(textwrap.wrap(c["h"], 42)[:3]) for c in cr]
    if t.n > 1:
        h = 0.38
        b1 = ax.barh([i + h / 2 for i in y], [c["v"] for c in cr], h, color=[CAT_HEX[_cat(c["v"])] for c in cr],
                     label="Proveedor")
        b2 = ax.barh([i - h / 2 for i in y], [c["avg"] for c in cr], h, color=GREY, label="Promedio del tipo")
        _bar_labels(ax, b1, lambda v: fnum(v) + "%", horizontal=True, size=7.5)
        _bar_labels(ax, b2, lambda v: fnum(v) + "%", horizontal=True, size=7)
        ax.legend(handles=[b1[0], b2], labels=["Proveedor", "Promedio del tipo"], fontsize=7.5, frameon=False,
                  loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=2)
    else:
        b1 = ax.barh(y, [c["v"] for c in cr], 0.55, color=[CAT_HEX[_cat(c["v"])] for c in cr])
        _bar_labels(ax, b1, lambda v: fnum(v) + "%", horizontal=True, size=7.5)
    ax.set_yticks(y, labels, fontsize=7.5)
    ax.set_xlim(0, 112)
    ax.set_xticks([0, 20, 40, 60, 80, 100], ["0%", "20%", "40%", "60%", "80%", "100%"], fontsize=7)
    ax.grid(axis="x", color="#E6E9E7")
    ax.set_axisbelow(True)
    return _png(fig)


def _cat(v):
    from .analisis import cat_of
    return cat_of(v)


# ================================================================ Excel
def build_xlsx(a, cmp=None, multi=None) -> bytes:
    wb = Workbook()
    thin = Side(style="thin", color="D0D5D2")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    hdr_fill = PatternFill("solid", fgColor=NAVY)
    soft = PatternFill("solid", fgColor="EDF0EE")

    def header(ws, row, values):
        for i, v in enumerate(values, 1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
        ws.row_dimensions[row].height = 30

    def widths(ws, ws_widths):
        for i, w in enumerate(ws_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # General: mismo esquema que el Resumen General de la plataforma
    g = wb.active
    g.title = "General"
    widths(g, [88] + [13] * 5 + [10])
    g.merge_cells("A2:G2")
    g["A2"] = "Resumen General Evaluacion Proveedores Categorias"
    g["A2"].font = Font(bold=True, size=14)
    g.merge_cells("A3:G3")
    g["A3"] = f"Periodo: {a['label']}"
    g["A3"].font = Font(bold=True)
    header(g, 5, ["Tipo de Evaluación", *CATS, "TOTAL"])
    for i, c in enumerate(CATS):
        g.cell(row=5, column=i + 2).fill = PatternFill("solid", fgColor=CAT_HEX[c][1:])
    r = 6
    for t in a["types"]:
        for j, v in enumerate([t.title, *[t.counts[c] for c in CATS], t.n], 1):
            cell = g.cell(row=r, column=j, value=v)
            cell.border = border
            if j == 1:
                cell.alignment = Alignment(wrap_text=True)
        r += 1
    for j, v in enumerate(["TOTAL", *[a["counts"][c] for c in CATS], a["total"]], 1):
        cell = g.cell(row=r, column=j, value=v)
        cell.font, cell.fill, cell.border = Font(bold=True), soft, border
    r += 1
    for j, v in enumerate(["% del total", *[a["counts"][c] / a["total"] for c in CATS], 1], 1):
        cell = g.cell(row=r, column=j, value=v)
        cell.font, cell.fill, cell.border = Font(bold=True), soft, border
        if j > 1:
            cell.number_format = "0.0%"
    g.freeze_panes = "A6"

    # Por tipo
    s = wb.create_sheet("Por tipo")
    cols = [("Proceso", 20), ("Código", 13), ("Tipo de proveedor", 40), ("No. proveedores", 11), ("Promedio", 10),
            ("Mínimo", 10), ("Proveedor con menor calificación", 40), ("Máximo", 10),
            ("Proveedor con mayor calificación", 40), *[(c, 11) for c in CATS], ("Criterio más bajo", 50),
            ("Promedio criterio", 11)]
    if cmp:
        cols += [(f"Promedio {cmp['prevLabel']}", 12), ("Variación (pp)", 11)]
    header(s, 1, [c[0] for c in cols])
    widths(s, [c[1] for c in cols])
    for t in a["types"]:
        c = cmp["byType"].get(t.code) if cmp else None
        row = [PROC[t.proc], t.code, t.type_name, t.n, t.avg, t.min,
               "; ".join(x["name"] for x in t.rows if x["final"] == t.min), t.max,
               "; ".join(x["name"] for x in t.rows if x["final"] == t.max), *[t.counts[k] for k in CATS],
               t.weakest["h"] if t.weakest else "", t.weakest["avg"] if t.weakest else None]
        if cmp:
            row += [c["prev"].avg if c and c["prev"] else None, t.avg - c["prev"].avg if c and c["prev"] else None]
        s.append(row)
    for row in s.iter_rows(min_row=2):
        for k in (5, 6, 8, 16, 17, 18):
            if k <= len(row):
                row[k - 1].number_format = "0.00"
    s.freeze_panes = "A2"
    s.auto_filter.ref = s.dimensions

    # Detalle
    d = wb.create_sheet("Detalle")
    cols = [("Proceso", 20), ("Código", 13), ("Tipo de proveedor", 40), ("NIT", 18), ("Proveedor", 50), ("Final", 9),
            ("Categoría", 13), ("Final plataforma", 11), ("Categoría plataforma", 13), ("Criterios en blanco", 40),
            ("Observación", 45)]
    header(d, 1, [c[0] for c in cols])
    widths(d, [c[1] for c in cols])
    for t in a["types"]:
        for x in t.rows:
            fp = x.get("finalPlat", x["final"])
            obs = []
            if not x["platCat"]:
                obs.append("Sin categoría en la plataforma")
            elif x["platCat"] != _cat(fp):
                obs.append("Categoría de la plataforma no coincide con su Final")
            if x.get("recalc"):
                obs.append("Final recalculado sin criterios en blanco" + (" (peso estimado)" if x.get("estimated") else ""))
            d.append([PROC[t.proc], t.code, t.type_name, x["nit"], x["name"], x["final"], x["cat"], fp, x["platCat"],
                      "; ".join(x.get("blanks", [])), "; ".join(obs)])
            d.cell(row=d.max_row, column=8).number_format = "0.00"
            d.cell(row=d.max_row, column=6).number_format = "0.00"
            d.cell(row=d.max_row, column=7).font = Font(bold=True, color=CAT_HEX[x["cat"]][1:])
    d.freeze_panes = "A2"
    d.auto_filter.ref = d.dimensions

    # Comparativo
    if cmp:
        c = wb.create_sheet("Comparativo")
        pl, cl = cmp["prevLabel"], a["label"]
        cols = [("Proceso", 20), ("Código", 13), ("Tipo de proveedor", 40), ("NIT", 18), ("Proveedor", 50),
                (f"Final {pl}", 12), (f"Final {cl}", 12), ("Variación (pp)", 12), (f"Categoría {pl}", 13),
                (f"Categoría {cl}", 13), ("Estado", 14)]
        header(c, 1, [x[0] for x in cols])
        widths(c, [x[1] for x in cols])
        colors = {"Mejoró": "11774A", "Bajó": "B3263E"}
        for t in a["types"]:
            b = cmp["byType"][t.code]
            for m in b["matched"]:
                st = "Mejoró" if m["d"] > 0.05 else "Bajó" if m["d"] < -0.05 else "Se mantuvo"
                c.append([PROC[t.proc], t.code, t.type_name, m["r"]["nit"], m["r"]["name"], m["pr"]["final"],
                          m["r"]["final"], m["d"], m["pr"]["cat"], m["r"]["cat"], st])
                c.cell(row=c.max_row, column=11).font = Font(bold=True, color=colors.get(st, "666666"))
            for x in b["news"]:
                c.append([PROC[t.proc], t.code, t.type_name, x["nit"], x["name"], None, x["final"], None, "", x["cat"],
                          "Nuevo"])
            for x in b["gone"]:
                c.append([PROC[t.proc], t.code, t.type_name, x["nit"], x["name"], x["final"], None, None, x["cat"], "",
                          "No evaluado"])
        for t in cmp["goneTypes"]:
            for x in t.rows:
                c.append([PROC[t.proc], t.code, t.type_name, x["nit"], x["name"], x["final"], None, None, x["cat"], "",
                          "No evaluado"])
        for row in c.iter_rows(min_row=2):
            for k in (6, 7, 8):
                row[k - 1].number_format = "0.00"
        c.freeze_panes = "A2"
        c.auto_filter.ref = c.dimensions

    if multi:  # varios periodos comparados: tablas por periodo
        from .historico import add_period_sheets
        add_period_sheets(wb, multi, base=cmp["prevLabel"] if cmp else None)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ================================================================ Word: utilidades
def _shade(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def _row_flags(row, cant_split=True, header=False):
    tr_pr = row._tr.get_or_add_trPr()
    if cant_split:
        tr_pr.append(OxmlElement("w:cantSplit"))
    if header:
        tr_pr.append(OxmlElement("w:tblHeader"))


def _widths(table, widths_in):
    table.autofit = False
    # ancho de la cuadrícula (lo usan LibreOffice y Google Docs) y de cada celda (lo usa Word)
    for gc, w in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths_in):
        gc.set(qn("w:w"), str(int(w * 1440)))
    for row in table.rows:
        for i, w in enumerate(widths_in):
            if i < len(row.cells):
                row.cells[i].width = Inches(w)


def _para(container, text="", bold=False, size=None, color=None, align=None, keep=False, space_after=4):
    p = container.add_paragraph()
    if text:
        run = p.add_run(text)
        run.bold = bold
        if size:
            run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.keep_with_next = keep
    return p


def _label(container, label, value, keep=False):
    p = container.add_paragraph()
    p.add_run(label).bold = True
    p.add_run(value)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = keep
    return p


def _first_para(cell):
    """Las celdas nuevas traen un párrafo vacío; lo usamos o lo quitamos."""
    return cell.paragraphs[0]


def _cell_text(cell, text, bold=False, color=None, size=None, align=None):
    p = _first_para(cell)
    run = p.add_run(text)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if size:
        run.font.size = Pt(size)
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(2)


def _cell_paras(cell, items):
    """items: lista de str o tuplas ('label', etiqueta, valor) / ('bold', texto)."""
    first = True
    for it in items:
        p = _first_para(cell) if first else cell.add_paragraph()
        first = False
        if isinstance(it, tuple) and it[0] == "label":
            p.add_run(it[1]).bold = True
            p.add_run(it[2])
        elif isinstance(it, tuple) and it[0] == "bold":
            p.add_run(it[1]).bold = True
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(3)


def _head_table(doc, title, cols=1, widths=(7.0,)):
    t = doc.add_table(rows=1, cols=cols)
    t.style = "Table Grid"
    row = t.rows[0]
    cell = row.cells[0].merge(row.cells[-1]) if cols > 1 else row.cells[0]
    _cell_text(cell, title, bold=True, color="FFFFFF")
    _first_para(cell).paragraph_format.keep_with_next = True
    _shade(cell, NAVY)
    _row_flags(row)
    return t


def _new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Inches(0.75)
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    return doc


def _save(doc) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ================================================================ Word: informe
def build_docx(a, cmp=None, meta=None, progress=None) -> bytes:
    meta = meta or {}
    doc = _new_doc()
    _para(doc, "INFORME DE GESTIÓN", bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    t = doc.add_table(rows=3, cols=4)
    t.style = "Table Grid"
    vals = [["Proceso o programa:", "Gestión de proveedores", None, None],
            ["Período evaluado:", a["label"], "Fecha de evaluación:", meta.get("fecha", "")],
            ["Informe elaborado por:", meta.get("autor", ""), None, None]]
    for i, row in enumerate(t.rows):
        if vals[i][2] is None:
            row.cells[1].merge(row.cells[3])
        for j, v in enumerate(vals[i]):
            if v is None:
                continue
            _cell_text(row.cells[j], v, bold=j in (0, 2))
            if j in (0, 2):
                _shade(row.cells[j], "EDF0EE")
    _widths(t, [1.6, 2.1, 1.6, 1.7])

    def box(title, items):
        _para(doc, space_after=2)
        bt = _head_table(doc, title)
        cell = bt.add_row().cells[0]
        _cell_paras(cell, items)

    box("Objetivos del proceso o programa:", ["Evaluar el desempeño de los proveedores con los que cuenta la "
                                              "organización en los distintos procesos, en conformidad con los "
                                              "requisitos establecidos."])
    box("Resultados de la gestión", [("label", "Aspecto verificado: ",
                                      f"Resultados de evaluación de proveedores en el periodo {a['label']}."),
                                     "Con base en la información con la que se cuenta, se obtuvieron los siguientes "
                                     "resultados:", *(meta.get("nota_final") and [meta["nota_final"]] or [])])

    total_types = len(a["types"])
    done = 0
    for g in a["procs"]:
        _para(doc, space_after=2)
        tb = _head_table(doc, f"PROVEEDORES {g['name'].upper()}  ·  promedio {pct(g['avg'])}", cols=2)
        for tt in g["types"]:
            c = cmp["byType"].get(tt.code) if cmp else None
            png = chart_categories(tt, c["prev"] if c else None, cmp["prevLabel"] if cmp else "", a["label"])
            row = tb.add_row()
            _row_flags(row)
            left, right = row.cells
            left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _first_para(left).add_run().add_picture(io.BytesIO(png), width=Inches(2.75))
            _first_para(left).alignment = WD_ALIGN_PARAGRAPH.CENTER
            _cell_paras(right, [("label", "Tipo de proveedor: ", tt.type_name + "."),
                                ("label", "No. de proveedores: ", str(tt.n)),
                                ("bold", "Análisis de resultados:"), *type_paras(tt, cmp)])
            done += 1
            if progress:
                progress(done / (total_types + 3))
        _widths(tb, [2.95, 4.05])

    # Consolidados
    _para(doc, space_after=2)
    ct = _head_table(doc, "CONSOLIDADOS")
    for png, w in ((chart_distribution(a, cmp), 5.8), (chart_areas(a), 6.2)):
        row = ct.add_row()
        _row_flags(row)
        p = _first_para(row.cells[0])
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(io.BytesIO(png), width=Inches(w))
    cell = ct.add_row().cells[0]
    _cell_paras(cell, [("label", "No. de evaluaciones: ", f"{a['total']} ({a['uniq']} proveedores distintos)"),
                       ("bold", "Análisis de resultados:"), *consolidated_paras(a, cmp)])
    if progress:
        progress((total_types + 1) / (total_types + 3))
    if cmp:
        res = chart_process_compare(a, cmp)
        if res:
            row = ct.add_row()
            _row_flags(row)
            _cell_paras(row.cells[0], [("bold", f"Comparativo {cmp['prevLabel']} – {a['label']}")])
            p = row.cells[0].add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(io.BytesIO(res[0]), width=Inches(6.0))
        drops = [m for m in cmp["movers"] if m["d"] <= -5]
        if drops:
            cell = ct.add_row().cells[0]
            _cell_paras(cell, [("bold", "Proveedores con mayor caída (5 pp o más):")] + [
                f"• {m['r']['name']} ({m['t'].type_name}): de {pct(m['pr']['final'])} a {pct(m['r']['final'])} "
                f"({pp(m['d'])})." for m in drops[:12]])

    # Resumen general
    _para(doc, space_after=2)
    rg = doc.add_table(rows=2, cols=7)
    rg.style = "Table Grid"
    head = rg.rows[0].cells[0].merge(rg.rows[0].cells[6])
    _cell_text(head, "RESUMEN GENERAL POR TIPO DE EVALUACIÓN", bold=True, color="FFFFFF", size=8)
    _shade(head, NAVY)
    for j, h in enumerate(["Tipo de evaluación", *CATS, "TOTAL"]):
        _cell_text(rg.rows[1].cells[j], h, bold=True, size=7.5, align=None if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
        _shade(rg.rows[1].cells[j], "EDF0EE")
    _row_flags(rg.rows[1], header=True)
    for tt in a["types"]:
        row = rg.add_row()
        _row_flags(row)
        for j, v in enumerate([tt.title, *[tt.counts[c] for c in CATS], tt.n]):
            _cell_text(row.cells[j], str(v), size=7.5, bold=j == 6, align=None if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    row = rg.add_row()
    for j, v in enumerate(["TOTAL", *[a["counts"][c] for c in CATS], a["total"]]):
        _cell_text(row.cells[j], str(v), size=7.5, bold=True, align=None if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
        _shade(row.cells[j], "EDF0EE")
    _widths(rg, [3.1, 0.66, 0.66, 0.64, 0.64, 0.64, 0.66])

    multi = meta.get("multi")
    if multi:  # tabla de promedio por área en todos los periodos comparados
        from .historico import area_matrix, insights
        labels = [l for l, _ in multi]
        am = area_matrix(multi)
        _para(doc, space_after=2)
        mt = doc.add_table(rows=2, cols=len(labels) + 1)
        mt.style = "Table Grid"
        head = mt.rows[0].cells[0].merge(mt.rows[0].cells[-1])
        _cell_text(head, "COMPARATIVO DE VARIOS PERIODOS · PROMEDIO POR ÁREA (%)", bold=True, color="FFFFFF", size=8)
        _shade(head, NAVY)
        for j, h in enumerate(["Área", *labels]):
            _cell_text(mt.rows[1].cells[j], h, bold=True, size=8, align=None if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
            _shade(mt.rows[1].cells[j], "EDF0EE")
        for area, v in am.items():
            row = mt.add_row()
            _row_flags(row)
            _cell_text(row.cells[0], area, size=8)
            for j, l in enumerate(labels, 1):
                val = v.get(l, (None,))[0]
                _cell_text(row.cells[j], fnum(val) if val is not None else "—", size=8,
                           align=WD_ALIGN_PARAGRAPH.CENTER,
                           color=CAT_HEX[_cat(val)][1:] if val is not None else None, bold=val is not None)
        _widths(mt, [2.6] + [4.4 / len(labels)] * len(labels))
        cell = mt.add_row().cells
        cell = cell[0].merge(cell[-1])
        _cell_paras(cell, [("bold", "Análisis de los periodos comparados:"), *insights(multi)])

    box("Conclusiones", [s.strip() for s in (meta.get("conclusiones") or "").split("\n") if s.strip()] or ["—"])

    _para(doc, space_after=2)
    at = doc.add_table(rows=2, cols=4)
    at.style = "Table Grid"
    head = at.rows[0].cells[0].merge(at.rows[0].cells[3])
    _cell_text(head, "Acciones para la mejora (Tipo de acción  C: correcciones / AC: acciones correctivas / "
                     "CR: control riesgo / OP: aprovechamiento de oportunidad)", bold=True, color="FFFFFF")
    _shade(head, NAVY)
    for j, h in enumerate(["Tipo de acción", "Descripción breve", "Fecha (dd/mm/aa)", "Resp."]):
        _cell_text(at.rows[1].cells[j], h, bold=True)
        _shade(at.rows[1].cells[j], "EDF0EE")
    for tipo, desc in action_rows(a):
        row = at.add_row()
        for j, v in enumerate([tipo, desc, "NA" if tipo == "NA" else "", "NA" if tipo == "NA" else ""]):
            _cell_text(row.cells[j], v)
    _widths(at, [1.05, 3.85, 1.05, 1.05])
    if progress:
        progress(1.0)
    return _save(doc)


# ================================================================ Word: fichas
def build_fichas(a, cmp, providers, timeline, proc_only=None, progress=None) -> bytes:
    doc = _new_doc()
    _para(doc, "FICHAS DE DESEMPEÑO DE PROVEEDORES", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    extra = f" · comparado con {cmp['prevLabel']}" if cmp else ""
    _para(doc, f"Periodo evaluado: {a['label']}{extra} · {len(providers)} "
               f"{'proveedor' if len(providers) == 1 else 'proveedores'}", color="555555",
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    ordered = sorted(providers, key=lambda p: (min(r["final"] for _, r in p["evals"]), p["name"]))
    for i, p in enumerate(ordered):
        evs = [(t, r) for t, r in p["evals"] if not proc_only or t.proc == proc_only]
        h = _para(doc, p["name"], bold=True, size=13, color=NAVY, keep=True, space_after=0)
        h.paragraph_format.page_break_before = i > 0
        _para(doc, f"NIT {p['nit']}", color="555555", keep=True, space_after=8)
        for t, r in evs:
            others = [(ot, orow) for ot, orow in p["evals"] if ot is not t]
            hist = history_for(timeline, p["key"], t.code)
            q = _para(doc, keep=True, space_after=2)
            q.add_run(t.type_name + " ").bold = True
            rr = q.add_run(f"({PROC[t.proc]} · {t.code})")
            rr.font.color.rgb = RGBColor.from_string("666666")
            rr.font.size = Pt(9)
            q = _para(doc, keep=True, space_after=4)
            q.add_run("Calificación final: ").bold = True
            rr = q.add_run(pct(r["final"]) + "  ")
            rr.bold = True
            rr.font.size = Pt(12)
            rr.font.color.rgb = RGBColor.from_string(CAT_HEX[r["cat"]][1:])
            rest = CAT_LABEL[r["cat"]].upper()
            if t.n > 1:
                rest += f"  ·  puesto {rank_in(t, r)} de {t.n}  ·  promedio del tipo {pct(t.avg)}"
            q.add_run(rest)
            png = chart_criteria(t, r)
            if png:
                q = _para(doc, keep=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
                q.add_run().add_picture(io.BytesIO(png), width=Inches(5.9))
            _para(doc, "Análisis:", bold=True, keep=True, space_after=2)
            for s in provider_insights(t, r, others, hist, cmp):
                b = doc.add_paragraph(s, style="List Bullet")
                b.paragraph_format.space_after = Pt(2)
        if progress:
            progress((i + 1) / len(ordered))
    return _save(doc)


# ================================================================ Power BI
def build_powerbi_xlsx(timeline) -> bytes:
    """Tablas limpias para Power BI (mismos nombres que el modelo de las consultas M).

    timeline: lista de (etiqueta_periodo, análisis) en orden cronológico.
    """
    from openpyxl.worksheet.table import Table, TableStyleInfo

    from .analisis import PROC_ORDER, pkey

    evals, crits, tipos, provs = [], [], {}, {}
    for label, a in timeline:
        for t in a["types"]:
            tipos.setdefault(t.code, [t.code, t.title, t.type_name, t.proc])
            for r in t.rows:
                k = pkey(r)
                ek = f"{label}|{t.code}|{k}"
                fp = r.get("finalPlat", r["final"])
                obs = ("Sin categoría en la plataforma" if not r["platCat"] else
                       f"Categoría de la plataforma ({r['platCat']}) no coincide con Final"
                       if r["platCat"] != _cat(fp) else
                       "Final recalculado sin criterios en blanco" if r.get("recalc") else None)
                evals.append([label, t.sheet, t.code, t.title, r["nit"], r["name"], r["final"], r["platCat"], k, ek,
                              r["cat"], obs, fp, "; ".join(r.get("blanks", []))])
                provs[k] = [k, r["nit"], r["name"], f"{r['name']} · {r['nit']}"]  # el último periodo gana
                for c in t.crit:
                    v = r["scores"][c["j"]] if c["j"] < len(r["scores"]) else None
                    if v is not None:
                        crits.append([ek, label, t.code, k, c["h"], c["j"] + 1, v])
    periodos = [[lbl, i + 1] for i, (lbl, _) in enumerate(timeline)]
    sheets = {
        "Evaluaciones": (["Periodo", "Hoja", "Codigo", "Titulo", "NIT", "Proveedor", "Final", "CategoriaPlataforma",
                          "ProveedorKey", "EvalKey", "Categoria", "Observacion", "FinalPlataforma",
                          "CriteriosEnBlanco"], evals),
        "Criterios": (["EvalKey", "Periodo", "Codigo", "ProveedorKey", "Criterio", "Orden", "Puntaje"], crits),
        "Dim_Tipo": (["Codigo", "Titulo", "Tipo", "Proceso"], list(tipos.values())),
        "Dim_Proceso": (["Proceso", "Area", "Orden"], [[k, PROC[k], i + 1] for i, k in enumerate(PROC_ORDER)]),
        "Dim_Categoria": (["Categoria", "Nombre", "Orden", "Desde", "Hasta", "Singular", "Plural", "Color"],
                          [["EXCELENTE", "Excelente", 1, 98, "100%", "excelente", "excelentes", "#11774A"],
                           ["MUY BUENO", "Muy bueno", 2, 80, "< 98%", "muy bueno", "muy buenos", "#4FA877"],
                           ["BUENO", "Bueno", 3, 70, "< 80%", "bueno", "buenos", "#CFA11D"],
                           ["REGULAR", "Regular", 4, 60, "< 70%", "regular", "regulares", "#D65F28"],
                           ["MALO", "Malo", 5, 0, "< 60%", "malo", "malos", "#A3233A"]]),
        "Dim_Periodo": (["Periodo", "Orden"], periodos),
        "Periodo_Comparar": (["Periodo comparación", "Orden comparación"], periodos),
        "Dim_Proveedor": (["ProveedorKey", "NIT", "Proveedor", "Proveedor y NIT"], list(provs.values())),
    }
    wb = Workbook()
    wb.remove(wb.active)
    for name, (cols, rows) in sheets.items():
        ws = wb.create_sheet(name)
        ws.append(cols)
        for r in rows:
            ws.append(r)
        ref = f"A1:{get_column_letter(len(cols))}{max(2, len(rows) + 1)}"
        tab = Table(displayName=name, ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tab)
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = max(12, min(50, len(c) + 4))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
