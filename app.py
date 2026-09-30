"""Evaluación de Proveedores — app local de Streamlit.

Ejecutar con:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import json
import re
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from core.analisis import (CAT_HEX, CAT_LABEL, CATS, PROC, action_rows, analyze, cat_of, compare, conclusion_text,
                           infer_weights, recalc_period, recalc_rows,
                           consolidated_paras, counts_text, crit_rows, fnum, from_stored, history_for, parse_workbook,
                           pct, pkey, pp, provider_index, provider_insights, rank_in, share, to_stored, type_paras)
from core.historico import (area_matrix, build_hist_xlsx, guess_label, insights, period_summary,
                            provider_matrix, recurrent_low, type_matrix)
from core.exportar import build_docx, build_fichas, build_powerbi_xlsx, build_xlsx

BASE = Path(__file__).parent
HIST = BASE / "historial"
HIST.mkdir(exist_ok=True)
CONF = BASE / "config"
CONF.mkdir(exist_ok=True)
POND = CONF / "ponderaciones.json"

st.set_page_config(page_title="Evaluación de Proveedores", page_icon="📋", layout="wide")
st.markdown("""
<style>
.block-container{padding-top:2rem;max-width:1300px}
.scale{display:flex;flex-wrap:wrap;gap:0;border-radius:6px;overflow:hidden;width:fit-content;font-size:12px;font-weight:600}
.scale span{padding:5px 10px;color:#fff}
.pill{display:inline-block;padding:1px 8px;border-radius:99px;font-size:12px;font-weight:600;color:#fff}
.small{font-size:13px;opacity:.8}
div[data-testid="stMetricValue"]{font-size:1.7rem}
</style>""", unsafe_allow_html=True)


# ---------------------------------------------------------------- historial
def natural(s: str):
    return [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", s)]


def slug(label: str) -> str:
    return re.sub(r"[^A-Za-z0-9_\-.]", "_", label.strip()) or "periodo"


def load_history() -> dict[str, dict]:
    out = {}
    for f in HIST.glob("*.json"):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            out[d["label"]] = d
        except Exception:  # archivo dañado: se ignora
            continue
    return dict(sorted(out.items(), key=lambda kv: natural(kv[0]), reverse=True))


def save_history(p: dict):
    (HIST / f"{slug(p['label'])}.json").write_text(json.dumps(to_stored(p), ensure_ascii=False), encoding="utf-8")


def delete_history(label: str):
    f = HIST / f"{slug(label)}.json"
    if f.exists():
        f.unlink()


@st.cache_data(show_spinner=False)
def parse_cached(data: bytes, label: str) -> dict:
    return parse_workbook(data, label)


def ppm(d: float) -> str:
    """Variación para st.metric (necesita el signo menos ASCII para colorear en rojo)."""
    return pp(d).replace("−", "-")


def pill(cat: str) -> str:
    fg = "#1c1a10" if cat == "BUENO" else "#fff"
    return f'<span class="pill" style="background:{CAT_HEX[cat]};color:{fg}">{CAT_LABEL[cat]}</span>'


# ---------------------------------------------------------------- barra lateral
hist = load_history()
with st.sidebar:
    st.header("Periodo")
    up = st.file_uploader("Archivo EvaluacionResumenPeriodo (.xlsx)", type=["xlsx"], key="up_cur")
    cur = None
    source = ""
    if up is not None:
        label = st.text_input("Nombre del periodo", placeholder="p. ej. 2026-1", key="label_cur").strip()
        try:
            cur = parse_cached(up.getvalue(), label or "Periodo actual")
            cur = {**cur, "label": label or "Periodo actual"}
            source = hashlib.md5(up.getvalue()).hexdigest()[:10]
        except Exception as e:  # archivo con otro formato
            st.error(str(e))
    elif hist:
        sel = st.selectbox("O abre un periodo guardado", list(hist), key="open_hist")
        cur = from_stored(hist[sel])
        source = f"hist:{sel}:{hist[sel].get('savedAt')}"

    fecha = st.date_input("Fecha de evaluación", value=date.today(), format="DD/MM/YYYY")
    autor = st.text_input("Informe elaborado por", placeholder="Nombre – cargo")
    modo_final = st.radio(
        "Criterios en blanco", ["No contarlos (recalcular el Final)", "Contarlos como 0 (Final de la plataforma)"],
        key="modo_final",
        help="La plataforma toma los criterios sin diligenciar como 0. Con la primera opción la app recalcula el "
             "Final de esas filas sin tenerlos en cuenta, conservando la ponderación de los demás criterios. "
             "Las filas completas siempre conservan el Final de la plataforma.")
    ignore_blanks = modo_final.startswith("No")

    prev = None
    prev_list = []
    if cur:
        st.header("Comparar con")
        others = [k for k in hist if k != cur["label"]]
        older = [k for k in others if natural(k) < natural(cur["label"])]
        sel_prev = st.multiselect(
            "Periodos a comparar", others, default=older[:1] if older else others[:1],
            key=f"cmp_multi_{cur['label']}", placeholder="Sin comparación",
            help="Puedes elegir varios. Las variaciones (pp, mejoraron, bajaron) se calculan frente al periodo "
                 "anterior más reciente; los demás aparecen como columnas y series adicionales.")
        prev_list = [from_stored(hist[k]) for k in sel_prev]
        if st.checkbox("Agregar un archivo que no está en el historial", key="cmp_file_chk"):
            up2 = st.file_uploader("Archivo del otro periodo", type=["xlsx"], key="up_prev")
            label2 = st.text_input("Nombre de ese periodo", placeholder="p. ej. 2025-2", key="label_prev").strip()
            if up2 is not None:
                try:
                    extra = {**parse_cached(up2.getvalue(), label2 or "Periodo anterior"),
                             "label": label2 or "Periodo anterior"}
                    prev_list = [p for p in prev_list if p["label"] != extra["label"]] + [extra]
                except Exception as e:
                    st.error(str(e))
        prev_list = sorted(prev_list, key=lambda p: natural(p["label"]))
        if prev_list:
            before = [p for p in prev_list if natural(p["label"]) < natural(cur["label"])]
            prev = before[-1] if before else prev_list[-1]
            if len(prev_list) > 1:
                st.caption(f"Variaciones frente a **{prev['label']}**; "
                           + ", ".join(p["label"] for p in prev_list if p is not prev)
                           + " se muestran como columnas adicionales.")

        st.header("Historial")
        if up is not None:
            if st.button("Guardar este periodo en el historial", type="primary", width="stretch"):
                if not st.session_state.get("label_cur", "").strip():
                    st.warning("Escribe el nombre del periodo antes de guardarlo.")
                else:
                    existed = cur["label"] in hist
                    save_history(cur)
                    st.success(f"Periodo {cur['label']} {'actualizado' if existed else 'guardado'}.")
                    st.rerun()
        if hist:
            st.caption("Guardados: " + ", ".join(hist))
            with st.expander("Eliminar un periodo guardado"):
                dl = st.selectbox("Periodo", list(hist), key="del_sel")
                if st.button(f"Eliminar {dl}", key="del_btn"):
                    delete_history(dl)
                    st.rerun()
        else:
            st.caption("El historial está vacío. Guarda cada periodo para poder compararlo después.")

    with st.expander("Cargar varios periodos al historial"):
        st.caption("Sube los archivos de años o semestres anteriores (uno por periodo) para compararlos en la "
                   "pestaña Histórico. Revisa el nombre de cada periodo antes de guardar.")
        ups = st.file_uploader("Archivos EvaluacionResumenPeriodo", type=["xlsx"], accept_multiple_files=True,
                               key="up_multi")
        pend = []
        for i, f in enumerate(ups or []):
            lbl = st.text_input(f"Periodo de {f.name}", value=guess_label(f.name), key=f"multi_lbl_{f.name}_{i}",
                                placeholder="p. ej. 2023").strip()
            pend.append((f, lbl))
        if pend and st.button("Guardar todos en el historial", key="multi_save", type="primary", width="stretch"):
            faltan = [f.name for f, l in pend if not l]
            repetidos = {l for _, l in pend if l and [x for _, x in pend].count(l) > 1}
            if faltan:
                st.warning("Falta el nombre del periodo de: " + ", ".join(faltan))
            elif repetidos:
                st.warning("Hay nombres de periodo repetidos: " + ", ".join(sorted(repetidos)))
            else:
                ok, errs = [], []
                for f, l in pend:
                    try:
                        save_history({**parse_cached(f.getvalue(), l), "label": l})
                        ok.append(l)
                    except Exception as e:
                        errs.append(f"{f.name}: {e}")
                if errs:
                    st.error("No se pudieron leer: " + "; ".join(errs))
                if ok:
                    st.success("Guardados: " + ", ".join(ok))
                    st.rerun()

# ---------------------------------------------------------------- encabezado
st.title("Evaluación de Proveedores")
st.markdown('<div class="scale">' + "".join(
    f'<span style="background:{CAT_HEX[c]};{"color:#1c1a10" if c == "BUENO" else ""}">{CAT_LABEL[c]} {r}</span>'
    for c, r in zip(CATS, ["≥ 98%", "80–98%", "70–80%", "60–70%", "< 60%"])) + "</div>", unsafe_allow_html=True)

if not cur:
    st.info("Sube en la barra lateral el archivo **EvaluacionResumenPeriodo** que descargas de la plataforma "
            "(una hoja por tipo de evaluación, p. ej. EV-1-CE-DS-03) y escribe el nombre del periodo.")
    st.stop()

# ---------------------------------------------------------------- recálculo del Final
def load_overrides() -> dict:
    try:
        return json.loads(POND.read_text(encoding="utf-8"))
    except Exception:
        return {}


overrides = load_overrides()
hist_raw = {lbl: from_stored(d) for lbl, d in hist.items()}
raw_periods = [p for lbl, p in hist_raw.items() if lbl != cur["label"]] + [cur] + prev_list
wmap = infer_weights(raw_periods)


def prep(p):
    return recalc_period(p, wmap, overrides, ignore_blanks)


a = analyze(prep(cur))
pas_all = {p["label"]: analyze(prep(p)) for p in prev_list if p["label"] != a["label"]}
pa = pas_all.get(prev["label"]) if prev else None
cmp = compare(a, pa)
# periodos comparados en orden cronológico (incluye el actual); None si solo hay uno anterior
multi = sorted([*pas_all.items(), (a["label"], a)], key=lambda kv: natural(kv[0])) if len(pas_all) >= 2 else None

# línea de tiempo para las fichas: historial + periodo anterior + actual
timeline_map = {lbl: analyze(prep(p)) for lbl, p in hist_raw.items() if lbl != a["label"]}
for lbl, x in pas_all.items():
    timeline_map[lbl] = x
timeline_map[a["label"]] = a
timeline = sorted(timeline_map.items(), key=lambda kv: natural(kv[0]))

st.caption(f"Periodo **{a['label']}** · {a['total']} evaluaciones en {len(a['types'])} tipos"
           + (f" · comparado con **{', '.join(l for l, _ in multi if l != a['label'])}**" if multi else
              f" · comparado con **{cmp['prevLabel']}**" if cmp else ""))

tab_names = ["Consolidado", "Por área y tipo", "Ficha por proveedor"] + (["Comparativo"] if cmp else []) + \
            ["Histórico", "Cálculo del Final", "Revisión de datos", "Descargas"]
tabs = dict(zip(tab_names, st.tabs(tab_names)))


# ---------------------------------------------------------------- gráficas plotly
def fig_layout(fig, h):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor="rgba(0,0,0,0)",
                      paper_bgcolor="rgba(0,0,0,0)", legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0, traceorder="normal"),
                      hoverlabel=dict(font_size=12))
    return fig


def fig_distribution():
    fig = go.Figure()
    rows = multi if multi else ([(cmp["prevLabel"], pa)] if cmp else []) + [(a["label"], a)]
    for c in CATS:
        fig.add_bar(y=[r[0] for r in rows], x=[share(r[1]["counts"][c], r[1]["total"]) for r in rows],
                    orientation="h", name=CAT_LABEL[c], marker_color=CAT_HEX[c],
                    customdata=[r[1]["counts"][c] for r in rows],
                    hovertemplate="%{y} · " + CAT_LABEL[c] + ": %{customdata} (%{x:.1f}%)<extra></extra>")
    fig = fig_layout(fig, 100 + 45 * len(rows))
    fig.update_layout(barmode="stack", xaxis=dict(range=[0, 100], ticksuffix="%"), yaxis=dict(autorange="reversed", type="category"),
                      margin=dict(l=10, r=10, t=40, b=10))
    return fig


def fig_areas():
    areas = sorted(a["procs"], key=lambda g: g["avg"])
    pmap = {g["k"]: g for g in pa["procs"]} if pa else {}
    fig = go.Figure()
    names = [f"{g['name']} ({g['n']})" for g in areas]
    for c in CATS:
        fig.add_bar(y=names, x=[share(g["counts"][c], g["n"]) for g in areas], orientation="h", name=CAT_LABEL[c],
                    marker_color=CAT_HEX[c], customdata=[g["counts"][c] for g in areas],
                    hovertemplate="%{y} · " + CAT_LABEL[c] + ": %{customdata} (%{x:.1f}%)<extra></extra>")
    for g, nm in zip(areas, names):
        txt = f"<b>{pct(g['avg'])}</b>"
        if g["k"] in pmap:
            txt += f"  {pp(g['avg'] - pmap[g['k']]['avg'])}"
        if g["low"]:
            txt += f"  · {g['low']} bajo 80%"
        fig.add_annotation(x=1.01, xref="paper", y=nm, text=txt, showarrow=False, xanchor="left", font=dict(size=12))
    fig = fig_layout(fig, 90 + 34 * len(areas))
    fig.update_layout(barmode="stack", xaxis=dict(range=[0, 100], ticksuffix="%"), yaxis=dict(type="category"),
                      margin=dict(l=10, r=240, t=40, b=10))
    return fig


def fig_strip(t):
    c = cmp["byType"].get(t.code) if cmp else None
    rows = sorted(t.rows, key=lambda r: r["final"])
    ys, bins = [], {}
    for r in rows:
        k = round(r["final"] / 0.8)
        n = bins.get(k, 0)
        bins[k] = n + 1
        ys.append(0 if n == 0 else ((n + 1) // 2) * 0.12 * (1 if n % 2 else -1))
    lo = min(60, int(min(t.min, c["prev"].min if c and c["prev"] else 100) // 10 * 10))
    fig = go.Figure()
    for x0, x1, cat in [(lo, 60, "MALO"), (60, 70, "REGULAR"), (70, 80, "BUENO"), (80, 98, "MUY BUENO"),
                        (98, 100, "EXCELENTE")]:
        if x1 > lo:
            fig.add_vrect(x0=max(x0, lo), x1=x1, fillcolor=CAT_HEX[cat], opacity=0.08, line_width=0)
    fig.add_scatter(x=[r["final"] for r in rows], y=[max(-1, min(1, y)) for y in ys], mode="markers",
                    marker=dict(size=10, color=[CAT_HEX[r["cat"]] for r in rows], line=dict(width=1, color="white")),
                    text=[f"{r['name']}<br>{pct(r['final'])} · {CAT_LABEL[r['cat']]}" for r in rows],
                    hovertemplate="%{text}<extra></extra>", showlegend=False)
    fig.add_vline(x=t.avg, line=dict(color="#555", width=2))
    if c and c["prev"]:
        fig.add_vline(x=c["prev"].avg, line=dict(color="#999", width=1.5, dash="dot"))
    fig.update_xaxes(range=[lo - 0.5, 100.8], ticksuffix="%", tickvals=[v for v in [lo, 60, 70, 80, 90, 100] if v >= lo])
    fig.update_yaxes(visible=False, range=[-1.3, 1.3])
    return fig_layout(fig, 120)


def fig_criteria(t, r):
    cr = crit_rows(t, r)[::-1]
    labels = ["<br>".join(re.findall(r".{1,48}(?:\s|$)", c["h"])[:2]).strip() for c in cr]
    fig = go.Figure()
    fig.add_bar(y=labels, x=[c["v"] for c in cr], orientation="h", name="Proveedor",
                marker_color=[CAT_HEX[_cat(c["v"])] for c in cr], text=[pct(c["v"]) for c in cr],
                textposition="outside", hovertemplate="%{y}: %{x:.2f}%<extra></extra>")
    if t.n > 1:
        fig.add_scatter(y=labels, x=[c["avg"] for c in cr], mode="markers", name="Promedio del tipo",
                        marker=dict(symbol="line-ns", size=18, line=dict(width=3, color="#333")),
                        hovertemplate="Promedio del tipo: %{x:.2f}%<extra></extra>")
    fig.update_xaxes(range=[0, 115], ticksuffix="%", tickvals=[0, 20, 40, 60, 80, 100])
    return fig_layout(fig, 60 + 46 * len(cr))


def _cat(v):
    from core.analisis import cat_of
    return cat_of(v)


def fig_trend(hist_rows):
    fig = go.Figure()
    fig.add_hrect(y0=98, y1=100.5, fillcolor=CAT_HEX["EXCELENTE"], opacity=0.07, line_width=0)
    fig.add_hrect(y0=80, y1=98, fillcolor=CAT_HEX["MUY BUENO"], opacity=0.07, line_width=0)
    fig.add_scatter(x=[h["label"] for h in hist_rows], y=[h["final"] for h in hist_rows], mode="lines+markers+text",
                    text=[pct(h["final"]) for h in hist_rows], textposition="top center",
                    marker=dict(size=11, color=[CAT_HEX[h["cat"]] for h in hist_rows]), line=dict(color="#666"),
                    showlegend=False, hovertemplate="%{x}: %{y:.2f}%<extra></extra>")
    lo = min(60, min(h["final"] for h in hist_rows) // 10 * 10)
    fig.update_yaxes(range=[lo, 104], ticksuffix="%")
    fig.update_xaxes(type="category")
    return fig_layout(fig, 220)


def _bg(v):
    if v is None or pd.isna(v):
        return ""
    return f"background-color: {CAT_HEX[cat_of(v)]}33"


def period_table(rows_dict: dict, labels: list, base: str | None, first_col: str):
    """rows_dict: {nombre: {periodo: valor}} -> Styler con columnas por periodo y cambio frente al base."""
    df = pd.DataFrame({k: {l: v.get(l) for l in labels} for k, v in rows_dict.items()}).T[labels]
    df = df.apply(pd.to_numeric, errors="coerce")
    cur_l = labels[-1] if a["label"] not in labels else a["label"]
    if base and base in df.columns and cur_l in df.columns:
        df[f"Cambio vs {base} (pp)"] = df[cur_l] - df[base]
    df.index.name = first_col
    return df.style.map(_bg, subset=labels).format("{:.2f}", na_rep="—")


# ---------------------------------------------------------------- Consolidado
with tabs["Consolidado"]:
    low_n = sum(r["final"] < 80 for r, _ in a["all"])
    good = share(a["counts"]["EXCELENTE"] + a["counts"]["MUY BUENO"], a["total"])
    cols = st.columns(5)
    cols[0].metric("Evaluaciones", a["total"], f"{a['total'] - pa['total']:+d} vs {pa['label']}" if pa else None,
                   delta_color="off")
    cols[1].metric("Proveedores distintos", a["uniq"])
    cols[2].metric("Promedio general", pct(a["avg"]), ppm(a["avg"] - pa["avg"]) if pa else None)
    if pa:
        pgood = share(pa["counts"]["EXCELENTE"] + pa["counts"]["MUY BUENO"], pa["total"])
        cols[3].metric("Excelente o muy bueno", fnum(good) + "%", ppm(good - pgood))
        cols[4].metric("Por debajo de 80%", low_n, f"{low_n - sum(r['final'] < 80 for r, _ in pa['all']):+d}",
                       delta_color="inverse")
    else:
        cols[3].metric("Excelente o muy bueno", fnum(good) + "%")
        cols[4].metric("Por debajo de 80%", low_n)

    st.subheader("Distribución por categoría")
    st.plotly_chart(fig_distribution(), width="stretch", config={"displayModeBar": False})
    st.subheader("Calificación por área")
    st.caption("Ordenado por promedio. La barra muestra la proporción de evaluaciones en cada categoría; a la derecha, "
               "el promedio del área" + (f", el cambio frente a {pa['label']}" if pa else "")
               + " y cuántos proveedores quedaron bajo 80%.")
    st.plotly_chart(fig_areas(), width="stretch", config={"displayModeBar": False})
    if multi:
        mlabels = [l for l, _ in multi]
        st.subheader("Promedio por área en los periodos comparados")
        am = area_matrix(multi)
        st.dataframe(period_table({ar: {l: v[l][0] for l in v} for ar, v in am.items()}, mlabels,
                                  pa["label"] if pa else None, "Área"), width="stretch")
        summ = period_summary(multi)
        fig = go.Figure(go.Scatter(x=mlabels, y=[x["Promedio"] for x in summ], mode="lines+markers+text",
                                   text=[pct(x["Promedio"]) for x in summ], textposition="top center",
                                   marker=dict(size=11, color=[CAT_HEX[cat_of(x["Promedio"])] for x in summ]),
                                   line=dict(color="#1F4F7A", width=2.5),
                                   hovertemplate="%{x}: %{y:.2f}%<extra></extra>"))
        fig.update_yaxes(range=[max(0, min(x["Promedio"] for x in summ) - 5), 101], ticksuffix="%")
        fig.update_xaxes(type="category")
        st.markdown("**Promedio general por periodo**")
        st.plotly_chart(fig_layout(fig, 260), width="stretch", config={"displayModeBar": False}, key="multi_avg")
    st.subheader("Análisis")
    for ptxt in consolidated_paras(a, cmp):
        st.write(ptxt)
    if multi:
        for ptxt in insights(multi):
            st.write(ptxt)
    with st.expander(f"Resumen general por tipo de evaluación ({len(a['types'])} tipos)"):
        df = pd.DataFrame([[t.title, *[t.counts[c] for c in CATS], t.n] for t in a["types"]],
                          columns=["Tipo de evaluación", *CATS, "TOTAL"])
        df.loc[len(df)] = ["TOTAL", *[a["counts"][c] for c in CATS], a["total"]]
        st.dataframe(df, hide_index=True, width="stretch", height=min(38 * (len(df) + 1), 900))

# ---------------------------------------------------------------- Por área y tipo
with tabs["Por área y tipo"]:
    area_sel = st.selectbox("Área", ["Todas"] + [g["name"] for g in a["procs"]], key="area_sel")
    st.caption("Cada punto es un proveedor (pasa el cursor para ver el nombre). Línea sólida: promedio del tipo"
               + (f"; línea punteada: promedio en {cmp['prevLabel']}." if cmp else "."))
    for g in a["procs"]:
        if area_sel != "Todas" and g["name"] != area_sel:
            continue
        st.subheader(f"Proveedores {g['name'].lower()}")
        st.caption(f"{g['n']} evaluaciones · promedio {pct(g['avg'])} · {counts_text(g['counts'])}")
        for t in g["types"]:
            c = cmp["byType"].get(t.code) if cmp else None
            with st.container(border=True):
                h1, h2 = st.columns([4, 1])
                h1.markdown(f"**{t.type_name}**  \n<span class='small'>{t.code} · {t.n} "
                            f"{'proveedor' if t.n == 1 else 'proveedores'}</span>", unsafe_allow_html=True)
                h2.metric("Promedio", pct(t.avg), ppm(t.avg - c["prev"].avg) if c and c["prev"] else None,
                          label_visibility="collapsed")
                st.plotly_chart(fig_strip(t), width="stretch", config={"displayModeBar": False},
                                key=f"strip-{t.code}")
                st.markdown(" ".join(f"{pill(k)} {t.counts[k]}" for k in CATS if t.counts[k]),
                            unsafe_allow_html=True)
                if multi:
                    per = []
                    for lbl, x in multi:
                        tt = next((q for q in x["types"] if q.code == t.code), None)
                        per.append(f"{lbl}: **{pct(tt.avg)}**" if tt else f"{lbl}: sin datos")
                    st.markdown("Promedio por periodo · " + " · ".join(per))
                for ptxt in type_paras(t, cmp):
                    st.write(ptxt)
                with st.expander(f"Ver {t.n} {'proveedor' if t.n == 1 else 'proveedores'}"):
                    pm = {id(m["r"]): m for m in c["matched"]} if c else {}
                    data = []
                    for r in t.rows:
                        row = {"Proveedor": r["name"], "NIT": r["nit"], "Final": round(r["final"], 2),
                               "Categoría": CAT_LABEL[r["cat"]]}
                        if multi:
                            for lbl, x in multi:
                                if lbl == a["label"]:
                                    continue
                                tt = next((q for q in x["types"] if q.code == t.code), None)
                                pr = next((q for q in tt.rows if pkey(q) == pkey(r)), None) if tt else None
                                row[f"Final {lbl}"] = round(pr["final"], 2) if pr else None
                        if cmp:
                            m = pm.get(id(r))
                            if not multi:
                                row[f"Final {cmp['prevLabel']}"] = round(m["pr"]["final"], 2) if m else None
                            row[f"Cambio vs {cmp['prevLabel']} (pp)"] = round(m["d"], 2) if m else None
                        data.append(row)
                    st.dataframe(pd.DataFrame(data), hide_index=True, width="stretch")

# ---------------------------------------------------------------- Ficha por proveedor
idx = provider_index(a)
with tabs["Ficha por proveedor"]:
    only_low = st.toggle("Mostrar solo proveedores por debajo de 80%", value=False, key="only_low")
    plist = sorted(idx.values(), key=lambda p: p["name"])
    if only_low:
        plist = [p for p in plist if any(r["final"] < 80 for _, r in p["evals"])]
    if not plist:
        st.info("No hay proveedores con ese filtro.")
    else:
        pk = st.selectbox("Proveedor (escribe para buscar por nombre o NIT)", [p["key"] for p in plist],
                          format_func=lambda k: f"{idx[k]['name']} · {idx[k]['nit']}", key="prov_sel")
        p = idx[pk]
        hcol, bcol = st.columns([4, 1])
        hcol.markdown(f"### {p['name']}\nNIT {p['nit']} · {len(p['evals'])} "
                      f"{'evaluación' if len(p['evals']) == 1 else 'evaluaciones'} en {a['label']}")
        with bcol:
            if st.button("Preparar ficha (.docx)", key=f"ficha1-{pk}"):
                st.session_state["ficha1"] = (pk, build_fichas(a, cmp, [p], timeline))
            if st.session_state.get("ficha1", (None,))[0] == pk:
                st.download_button("Descargar ficha", st.session_state["ficha1"][1],
                                   file_name=f"Ficha-{slug(p['name'])[:40]}-{slug(a['label'])}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        for t, r in p["evals"]:
            others = [(ot, orow) for ot, orow in p["evals"] if ot is not t]
            hist_rows = history_for(timeline, p["key"], t.code)
            with st.container(border=True):
                st.markdown(f"<span class='small'>{PROC[t.proc]} · {t.code}</span>", unsafe_allow_html=True)
                st.markdown(f"#### {t.type_name}")
                m1, m2, m3 = st.columns(3)
                prev_row = hist_rows[-2] if len(hist_rows) >= 2 else None
                m1.metric("Calificación final", pct(r["final"]),
                          ppm(r["final"] - prev_row["final"]) + f" vs {prev_row['label']}" if prev_row else None)
                m2.markdown(f"Categoría<br>{pill(r['cat'])}", unsafe_allow_html=True)
                m3.metric("Puesto en su tipo", f"{rank_in(t, r)} de {t.n}" if t.n > 1 else "Único proveedor")
                if t.n > 1:
                    m3.caption(f"Promedio del tipo: {pct(t.avg)}")
                left, right = st.columns([3, 2])
                with left:
                    if crit_rows(t, r):
                        st.plotly_chart(fig_criteria(t, r), width="stretch",
                                        config={"displayModeBar": False}, key=f"crit-{pk}-{t.code}")
                with right:
                    st.markdown("**Análisis**")
                    st.markdown("\n".join(f"- {x}" for x in provider_insights(t, r, others, hist_rows, cmp)))
                    if len(hist_rows) >= 2:
                        st.plotly_chart(fig_trend(hist_rows), width="stretch",
                                        config={"displayModeBar": False}, key=f"trend-{pk}-{t.code}")

# ---------------------------------------------------------------- Comparativo
if cmp:
    with tabs["Comparativo"]:
        base_opts = list(pas_all)
        if len(base_opts) > 1:
            base_lbl = st.radio("Comparar los movimientos de proveedores contra", base_opts,
                                index=base_opts.index(pa["label"]), horizontal=True, key="cmp_base")
            cmp_view = compare(a, pas_all[base_lbl]) if base_lbl != pa["label"] else cmp
        else:
            cmp_view = cmp
        cmp, _cmp_main = cmp_view, cmp
        pa_view = cmp["prev"]
        cols = st.columns(4)
        cols[0].metric("En ambos periodos", cmp["matched"])
        cols[1].metric("Mejoraron", cmp["improved"])
        cols[1].caption(f"{cmp['catUp']} subieron de categoría")
        cols[2].metric("Bajaron", cmp["declined"])
        cols[2].caption(f"{cmp['catDown']} bajaron de categoría")
        cols[3].metric("Proveedores nuevos", cmp["nNew"])
        st.subheader("Promedio por proceso")
        pmap = {g["k"]: g for g in pa_view["procs"]}
        rows = [g for g in a["procs"] if g["k"] in pmap]
        fig = go.Figure()
        for g in rows:
            pv, d = pmap[g["k"]]["avg"], g["avg"] - pmap[g["k"]]["avg"]
            col = "#11774A" if d > 0.05 else "#B3263E" if d < -0.05 else "#888"
            fig.add_scatter(x=[pv, g["avg"]], y=[g["name"]] * 2, mode="lines", line=dict(color=col, width=3),
                            showlegend=False, hoverinfo="skip")
        fig.add_scatter(x=[pmap[g["k"]]["avg"] for g in rows], y=[g["name"] for g in rows], mode="markers",
                        name=cmp["prevLabel"], marker=dict(size=11, color="white", line=dict(color="#888", width=2)),
                        hovertemplate="%{y} · " + cmp["prevLabel"] + ": %{x:.2f}%<extra></extra>")
        for g in rows:
            d = g["avg"] - pmap[g["k"]]["avg"]
            fig.add_annotation(x=1.01, xref="paper", y=g["name"], text=f"<b>{pct(g['avg'])}</b>  {pp(d)}",
                               showarrow=False, xanchor="left",
                               font=dict(size=12, color="#11774A" if d > 0.05 else "#B3263E" if d < -0.05 else "#777"))
        fig.add_scatter(x=[g["avg"] for g in rows], y=[g["name"] for g in rows], mode="markers", name=a["label"],
                        marker=dict(size=12, color="#1F4F7A"),
                        hovertemplate="%{y} · " + a["label"] + ": %{x:.2f}%<extra></extra>")
        fig.update_xaxes(ticksuffix="%", range=[min(min(g["avg"], pmap[g["k"]]["avg"]) for g in rows) - 3, 104])
        fig.update_yaxes(autorange="reversed")
        fig = fig_layout(fig, 90 + 36 * len(rows))
        fig.update_layout(margin=dict(l=10, r=170, t=40, b=10))
        if multi:
            # todos los periodos comparados: un punto por periodo, del más claro (antiguo) al más oscuro (actual)
            shades = ["#C9D3DC", "#9FB2C4", "#7591AC", "#4B7094", "#1F4F7A"]
            fig = go.Figure()
            names = [g["name"] for g in a["procs"]]
            for g in a["procs"]:
                vals = [g2["avg"] for _, x2 in multi for g2 in x2["procs"] if g2["k"] == g["k"]]
                fig.add_scatter(x=[min(vals), max(vals)], y=[g["name"]] * 2, mode="lines",
                                line=dict(color="#D5DAD7", width=6), showlegend=False, hoverinfo="skip")
            n = len(multi)
            for i, (lbl, x2) in enumerate(multi):
                gm = {g2["k"]: g2 for g2 in x2["procs"]}
                ys = [g["name"] for g in a["procs"] if g["k"] in gm]
                xs = [gm[g["k"]]["avg"] for g in a["procs"] if g["k"] in gm]
                col = shades[-1] if lbl == a["label"] else shades[max(0, len(shades) - 1 - (n - 1 - i))]
                fig.add_scatter(x=xs, y=ys, mode="markers", name=lbl,
                                marker=dict(size=14 if lbl == a["label"] else 11, color=col,
                                            line=dict(color="white", width=1)),
                                hovertemplate="%{y} · " + lbl + ": %{x:.2f}%<extra></extra>")
            allv = [g2["avg"] for _, x2 in multi for g2 in x2["procs"]]
            fig.update_xaxes(ticksuffix="%", range=[min(allv) - 3, 101.5])
            fig.update_yaxes(autorange="reversed")
            fig = fig_layout(fig, 90 + 36 * len(names))
            fig.update_layout(margin=dict(l=10, r=10, t=40, b=10))
            st.caption("Cada punto es el promedio del proceso en un periodo; el más oscuro es el periodo actual.")
        st.plotly_chart(fig, width="stretch",
                        config={"displayModeBar": False})

        def movers(ms):
            return pd.DataFrame([{"Proveedor": m["r"]["name"], "Tipo": m["t"].type_name,
                                  cmp["prevLabel"]: round(m["pr"]["final"], 2), a["label"]: round(m["r"]["final"], 2),
                                  "Cambio (pp)": round(m["d"], 2)} for m in ms])
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Mayores caídas")
            downs = [m for m in cmp["movers"] if m["d"] < -0.05][:15]
            if downs:
                st.dataframe(movers(downs), hide_index=True, width="stretch")
            else:
                st.caption("Ninguna.")
        with c2:
            st.subheader("Mayores mejoras")
            ups = [m for m in reversed(cmp["movers"]) if m["d"] > 0.05][:15]
            if ups:
                st.dataframe(movers(ups), hide_index=True, width="stretch")
            else:
                st.caption("Ninguna.")
        if cmp["goneTypes"]:
            st.caption(f"Tipos evaluados en {cmp['prevLabel']} sin registros en {a['label']}: "
                       + ", ".join(t.type_name for t in cmp["goneTypes"]) + ".")
        if multi:
            st.subheader("Proveedores evaluados en todos los periodos comparados")
            mlabels = [l for l, _ in multi]
            pmx = provider_matrix(multi)
            todos = {f"{e['nombre']} · {e['nit']}": e["valores"] for e in pmx.values()
                     if all(l in e["valores"] for l in mlabels)}
            st.caption(f"{len(todos)} proveedores. Si un proveedor tiene varias evaluaciones en un periodo se muestra "
                       "su promedio.")
            if todos:
                st.dataframe(period_table(todos, mlabels, pa_view["label"], "Proveedor"), width="stretch", height=380)
        cmp = _cmp_main

# ---------------------------------------------------------------- Histórico

with tabs["Histórico"]:
    all_labels = [l for l, _ in timeline]
    if len(all_labels) < 2:
        st.info("Para comparar varios años o semestres, guarda al menos dos periodos en el historial. En la barra "
                "lateral usa **Cargar varios periodos al historial** para subir de una vez los archivos de años "
                "anteriores (por ejemplo 2022, 2023 y 2024).")
    else:
        sel = st.multiselect("Periodos a comparar", all_labels, default=all_labels, key="hist_sel")
        tl = [(l, x) for l, x in timeline if l in sel]
        if len(tl) < 2:
            st.warning("Elige al menos dos periodos.")
        else:
            labels = [l for l, _ in tl]
            summ = period_summary(tl)
            cols = st.columns(4)
            cols[0].metric("Promedio general", pct(summ[-1]["Promedio"]),
                           ppm(summ[-1]["Promedio"] - summ[0]["Promedio"]) + f" vs {labels[0]}")
            cols[1].metric("Excelente o muy bueno", fnum(summ[-1]["% excelente o muy bueno"]) + "%",
                           ppm(summ[-1]["% excelente o muy bueno"] - summ[0]["% excelente o muy bueno"]) + f" vs {labels[0]}")
            cols[2].metric("Por debajo de 80%", summ[-1]["Bajo 80%"],
                           f"{summ[-1]['Bajo 80%'] - summ[0]['Bajo 80%']:+d} vs {labels[0]}", delta_color="inverse")
            cols[3].metric("Evaluaciones", summ[-1]["Evaluaciones"],
                           f"{summ[-1]['Evaluaciones'] - summ[0]['Evaluaciones']:+d} vs {labels[0]}", delta_color="off")

            st.subheader("Análisis")
            for t_ in insights(tl):
                st.write(t_)

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Promedio general por periodo**")
                fig = go.Figure(go.Scatter(x=labels, y=[x["Promedio"] for x in summ], mode="lines+markers+text",
                                           text=[pct(x["Promedio"]) for x in summ], textposition="top center",
                                           marker=dict(size=11, color=[CAT_HEX[cat_of(x["Promedio"])] for x in summ]),
                                           line=dict(color="#1F4F7A", width=2.5),
                                           hovertemplate="%{x}: %{y:.2f}%<extra></extra>"))
                lo = min(x["Promedio"] for x in summ)
                fig.update_yaxes(range=[max(0, lo - 5), 101], ticksuffix="%")
                fig.update_xaxes(type="category")
                st.plotly_chart(fig_layout(fig, 300), width="stretch", config={"displayModeBar": False})
            with c2:
                st.markdown("**Distribución por categoría**")
                fig = go.Figure()
                for c in CATS:
                    fig.add_bar(x=labels, y=[share(x[c], x["Evaluaciones"]) for x in summ], name=CAT_LABEL[c],
                                marker_color=CAT_HEX[c], customdata=[x[c] for x in summ],
                                hovertemplate="%{x} · " + CAT_LABEL[c] + ": %{customdata} (%{y:.1f}%)<extra></extra>")
                fig.update_layout(barmode="stack", yaxis=dict(range=[0, 100], ticksuffix="%"),
                                  xaxis=dict(type="category"))
                fig = fig_layout(fig, 300)
                fig.update_layout(margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

            st.dataframe(pd.DataFrame(summ).rename(columns={c: CAT_LABEL[c] for c in CATS}).style.format(
                {"Promedio": "{:.2f}%", "% excelente o muy bueno": "{:.1f}%"}), hide_index=True, width="stretch")

            st.subheader("Promedio por área")
            am = area_matrix(tl)
            dfa = pd.DataFrame({area: {l: v[l][0] if l in v else None for l in labels} for area, v in am.items()}).T
            dfa = dfa[labels].apply(pd.to_numeric, errors="coerce")
            dfa[f"Cambio {labels[0]}→{labels[-1]} (pp)"] = dfa[labels[-1]] - dfa[labels[0]]
            st.dataframe(dfa.style.map(_bg, subset=labels).format("{:.2f}", na_rep="—"), width="stretch")
            ch = dfa[f"Cambio {labels[0]}→{labels[-1]} (pp)"].abs().sort_values(ascending=False)
            areas_sel = st.multiselect("Ver la evolución de estas áreas", list(dfa.index),
                                       default=list(ch.dropna().index[:4]), key="hist_areas")
            if areas_sel:
                fig = go.Figure()
                for ar in areas_sel:
                    fig.add_scatter(x=labels, y=[dfa.loc[ar, l] for l in labels], mode="lines+markers", name=ar,
                                    hovertemplate=ar + " · %{x}: %{y:.2f}%<extra></extra>")
                fig.update_yaxes(ticksuffix="%")
                fig.update_xaxes(type="category")
                fig = fig_layout(fig, 320)
                fig.update_layout(margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

            st.subheader("Promedio por tipo de evaluación")
            tm = type_matrix(tl)
            dft = pd.DataFrame([{"Proceso": e["proceso"], "Tipo": e["tipo"],
                                 **{l: e["valores"][l][0] if l in e["valores"] else None for l in labels}}
                                for e in tm.values()])
            dft[labels] = dft[labels].apply(pd.to_numeric, errors="coerce")
            dft[f"Cambio {labels[0]}→{labels[-1]} (pp)"] = dft[labels[-1]] - dft[labels[0]]
            dft = dft.sort_values(["Proceso", "Tipo"])
            st.dataframe(dft.style.map(_bg, subset=labels).format({c: "{:.2f}" for c in dft.columns[2:]}, na_rep="—"),
                         hide_index=True, width="stretch", height=420)

            st.subheader("Proveedores")
            rl = recurrent_low(tl)
            if rl:
                st.markdown(f"**Por debajo de 80% en dos o más periodos ({len(rl)})**")
                dfr = pd.DataFrame([{"Proveedor": x["nombre"], "NIT": x["nit"], "Periodos bajo 80%": len(x["bajos"]),
                                     **{l: x["bajos"].get(l) for l in labels}} for x in rl])
                dfr[labels] = dfr[labels].apply(pd.to_numeric, errors="coerce")
                st.caption("Se muestra la calificación más baja del proveedor en cada periodo en que quedó bajo 80%.")
                st.dataframe(dfr
                             .style.map(_bg, subset=labels).format({l: "{:.2f}" for l in labels}, na_rep="—"),
                             hide_index=True, width="stretch")
            else:
                st.caption("Ningún proveedor quedó por debajo de 80% en más de un periodo.")
            pm = provider_matrix(tl)
            q = st.text_input("Buscar proveedor por nombre o NIT", key="hist_q").strip().upper()
            rows_p = [{"Proveedor": e["nombre"], "NIT": e["nit"], **{l: e["valores"].get(l) for l in labels},
                       "Periodos evaluado": len(e["valores"])}
                      for e in pm.values() if not q or q in e["nombre"].upper() or q in str(e["nit"]).upper()]
            dfp = pd.DataFrame(rows_p).sort_values("Proveedor") if rows_p else pd.DataFrame()
            if not dfp.empty:
                dfp[labels] = dfp[labels].apply(pd.to_numeric, errors="coerce")
            if not dfp.empty:
                st.dataframe(dfp.style.map(_bg, subset=labels).format({l: "{:.2f}" for l in labels}, na_rep="—"),
                             hide_index=True, width="stretch", height=380)
                st.caption("Si un proveedor tiene varias evaluaciones en un periodo, se muestra su promedio. "
                           "El detalle por tipo está en la pestaña Ficha por proveedor.")
            st.download_button("Descargar histórico (.xlsx)", build_hist_xlsx(tl),
                               file_name=f"Historico-proveedores-{slug(labels[0])}-{slug(labels[-1])}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary")

# ---------------------------------------------------------------- Cálculo del Final
with tabs["Cálculo del Final"]:
    rc = recalc_rows(a)
    st.markdown(
        "La plataforma calcula el **Final como promedio ponderado** de los criterios y toma los **criterios en blanco "
        "como 0**. La app deduce el peso de cada criterio a partir de los mismos datos (en cada tipo, el Final de la "
        "plataforma = suma de peso × puntaje) y, en las filas con criterios en blanco, recalcula el Final sin "
        "tenerlos en cuenta: **Final = Final plataforma ÷ (1 − peso de los criterios en blanco)**. "
        "Las filas completas conservan el Final de la plataforma.")
    if not ignore_blanks:
        st.info("Ahora estás usando el Final de la plataforma (criterios en blanco = 0). Cambia la opción "
                "**Criterios en blanco** en la barra lateral para recalcular.")
    else:
        plat_avg = sum(r["finalPlat"] for r, _ in a["all"]) / a["total"]
        c1, c2, c3 = st.columns(3)
        c1.metric("Evaluaciones recalculadas", len(rc))
        c2.metric("Cambiaron de categoría", sum(r["cat"] != cat_of(r["finalPlat"]) for r, _ in rc))
        c3.metric("Promedio general", pct(a["avg"]), ppm(a["avg"] - plat_avg) + " vs plataforma")
        if rc:
            st.dataframe(pd.DataFrame([{
                "Proveedor": r["name"], "Tipo": t.type_name,
                "Final plataforma": round(r["finalPlat"], 2), "Final recalculado": round(r["final"], 2),
                "Cambio (pp)": round(r["final"] - r["finalPlat"], 2),
                "Categoría": f"{CAT_LABEL[cat_of(r['finalPlat'])]} → {CAT_LABEL[r['cat']]}",
                "Criterios en blanco": "; ".join(r["blanks"]),
                "Peso": "estimado (igual para todos)" if r["estimated"] else "deducido de la plataforma"}
                for r, t in sorted(rc, key=lambda x: (x[1].code, -x[0]["final"] + x[0]["finalPlat"]))]),
                hide_index=True, width="stretch")
        else:
            st.success("Ninguna evaluación de este periodo tiene criterios en blanco.")

    st.subheader("Ponderación por tipo de evaluación")
    st.caption("Pesos deducidos de los datos del periodo y del historial. «Exacto» reproduce el Final de la "
               "plataforma en todas las filas. Si un peso no se puede deducir (pocas filas o criterios que siempre "
               "tienen el mismo puntaje), se supone igual para todos los criterios; puedes escribir el peso real y "
               "guardarlo.")
    con_blancos = {t.code for _, t in rc}
    tipos_opts = {f"{t.type_name} ({t.code})" + (" · tiene criterios en blanco" if t.code in con_blancos else ""): t
                  for t in sorted(a["types"], key=lambda t: t.code not in con_blancos)}
    tsel = tipos_opts[st.selectbox("Tipo de evaluación", list(tipos_opts), key="pond_tipo")]
    sheet = next(sh for sh in cur["sheets"] if sh["code"] == tsel.code)
    info = wmap.get((tsel.code, tuple(sheet["criteria"])))
    ov = overrides.get(tsel.code) or {}
    rows_w = []
    for j, crit in enumerate(sheet["criteria"]):
        if info and info["exact"] and info["ident"][j]:
            estado, ded = "exacto", round(info["weights"][j] * 100, 2)
        elif info and info["exact"]:
            estado, ded = "no identificable", None
        else:
            estado, ded = "no se pudo deducir", None
        blank_n = sum(1 for r in sheet["rows"] if j < len(r["scores"]) and r["scores"][j] is None)
        rows_w.append({"Criterio": crit, "Peso deducido (%)": ded, "Estado": estado,
                       "Filas en blanco": blank_n, "Peso a usar (%)": ov.get(crit, ded)})
    ed = st.data_editor(pd.DataFrame(rows_w), hide_index=True, width="stretch", key=f"pond_{tsel.code}",
                        disabled=["Criterio", "Peso deducido (%)", "Estado", "Filas en blanco"],
                        column_config={"Peso a usar (%)": st.column_config.NumberColumn(min_value=0, max_value=100,
                                                                                        step=0.5, format="%.2f")})
    b1, b2, _ = st.columns([1, 1, 3])
    if b1.button("Guardar ponderación", key="pond_save"):
        vals = ed["Peso a usar (%)"].tolist()
        if any(v is None or pd.isna(v) for v in vals):
            st.warning("Escribe un peso para todos los criterios antes de guardar.")
        else:
            overrides[tsel.code] = {c: float(v) for c, v in zip(ed["Criterio"], vals)}
            POND.write_text(json.dumps(overrides, ensure_ascii=False, indent=2), encoding="utf-8")
            st.rerun()
    if tsel.code in overrides and b2.button("Volver a los pesos deducidos", key="pond_reset"):
        overrides.pop(tsel.code)
        POND.write_text(json.dumps(overrides, ensure_ascii=False, indent=2), encoding="utf-8")
        st.rerun()
    if tsel.code in overrides:
        st.caption("Este tipo usa la ponderación que guardaste (archivo config/ponderaciones.json).")

# ---------------------------------------------------------------- Revisión de datos
with tabs["Revisión de datos"]:
    st.caption("La categoría se recalcula desde la columna Final. Estos registros no coinciden con lo que trae el "
               "archivo de la plataforma.")
    if not a["alerts"]:
        st.success("Todas las categorías del archivo coinciden con la columna Final.")
    for x in a["alerts"]:
        r, t = x["r"], x["t"]
        if x["kind"] == "nofinal":
            st.warning(f"**{r['name']}** ({t.code}): no tiene calificación Final y no se incluyó.")
        elif x["kind"] == "sin":
            fp = r.get("finalPlat", r["final"])
            st.warning(f"**{r['name']}** ({t.code}): la plataforma no asignó categoría; con {pct(fp)} "
                       f"corresponde a **{CAT_LABEL[cat_of(fp)]}**.")
        else:
            fp = r.get("finalPlat", r["final"])
            st.warning(f"**{r['name']}** ({t.code}): la plataforma dice {r['platCat']}, pero con {pct(fp)} "
                       f"corresponde a **{CAT_LABEL[cat_of(fp)]}**.")
    odd = [(t, r) for t in a["types"] for r in t.rows
           if crit_rows(t, r) and r["final"] < min(c["v"] for c in crit_rows(t, r)) - 0.5]
    if odd:
        st.markdown("**Calificación final menor que todos sus criterios** (revisar criterios o ponderación en la "
                    "plataforma):")
        for t, r in odd:
            st.markdown(f"- {r['name']} ({t.code}): final {pct(r['final'])}, criterio más bajo "
                        f"{pct(min(c['v'] for c in crit_rows(t, r)))}")
    zeros = [(t, c) for t in a["types"] for c in t.crit if c["avg"] == 0]
    if zeros:
        st.markdown("**Criterios con 0% para todos los proveedores del tipo** (¿aplica el criterio?):")
        for t, c in zeros:
            st.markdown(f"- {t.type_name} ({t.code}): «{c['h']}»")

# ---------------------------------------------------------------- Descargas
with tabs["Descargas"]:
    sig = (f"{source}|{a['label']}|{cmp['prevLabel'] if cmp else ''}|{','.join(pas_all)}|{ignore_blanks}|"
           + hashlib.md5(json.dumps(overrides, sort_keys=True).encode()).hexdigest()[:8])
    if st.session_state.get("concl_sig") != sig:
        st.session_state["concl_sig"] = sig
        st.session_state["concl"] = conclusion_text(a, cmp)
    st.subheader("Conclusiones para el informe")
    st.text_area("Puedes editar este texto antes de generar el Word.", key="concl", height=140)
    acts = action_rows(a)
    st.caption("Acciones para la mejora sugeridas: " + ("ninguna (NA)." if acts[0][0] == "NA" else
               f"{len(acts)} (AC para calificaciones < 70%, CR de seguimiento para 70–80%); fechas y responsables "
               "quedan en blanco en el Word."))

    fname = slug(a["label"])
    c1, c2, c3 = st.columns(3)
    with c1, st.container(border=True):
        st.markdown("**Resumen general (.xlsx)**")
        st.caption("Hojas General, Por tipo, Detalle" + (" y Comparativo." if cmp else "."))
        st.download_button("Descargar Excel", build_xlsx(a, cmp, multi), file_name=f"ResumenGeneral-{fname}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")
    with c2, st.container(border=True):
        st.markdown("**Informe de gestión (.docx)**")
        st.caption("Gráfica por tipo, consolidados, calificación por área, resumen general, conclusiones y acciones.")
        n_rc = len(recalc_rows(a))
        nota = (f"Nota: {n_rc} evaluaciones tenían criterios sin diligenciar, que la plataforma toma como 0. Para ellas "
                f"la calificación final se recalculó sin tener en cuenta esos criterios, conservando la ponderación de "
                f"los demás." if n_rc else "")
        meta = {"fecha": fecha.strftime("%d-%m-%Y"), "autor": autor, "conclusiones": st.session_state["concl"],
                "nota_final": nota, "multi": multi}
        dsig = sig + json.dumps({k: v for k, v in meta.items() if k != "multi"}, ensure_ascii=False)
        if st.button("Generar informe", width="stretch", key="gen_docx"):
            bar = st.progress(0.0, text="Generando gráficas…")
            st.session_state["docx"] = (dsig, build_docx(a, cmp, meta, progress=lambda v: bar.progress(min(v, 1.0))))
            bar.empty()
        if st.session_state.get("docx", (None,))[0] == dsig:
            st.download_button("Descargar informe", st.session_state["docx"][1],
                               file_name=f"Informe-gestion-proveedores-{fname}.docx", type="primary",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               width="stretch")
    with c3, st.container(border=True):
        st.markdown("**Fichas por proveedor (.docx)**")
        low_list = [p for p in idx.values() if any(r["final"] < 80 for _, r in p["evals"])]
        fopts = {f"Solo bajo 80% ({len(low_list)})": ("bajo-80", low_list, None),
                 f"Todos ({len(idx)})": ("todos", list(idx.values()), None)}
        for g in a["procs"]:
            ps = [p for p in idx.values() if any(t.proc == g["k"] for t, _ in p["evals"])]
            fopts[f"Área: {g['name']} ({len(ps)})"] = (slug(g["name"]).lower(), ps, g["k"])
        fsel = st.selectbox("Proveedores", list(fopts), key="fichas_sel")
        fslug, flist, fproc = fopts[fsel]
        fsig = sig + fsel
        if st.button("Generar fichas", width="stretch", key="gen_fichas"):
            if not flist:
                st.warning("No hay proveedores con ese filtro.")
            else:
                bar = st.progress(0.0, text=f"Generando {len(flist)} fichas…")
                st.session_state["fichas"] = (fsig, build_fichas(a, cmp, flist, timeline, fproc,
                                                                 progress=lambda v: bar.progress(min(v, 1.0))))
                bar.empty()
        if st.session_state.get("fichas", (None,))[0] == fsig:
            st.download_button("Descargar fichas", st.session_state["fichas"][1],
                               file_name=f"Fichas-proveedores-{fslug}-{fname}.docx", type="primary",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               width="stretch")

    with st.container(border=True):
        st.markdown("**Datos para Power BI (.xlsx)**")
        st.caption("Tablas limpias de todos los periodos disponibles ("
                   + ", ".join(lbl for lbl, _ in timeline)
                   + "), con los mismos nombres del modelo de Power BI: Evaluaciones, Criterios y dimensiones.")
        st.download_button("Descargar datos para Power BI", build_powerbi_xlsx(timeline),
                           file_name="Datos_PowerBI_Proveedores.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
