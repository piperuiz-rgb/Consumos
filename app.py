import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(
    page_title="Planificación de Consumos",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* ── Reference card ── */
.ref-card {
    background: #e8e8e8;
    border-left: 4px solid #888;
    border-radius: 0 6px 6px 0;
    padding: 7px 14px;
    margin: 18px 0 2px 0;
}
.ref-title { font-size: 0.9em; font-weight: 700; color: #2c2c2c; }

/* ── Column/row headers ── */
.col-hdr {
    background: #f0f0f0;
    border: 1px solid #d0d0d0;
    border-radius: 4px;
    text-align: center;
    padding: 3px 6px;
    font-size: 0.78em;
    font-weight: 600;
    color: #555;
}
.col-hdr-tot {
    background: #e0e0e0;
    border: 1px solid #b8b8b8;
    border-radius: 4px;
    text-align: center;
    padding: 3px 6px;
    font-size: 0.78em;
    font-weight: 700;
    color: #333;
}
.sz-lbl {
    text-align: right;
    font-weight: 600;
    font-size: 0.82em;
    color: #444;
    padding: 10px 8px 0 0;
}

/* ── Totals ── */
.tot-val {
    text-align: center;
    font-weight: 700;
    font-size: 0.9em;
    color: #333;
    padding-top: 8px;
}
.tot-lbl {
    text-align: right;
    font-weight: 700;
    font-size: 0.82em;
    color: #333;
    padding: 8px 8px 0 0;
    border-top: 1px solid #ccc;
}
.tot-zero {
    text-align: center;
    color: #ccc;
    font-size: 0.85em;
    padding-top: 8px;
}
.grand-tot {
    text-align: center;
    font-weight: 700;
    font-size: 1em;
    color: #111;
    padding-top: 8px;
    border-top: 1px solid #bbb;
}

/* ── Number inputs: compact & centered ── */
div[data-testid="stNumberInput"] {
    margin-bottom: 0 !important;
}
div[data-testid="stNumberInput"] input {
    text-align: center !important;
    font-weight: 600 !important;
    font-size: 0.9em !important;
}

/* ── No min-width on columns ── */
div[data-testid="stHorizontalBlock"] > div {
    min-width: 0 !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Planificación de Producción — Cálculo de Consumos")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VARIANTES_PATH = os.path.join(BASE_DIR, "Variantes.xlsx")
BOM_PATH = os.path.join(BASE_DIR, "listamateriales.xlsx")


@st.cache_data
def load_data():
    variantes = pd.read_excel(VARIANTES_PATH, sheet_name="Variantes de producto", dtype=str)
    bom = pd.read_excel(BOM_PATH, sheet_name="Fichero ejemplo")

    bom["Cod Barras Variante"] = bom["Cod Barras Variante"].astype(str).str.strip()
    bom["EAN Componente"] = bom["EAN Componente"].astype(str).str.strip()
    variantes["Código de barras principal"] = (
        variantes["Código de barras principal"].astype(str).str.strip()
    )
    variantes["Color"] = variantes["Color"].str.replace("Color: ", "", regex=False).str.strip()
    variantes["Talla"] = variantes["Talla"].str.replace("Talla: ", "", regex=False).str.strip()

    barcode_lookup = variantes.set_index("Código de barras principal").to_dict("index")
    bom_barcodes = set(bom["Cod Barras Variante"].unique())
    finished = variantes[variantes["Código de barras principal"].isin(bom_barcodes)].copy()

    return variantes, bom, barcode_lookup, finished


variantes, bom, barcode_lookup, finished = load_data()

_SZ = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "2XL", "XXXL", "3XL", "4XL"]


def _sz_key(s):
    u = str(s).strip().upper()
    if u in _SZ:
        return (0, _SZ.index(u), "")
    try:
        return (1, int(u), "")
    except ValueError:
        return (2, 0, str(s))


# ── SECTION 1 — Filters ──────────────────────────────────────────────────────
st.header("1. Plan de producción")
st.markdown(
    "Filtra las referencias e introduce las unidades a fabricar. "
    "Usa la **asignación masiva** para rellenar rangos completos de talla o color."
)

col1, col2, col3, col4 = st.columns([3, 2, 2, 2])
with col1:
    q = st.text_input("Buscar por nombre", placeholder="ej. Vestido Goya…")
with col2:
    sel_color = st.selectbox(
        "Color", ["Todos"] + sorted(finished["Color"].dropna().unique().tolist())
    )
with col3:
    sel_talla = st.selectbox(
        "Talla",
        ["Todas"] + sorted(finished["Talla"].dropna().unique().tolist(), key=_sz_key),
    )
with col4:
    sel_ref = st.selectbox(
        "Referencia",
        ["Todas"] + sorted(finished["Referencia interna"].dropna().unique().tolist()),
    )

df = finished.copy()
if q:
    df = df[df["Nombre"].str.contains(q, case=False, na=False)]
if sel_color != "Todos":
    df = df[df["Color"] == sel_color]
if sel_talla != "Todas":
    df = df[df["Talla"] == sel_talla]
if sel_ref != "Todas":
    df = df[df["Referencia interna"] == sel_ref]

st.caption(f"{len(df)} variantes mostradas")


# ── Bulk assignment ───────────────────────────────────────────────────────────
def _bulk_apply(target_df, delta):
    """delta=int → add/subtract; delta=None → reset to 0."""
    for bc in target_df["Código de barras principal"]:
        sk = f"qty_{bc}"
        cur = st.session_state.get(sk, 0)
        st.session_state[sk] = 0 if delta is None else max(0, cur + delta)


with st.expander("Asignación masiva por talla / color", expanded=False):
    mc1, mc2 = st.columns(2)
    with mc1:
        bulk_tallas = st.multiselect(
            "Filtrar por talla",
            sorted(df["Talla"].dropna().unique().tolist(), key=_sz_key),
            placeholder="Todas las tallas visibles",
        )
    with mc2:
        bulk_colors = st.multiselect(
            "Filtrar por color",
            sorted(df["Color"].dropna().unique().tolist()),
            placeholder="Todos los colores visibles",
        )

    sub_bulk = df.copy()
    if bulk_tallas:
        sub_bulk = sub_bulk[sub_bulk["Talla"].isin(bulk_tallas)]
    if bulk_colors:
        sub_bulk = sub_bulk[sub_bulk["Color"].isin(bulk_colors)]
    st.caption(f"Afecta a **{len(sub_bulk)}** de {len(df)} variantes visibles")

    bc1, bc2, bc3, _s, bc4, bc5, bc6, _s2, bc7 = st.columns(
        [1, 1, 1, 0.3, 1, 1, 1, 0.3, 1.4]
    )
    if bc1.button("− 10", use_container_width=True):
        _bulk_apply(sub_bulk, -10)
    if bc2.button("− 5", use_container_width=True):
        _bulk_apply(sub_bulk, -5)
    if bc3.button("− 1", use_container_width=True):
        _bulk_apply(sub_bulk, -1)
    if bc4.button("+ 1", use_container_width=True):
        _bulk_apply(sub_bulk, 1)
    if bc5.button("+ 5", use_container_width=True):
        _bulk_apply(sub_bulk, 5)
    if bc6.button("+ 10", use_container_width=True):
        _bulk_apply(sub_bulk, 10)
    if bc7.button("Poner a 0", use_container_width=True):
        _bulk_apply(sub_bulk, None)

# Placeholder filled after matrix so totals reflect current state
summary_slot = st.empty()

# ── Matrix view ──────────────────────────────────────────────────────────────
for (ref, nombre), grp in df.groupby(["Referencia interna", "Nombre"], sort=True):
    colors = sorted(grp["Color"].dropna().unique().tolist())
    tallas = sorted(grp["Talla"].dropna().unique().tolist(), key=_sz_key)

    bc_map = {
        (r["Color"], r["Talla"]): r["Código de barras principal"]
        for _, r in grp.iterrows()
    }

    # Gray reference card
    st.markdown(
        f'<div class="ref-card">'
        f'<span class="ref-title">{ref} &nbsp;·&nbsp; {nombre}</span>'
        f"</div>",
        unsafe_allow_html=True,
    )

    # ratios: [talla label] + [one per color] + [total col]
    ratios = [0.8] + [2.5] * len(colors) + [1.0]

    # Header row
    hcols = st.columns(ratios)
    hcols[0].write("")
    for ci, color in enumerate(colors):
        hcols[ci + 1].markdown(
            f'<div class="col-hdr">{color}</div>', unsafe_allow_html=True
        )
    hcols[-1].markdown('<div class="col-hdr-tot">Total</div>', unsafe_allow_html=True)

    col_totals = [0] * len(colors)

    # One row per size
    for talla in tallas:
        rcols = st.columns(ratios)
        rcols[0].markdown(f'<div class="sz-lbl">{talla}</div>', unsafe_allow_html=True)
        row_total = 0

        for ci, color in enumerate(colors):
            bc = bc_map.get((color, talla))
            if bc is None:
                continue
            sk = f"qty_{bc}"
            if sk not in st.session_state:
                st.session_state[sk] = 0
            with rcols[ci + 1]:
                v = st.number_input(
                    talla,
                    min_value=0,
                    step=1,
                    key=sk,
                    label_visibility="collapsed",
                )
            row_total += int(v)
            col_totals[ci] += int(v)

        cls = "tot-val" if row_total > 0 else "tot-zero"
        rcols[-1].markdown(
            f'<div class="{cls}">{row_total if row_total > 0 else "—"}</div>',
            unsafe_allow_html=True,
        )

    # Totals row
    tcols = st.columns(ratios)
    tcols[0].markdown('<div class="tot-lbl">Total</div>', unsafe_allow_html=True)
    grand = sum(col_totals)
    for ci, ct in enumerate(col_totals):
        cls = "tot-val" if ct > 0 else "tot-zero"
        tcols[ci + 1].markdown(
            f'<div class="{cls}">{ct if ct > 0 else "—"}</div>',
            unsafe_allow_html=True,
        )
    tcols[-1].markdown(
        f'<div class="grand-tot">{grand if grand > 0 else "—"}</div>',
        unsafe_allow_html=True,
    )

# Fill summary after all inputs are rendered (values are now current)
fresh_qtys = {
    k[4:]: int(v)
    for k, v in st.session_state.items()
    if k.startswith("qty_") and isinstance(v, (int, float)) and v > 0
}
if fresh_qtys:
    summary_slot.info(
        f"**{len(fresh_qtys)}** variantes con cantidad asignada — "
        f"**{sum(fresh_qtys.values())}** unidades totales"
    )

st.divider()

# ── SECTION 2 — Calculate ────────────────────────────────────────────────────
if st.button("Calcular Consumos", type="primary", use_container_width=True):
    final_qtys = {
        k[4:]: int(v)
        for k, v in st.session_state.items()
        if k.startswith("qty_") and isinstance(v, (int, float)) and v > 0
    }
    if not final_qtys:
        st.warning("Introduce la cantidad a producir en al menos una variante.")
        st.stop()

    comp_totals: dict[str, float] = {}
    for bc, qty in final_qtys.items():
        for _, brow in bom[bom["Cod Barras Variante"] == bc].iterrows():
            comp = str(brow["EAN Componente"])
            comp_totals[comp] = comp_totals.get(comp, 0.0) + float(brow["Cantidad"]) * qty

    results = []
    for comp_bc, total_qty in comp_totals.items():
        info = barcode_lookup.get(comp_bc, {})
        col_ = str(info.get("Color", "")).strip()
        tal_ = str(info.get("Talla", "")).strip()
        results.append(
            {
                "Referencia": info.get("Referencia interna", ""),
                "Nombre": info.get("Nombre", f"Componente {comp_bc}"),
                "Color": col_ if col_ not in ("", "nan") else "—",
                "Talla": tal_ if tal_ not in ("", "nan") else "—",
                "Código de barras": comp_bc,
                "Cantidad necesaria": round(total_qty, 4),
            }
        )

    results_df = (
        pd.DataFrame(results).sort_values(["Nombre", "Color", "Talla"]).reset_index(drop=True)
    )

    st.header("2. Necesidades de componentes")
    st.caption(f"{len(results_df)} componentes distintos necesarios para fabricar la colección")
    st.dataframe(results_df, hide_index=True, use_container_width=True)

    st.subheader("Resumen del plan de producción")
    plan_rows = []
    for bc, qty in final_qtys.items():
        info = barcode_lookup.get(bc, {})
        plan_rows.append(
            {
                "Referencia": info.get("Referencia interna", ""),
                "Nombre": info.get("Nombre", ""),
                "Color": info.get("Color", ""),
                "Talla": info.get("Talla", ""),
                "Código de barras": bc,
                "Unidades": qty,
            }
        )
    plan_df = pd.DataFrame(plan_rows).sort_values(["Nombre", "Color", "Talla"])
    st.dataframe(plan_df, hide_index=True, use_container_width=True)

    st.subheader("Descargar resultados")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Descargar consumos CSV",
            results_df.to_csv(index=False).encode("utf-8-sig"),
            "consumos.csv",
            "text/csv",
            use_container_width=True,
        )
    with c2:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            plan_df.to_excel(writer, sheet_name="Plan de producción", index=False)
            results_df.to_excel(writer, sheet_name="Consumos", index=False)
        st.download_button(
            "Descargar Excel completo",
            buf.getvalue(),
            "consumos.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
