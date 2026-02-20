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
/* ── Reference card (gray) ── */
.ref-card {
    background: #e4e4e4;
    border-left: 4px solid #999;
    border-radius: 0 6px 6px 0;
    padding: 7px 14px;
    margin: 16px 0 4px 0;
}
.ref-title { font-size: 0.92em; font-weight: 700; color: #333; }

/* ── Color column header ── */
.col-hdr {
    background: #f4f4f4;
    border: 1px solid #d8d8d8;
    border-radius: 4px;
    text-align: center;
    padding: 3px 4px;
    font-size: 0.78em;
    font-weight: 600;
    color: #555;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* ── Size label ── */
.sz-lbl {
    text-align: right;
    font-weight: 700;
    font-size: 0.82em;
    color: #444;
    padding: 8px 8px 0 0;
}

/* ── Quantity display ── */
.qty-val {
    text-align: center;
    font-size: 1em;
    font-weight: 700;
    color: #111;
    padding-top: 6px;
}
.qty-zero { color: #ccc; }

/* ── Compact +/- buttons ── */
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
    variantes["Color"] = (
        variantes["Color"].str.replace("Color: ", "", regex=False).str.strip()
    )
    variantes["Talla"] = (
        variantes["Talla"].str.replace("Talla: ", "", regex=False).str.strip()
    )

    barcode_lookup = variantes.set_index("Código de barras principal").to_dict("index")
    bom_barcodes = set(bom["Cod Barras Variante"].unique())
    finished = variantes[
        variantes["Código de barras principal"].isin(bom_barcodes)
    ].copy()

    return variantes, bom, barcode_lookup, finished


variantes, bom, barcode_lookup, finished = load_data()

# ── Size sort order ──────────────────────────────────────────────────────────
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
    "Filtra las referencias que quieres fabricar e introduce la cantidad "
    "usando los botones **+** / **−**. Pulsa **Calcular Consumos** cuando termines."
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

# Placeholder for the summary info — filled after all buttons are processed
summary_slot = st.empty()

# ── Matrix view ──────────────────────────────────────────────────────────────
for (ref, nombre), grp in df.groupby(["Referencia interna", "Nombre"], sort=True):
    colors = sorted(grp["Color"].dropna().unique().tolist())
    tallas = sorted(grp["Talla"].dropna().unique().tolist(), key=_sz_key)

    # Gray reference card
    st.markdown(
        f'<div class="ref-card">'
        f'<span class="ref-title">{ref} &nbsp;·&nbsp; {nombre}</span>'
        f"</div>",
        unsafe_allow_html=True,
    )

    ratios = [0.7] + [2.0] * len(colors)

    # Header row: color names
    hcols = st.columns(ratios)
    hcols[0].write("")
    for ci, color in enumerate(colors):
        hcols[ci + 1].markdown(
            f'<div class="col-hdr">{color}</div>', unsafe_allow_html=True
        )

    # One row per size
    for talla in tallas:
        rcols = st.columns(ratios)
        rcols[0].markdown(
            f'<div class="sz-lbl">{talla}</div>', unsafe_allow_html=True
        )

        for ci, color in enumerate(colors):
            match = grp[(grp["Color"] == color) & (grp["Talla"] == talla)]
            if match.empty:
                rcols[ci + 1].markdown(
                    '<div style="min-height:34px"></div>', unsafe_allow_html=True
                )
                continue

            bc = match.iloc[0]["Código de barras principal"]
            sk = f"qty_{bc}"
            if sk not in st.session_state:
                st.session_state[sk] = 0

            with rcols[ci + 1]:
                b0, b1, b2 = st.columns([1, 1, 1])

                # Render both buttons first, then process clicks, then display
                minus = b0.button("−", key=f"m_{bc}", use_container_width=True)
                plus = b2.button("+", key=f"p_{bc}", use_container_width=True)

                if minus:
                    st.session_state[sk] = max(0, st.session_state[sk] - 1)
                if plus:
                    st.session_state[sk] += 1

                v = st.session_state[sk]
                cls = "qty-val" if v > 0 else "qty-val qty-zero"
                b1.markdown(f'<div class="{cls}">{v}</div>', unsafe_allow_html=True)

# After all buttons are processed, update the summary with fresh totals
fresh_qtys = {
    k[4:]: v
    for k, v in st.session_state.items()
    if k.startswith("qty_") and isinstance(v, int) and v > 0
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
        k[4:]: v
        for k, v in st.session_state.items()
        if k.startswith("qty_") and isinstance(v, int) and v > 0
    }
    if not final_qtys:
        st.warning("Introduce la cantidad a producir en al menos una variante.")
        st.stop()

    # Aggregate component requirements
    comp_totals: dict[str, float] = {}
    for bc, qty in final_qtys.items():
        for _, brow in bom[bom["Cod Barras Variante"] == bc].iterrows():
            comp = str(brow["EAN Componente"])
            comp_totals[comp] = comp_totals.get(comp, 0.0) + float(brow["Cantidad"]) * qty

    # Resolve component details
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

    results_df = pd.DataFrame(results).sort_values(
        ["Nombre", "Color", "Talla"]
    ).reset_index(drop=True)

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
