import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(
    page_title="Planificación de Consumos",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.title("Planificación de Producción — Cálculo de Consumos")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VARIANTES_PATH = os.path.join(BASE_DIR, "Variantes.xlsx")
BOM_PATH = os.path.join(BASE_DIR, "listamateriales.xlsx")


@st.cache_data
def load_data():
    variantes = pd.read_excel(VARIANTES_PATH, sheet_name="Variantes de producto", dtype=str)
    bom = pd.read_excel(BOM_PATH, sheet_name="Fichero ejemplo")

    # Normalize barcode columns to string (BOM columns are numeric in the file)
    bom["Cod Barras Variante"] = bom["Cod Barras Variante"].astype(str).str.strip()
    bom["EAN Componente"] = bom["EAN Componente"].astype(str).str.strip()
    variantes["Código de barras principal"] = (
        variantes["Código de barras principal"].astype(str).str.strip()
    )

    # Clean display values (strip "Color: " and "Talla: " prefixes)
    variantes["Color"] = (
        variantes["Color"].str.replace("Color: ", "", regex=False).str.strip()
    )
    variantes["Talla"] = (
        variantes["Talla"].str.replace("Talla: ", "", regex=False).str.strip()
    )

    # Barcode → product info lookup (used for resolving component details)
    barcode_lookup = variantes.set_index("Código de barras principal").to_dict("index")

    # Only show finished products that have a BOM entry
    bom_barcodes = set(bom["Cod Barras Variante"].unique())
    finished = variantes[
        variantes["Código de barras principal"].isin(bom_barcodes)
    ].copy()

    return variantes, bom, barcode_lookup, finished


variantes, bom, barcode_lookup, finished = load_data()

# ---------------------------------------------------------------------------
# SECTION 1 — Production plan
# ---------------------------------------------------------------------------
st.header("1. Plan de producción")
st.markdown(
    "Filtra las variantes de prenda que quieres fabricar e introduce la cantidad "
    "a producir de cada una. Cuando hayas terminado, pulsa **Calcular Consumos**."
)

# --- Filters ---
col1, col2, col3, col4 = st.columns([3, 2, 2, 2])
with col1:
    q = st.text_input("Buscar por nombre", placeholder="ej. Vestido Goya…")
with col2:
    colores = ["Todos"] + sorted(
        finished["Color"].dropna().unique().tolist()
    )
    sel_color = st.selectbox("Color", colores)
with col3:
    tallas = ["Todas"] + sorted(
        finished["Talla"].dropna().unique().tolist()
    )
    sel_talla = st.selectbox("Talla", tallas)
with col4:
    refs = ["Todas"] + sorted(
        finished["Referencia interna"].dropna().unique().tolist()
    )
    sel_ref = st.selectbox("Referencia", refs)

# Apply filters
df = finished.copy()
if q:
    df = df[df["Nombre"].str.contains(q, case=False, na=False)]
if sel_color != "Todos":
    df = df[df["Color"] == sel_color]
if sel_talla != "Todas":
    df = df[df["Talla"] == sel_talla]
if sel_ref != "Todas":
    df = df[df["Referencia interna"] == sel_ref]

# --- Editable table ---
display = df[
    ["Referencia interna", "Nombre", "Color", "Talla", "Código de barras principal"]
].copy()
display.columns = ["Referencia", "Nombre", "Color", "Talla", "Código de barras"]
display["Cantidad a producir"] = 0

st.caption(f"{len(display)} variantes mostradas")

edited = st.data_editor(
    display,
    column_config={
        "Referencia": st.column_config.TextColumn(width="small"),
        "Nombre": st.column_config.TextColumn(width="large"),
        "Color": st.column_config.TextColumn(width="small"),
        "Talla": st.column_config.TextColumn(width="small"),
        "Código de barras": st.column_config.TextColumn(width="medium"),
        "Cantidad a producir": st.column_config.NumberColumn(
            "Cantidad a producir", min_value=0, step=1, default=0
        ),
    },
    disabled=["Referencia", "Nombre", "Color", "Talla", "Código de barras"],
    hide_index=True,
    use_container_width=True,
    key="plan",
)

to_produce = edited[edited["Cantidad a producir"] > 0]
if len(to_produce) > 0:
    total_units = int(to_produce["Cantidad a producir"].sum())
    st.info(
        f"**{len(to_produce)}** variantes seleccionadas — "
        f"**{total_units}** unidades totales a producir"
    )

st.divider()

# ---------------------------------------------------------------------------
# SECTION 2 — Calculate
# ---------------------------------------------------------------------------
if st.button("Calcular Consumos", type="primary", use_container_width=True):
    if to_produce.empty:
        st.warning("Introduce la cantidad a producir en al menos una variante.")
        st.stop()

    # Aggregate component requirements
    comp_totals: dict[str, float] = {}
    for _, row in to_produce.iterrows():
        barcode = str(row["Código de barras"])
        qty = float(row["Cantidad a producir"])
        bom_rows = bom[bom["Cod Barras Variante"] == barcode]
        for _, brow in bom_rows.iterrows():
            comp = str(brow["EAN Componente"])
            comp_totals[comp] = comp_totals.get(comp, 0.0) + float(brow["Cantidad"]) * qty

    # Resolve component details from Variantes
    results = []
    for comp_bc, total_qty in comp_totals.items():
        info = barcode_lookup.get(comp_bc, {})
        color = str(info.get("Color", "")).strip()
        talla = str(info.get("Talla", "")).strip()
        results.append(
            {
                "Referencia": info.get("Referencia interna", ""),
                "Nombre": info.get("Nombre", f"Componente {comp_bc}"),
                "Color": color if color not in ("", "nan") else "—",
                "Talla": talla if talla not in ("", "nan") else "—",
                "Código de barras": comp_bc,
                "Cantidad necesaria": round(total_qty, 4),
            }
        )

    results_df = pd.DataFrame(results).sort_values(
        ["Nombre", "Color", "Talla"]
    ).reset_index(drop=True)

    st.header("2. Necesidades de componentes")
    st.caption(
        f"{len(results_df)} componentes distintos necesarios para fabricar la colección"
    )
    st.dataframe(results_df, hide_index=True, use_container_width=True)

    # --- Plan de producción detallado ---
    st.subheader("Resumen del plan de producción")
    plan_detail = []
    for _, row in to_produce.iterrows():
        bc = str(row["Código de barras"])
        info = barcode_lookup.get(bc, {})
        plan_detail.append(
            {
                "Referencia": info.get("Referencia interna", row["Referencia"]),
                "Nombre": info.get("Nombre", row["Nombre"]),
                "Color": info.get("Color", row["Color"]),
                "Talla": info.get("Talla", row["Talla"]),
                "Código de barras": bc,
                "Unidades": int(row["Cantidad a producir"]),
            }
        )
    plan_df = pd.DataFrame(plan_detail).sort_values(["Nombre", "Color", "Talla"])
    st.dataframe(plan_df, hide_index=True, use_container_width=True)

    # --- Downloads ---
    st.subheader("Descargar resultados")
    c1, c2 = st.columns(2)

    with c1:
        csv_bytes = results_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Descargar consumos CSV",
            csv_bytes,
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
