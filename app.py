import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
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

/* ── BOM missing warning ── */
.bom-warn { font-size: 0.76em; color: #c0392b; font-weight: 600; margin-left: 10px; }

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


# ── TABS ─────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["Plan de producción", "Crear / Editar BOM"])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Plan de producción + Cálculo de consumos
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:

    # Indicator when a custom BOM is active
    if st.session_state.get("custom_bom") is not None:
        n_lines = len(st.session_state["custom_bom"])
        st.info(
            f"BOM personalizada activa ({n_lines} entradas). "
            "El cálculo usará esta BOM en lugar de la BOM por defecto. "
            "Ve a la pestaña 'Crear / Editar BOM' para modificarla o desactivarla."
        )

    # ── Bulk assignment helper ────────────────────────────────────────────────
    def _bulk_apply(target_df, delta):
        """delta=int → add/subtract; delta=None → reset to 0."""
        for bc in target_df["Código de barras principal"]:
            sk = f"qty_{bc}"
            cur = st.session_state.get(sk, 0)
            st.session_state[sk] = 0 if delta is None else max(0, cur + delta)

    # ── SECTION 1 — Filters ──────────────────────────────────────────────────
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

    # ── Bulk assignment ───────────────────────────────────────────────────────
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

    # ── Import plan from Excel / CSV ─────────────────────────────────────────
    with st.expander("Importar plan desde Excel / CSV", expanded=False):
        st.markdown(
            "El fichero debe tener al menos dos columnas: una con el **código de barras** "
            "y otra con las **unidades** a producir. "
            "Se aceptan nombres de columna en español o inglés."
        )
        up_plan = st.file_uploader(
            "Fichero Excel o CSV",
            type=["xlsx", "csv"],
            key="upload_plan",
            label_visibility="collapsed",
        )
        if up_plan is not None:
            if st.button("Cargar plan", key="btn_import_plan", use_container_width=True):
                try:
                    if up_plan.name.lower().endswith(".csv"):
                        _df_imp = pd.read_csv(up_plan, dtype=str)
                    else:
                        _df_imp = pd.read_excel(up_plan, dtype=str)
                    _bc_col = next(
                        (c for c in _df_imp.columns
                         if any(k in c.lower() for k in ("barras", "ean", "barcode", "código"))),
                        _df_imp.columns[0],
                    )
                    _qty_col = next(
                        (c for c in _df_imp.columns
                         if any(k in c.lower() for k in ("unidad", "cantidad", "qty", "units"))),
                        _df_imp.columns[1],
                    )
                    _loaded = 0
                    for _, _imp_row in _df_imp.iterrows():
                        _bc = str(_imp_row[_bc_col]).strip()
                        try:
                            _qty = int(float(str(_imp_row[_qty_col]).strip()))
                        except ValueError:
                            continue
                        if _qty > 0 and _bc in barcode_lookup:
                            st.session_state[f"qty_{_bc}"] = _qty
                            _loaded += 1
                    st.success(f"{_loaded} variantes cargadas desde el fichero.")
                    st.rerun()
                except Exception as _e:
                    st.error(f"Error al leer el fichero: {_e}")

    # Placeholder filled after matrix so totals reflect current state
    summary_slot = st.empty()

    # ── BOM coverage set (used for matrix indicators) ────────────────────────
    _active_bom_cov = st.session_state.get("custom_bom", bom)
    _bom_covered = set(_active_bom_cov["Cod Barras Variante"].unique())

    # ── Matrix view ──────────────────────────────────────────────────────────
    for (ref, nombre), grp in df.groupby(["Referencia interna", "Nombre"], sort=True):
        colors = sorted(grp["Color"].dropna().unique().tolist())
        tallas = sorted(grp["Talla"].dropna().unique().tolist(), key=_sz_key)

        bc_map = {
            (r["Color"], r["Talla"]): r["Código de barras principal"]
            for _, r in grp.iterrows()
        }

        # Gray reference card (with BOM coverage indicator)
        _missing_bom = sum(1 for bc in bc_map.values() if bc not in _bom_covered)
        _bom_warn_html = (
            f'<span class="bom-warn">(sin BOM: {_missing_bom} var.)</span>'
            if _missing_bom else ""
        )
        st.markdown(
            f'<div class="ref-card">'
            f'<span class="ref-title">{ref} &nbsp;·&nbsp; {nombre}</span>'
            f'{_bom_warn_html}'
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

    # ── SECTION 2 — Calculate ────────────────────────────────────────────────
    _bcol1, _bcol2 = st.columns([4, 1])
    with _bcol1:
        _calcular = st.button("Calcular Consumos", type="primary", use_container_width=True)
    with _bcol2:
        if st.button("Limpiar plan", use_container_width=True):
            for _k in list(st.session_state.keys()):
                if _k.startswith("qty_"):
                    del st.session_state[_k]
            for _k in ("results_df", "plan_df", "no_bom_warn"):
                st.session_state.pop(_k, None)
            st.rerun()

    if _calcular:
        final_qtys = {
            k[4:]: int(v)
            for k, v in st.session_state.items()
            if k.startswith("qty_") and isinstance(v, (int, float)) and v > 0
        }
        if not final_qtys:
            st.warning("Introduce la cantidad a producir en al menos una variante.")
        else:
            active_bom = st.session_state.get("custom_bom", bom)

            # Detect variants without BOM entries
            no_bom = [bc for bc in final_qtys
                      if active_bom[active_bom["Cod Barras Variante"] == bc].empty]
            st.session_state["no_bom_warn"] = no_bom

            comp_totals: dict[str, float] = {}
            for bc, qty in final_qtys.items():
                for _, brow in active_bom[active_bom["Cod Barras Variante"] == bc].iterrows():
                    comp = str(brow["EAN Componente"])
                    comp_totals[comp] = comp_totals.get(comp, 0.0) + float(brow["Cantidad"]) * qty

            results = []
            for comp_bc, total_qty in comp_totals.items():
                info = barcode_lookup.get(comp_bc, {})
                col_ = str(info.get("Color", "")).strip()
                tal_ = str(info.get("Talla", "")).strip()
                results.append({
                    "Referencia": info.get("Referencia interna", ""),
                    "Nombre": info.get("Nombre", f"Componente {comp_bc}"),
                    "Color": col_ if col_ not in ("", "nan") else "—",
                    "Talla": tal_ if tal_ not in ("", "nan") else "—",
                    "Código de barras": comp_bc,
                    "Cantidad necesaria": round(total_qty, 4),
                })

            st.session_state["results_df"] = (
                pd.DataFrame(results)
                .sort_values(["Nombre", "Color", "Talla"])
                .reset_index(drop=True)
            )

            plan_rows = []
            for bc, qty in final_qtys.items():
                info = barcode_lookup.get(bc, {})
                plan_rows.append({
                    "Referencia": info.get("Referencia interna", ""),
                    "Nombre": info.get("Nombre", ""),
                    "Color": info.get("Color", ""),
                    "Talla": info.get("Talla", ""),
                    "Código de barras": bc,
                    "Unidades": qty,
                })
            st.session_state["plan_df"] = (
                pd.DataFrame(plan_rows).sort_values(["Nombre", "Color", "Talla"])
            )

    # ── Results (persistent — survive widget interactions) ───────────────────
    if "results_df" in st.session_state:
        results_df = st.session_state["results_df"]
        plan_df = st.session_state["plan_df"]

        # Warning: variants without BOM
        _no_bom = st.session_state.get("no_bom_warn", [])
        if _no_bom:
            _names = [barcode_lookup.get(bc, {}).get("Nombre", bc) for bc in _no_bom[:5]]
            _extra = f" y {len(_no_bom) - 5} más" if len(_no_bom) > 5 else ""
            st.warning(
                f"{len(_no_bom)} variante(s) sin entradas en la BOM — "
                f"sus consumos no se han calculado: {', '.join(_names)}{_extra}."
            )

        st.header("2. Necesidades de componentes")
        st.caption(f"{len(results_df)} componentes distintos necesarios para fabricar la colección")
        st.dataframe(results_df, hide_index=True, use_container_width=True)

        st.subheader("Resumen del plan de producción")
        st.dataframe(plan_df, hide_index=True, use_container_width=True)

        st.subheader("Descargar resultados")
        pet_fecha = st.date_input(
            "Fecha de transferencia (plantilla PET)",
            value=date.today(),
            key="pet_fecha",
        )
        c1, c2, c3 = st.columns(3)
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
        with c3:
            pet_df = pd.DataFrame([
                {
                    "Fecha": pet_fecha,
                    "Almacén de origen": "PET Almacén Ibiza",
                    "Almacén de destino": "PET Almacén Túnez",
                    "Observaciones": "",
                    "EAN": row["Código de barras"],
                    "Cantidad": row["Cantidad necesaria"],
                }
                for _, row in results_df.iterrows()
            ])
            buf_pet = BytesIO()
            with pd.ExcelWriter(buf_pet, engine="openpyxl") as writer:
                pet_df.to_excel(writer, sheet_name="Fichero ejemplo", index=False)
            st.download_button(
                "Descargar plantilla PET",
                buf_pet.getvalue(),
                f"PET_{pet_fecha.strftime('%Y%m%d')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Crear / Editar BOM
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header("Crear / Editar Lista de Materiales")
    st.markdown(
        "Define los componentes que lleva cada variante de prenda para construir "
        "tu propia lista de materiales (BOM) desde cero. "
        "Una vez completa, actívala para usarla en el cálculo de consumos."
    )

    # Initialize BOM draft in session state
    if "bom_draft" not in st.session_state:
        st.session_state["bom_draft"] = []

    # ── Build variant option lists ────────────────────────────────────────────
    all_variants = variantes.dropna(subset=["Código de barras principal"]).copy()
    all_variants["_label"] = (
        all_variants["Referencia interna"].fillna("").str.strip()
        + " | "
        + all_variants["Nombre"].fillna("").str.strip()
        + " | "
        + all_variants["Color"].fillna("").str.strip()
        + " | "
        + all_variants["Talla"].fillna("").str.strip()
    )
    variant_label_list = sorted(all_variants["_label"].unique().tolist())
    variant_label_to_bc = dict(
        zip(all_variants["_label"], all_variants["Código de barras principal"])
    )

    # ── Load existing BOM as starting point ──────────────────────────────────
    with st.expander("Cargar BOM existente como base de partida"):
        lc1, lc2 = st.columns(2)

        with lc1:
            st.markdown("**Cargar la BOM por defecto** (`listamateriales.xlsx`)")
            if st.button("Cargar BOM por defecto", use_container_width=True):
                entries = []
                for _, row in bom.iterrows():
                    bc_var = str(row["Cod Barras Variante"])
                    bc_comp = str(row["EAN Componente"])
                    qty = float(row["Cantidad"])
                    var_info = barcode_lookup.get(bc_var, {})
                    comp_info = barcode_lookup.get(bc_comp, {})
                    var_label = " | ".join(
                        x for x in [
                            str(var_info.get("Referencia interna", "")).strip(),
                            str(var_info.get("Nombre", bc_var)).strip(),
                            str(var_info.get("Color", "")).strip(),
                            str(var_info.get("Talla", "")).strip(),
                        ] if x and x != "nan"
                    ) or bc_var
                    comp_label = " | ".join(
                        x for x in [
                            str(comp_info.get("Referencia interna", "")).strip(),
                            str(comp_info.get("Nombre", bc_comp)).strip(),
                            str(comp_info.get("Color", "")).strip(),
                            str(comp_info.get("Talla", "")).strip(),
                        ] if x and x != "nan"
                    ) or bc_comp
                    entries.append({
                        "Cod Barras Variante": bc_var,
                        "EAN Componente": bc_comp,
                        "Cantidad": qty,
                        "_nombre_variante": var_label,
                        "_nombre_componente": comp_label,
                    })
                st.session_state["bom_draft"] = entries
                st.success(f"BOM cargada: {len(entries)} entradas.")
                st.rerun()

        with lc2:
            st.markdown("**O sube un fichero Excel** (mismo formato que `listamateriales.xlsx`)")
            uploaded = st.file_uploader(
                "Fichero Excel",
                type=["xlsx"],
                key="bom_upload",
                label_visibility="collapsed",
            )
            if uploaded is not None:
                if st.button("Cargar fichero subido", use_container_width=True):
                    try:
                        uploaded_df = pd.read_excel(uploaded)
                        uploaded_df["Cod Barras Variante"] = (
                            uploaded_df["Cod Barras Variante"].astype(str).str.strip()
                        )
                        uploaded_df["EAN Componente"] = (
                            uploaded_df["EAN Componente"].astype(str).str.strip()
                        )
                        entries = []
                        for _, row in uploaded_df.iterrows():
                            bc_var = str(row["Cod Barras Variante"])
                            bc_comp = str(row["EAN Componente"])
                            qty = float(row["Cantidad"])
                            var_info = barcode_lookup.get(bc_var, {})
                            comp_info = barcode_lookup.get(bc_comp, {})
                            var_label = str(var_info.get("Nombre", bc_var))
                            comp_label = str(comp_info.get("Nombre", bc_comp))
                            entries.append({
                                "Cod Barras Variante": bc_var,
                                "EAN Componente": bc_comp,
                                "Cantidad": qty,
                                "_nombre_variante": var_label,
                                "_nombre_componente": comp_label,
                            })
                        st.session_state["bom_draft"] = entries
                        st.success(f"Fichero cargado: {len(entries)} entradas.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al leer el fichero: {e}")

    st.divider()

    # ── Form to add a new BOM entry ──────────────────────────────────────────
    st.subheader("Añadir componente")

    fc1, fc2, fc3 = st.columns([3, 3, 1.5])

    with fc1:
        st.markdown("**Prenda (variante de producto)**")
        sel_var_label = st.selectbox(
            "Variante",
            options=variant_label_list,
            label_visibility="collapsed",
            key="bom_sel_variant",
        )
        sel_var_bc = variant_label_to_bc.get(sel_var_label, "")

    with fc2:
        st.markdown("**Componente / Material**")
        comp_source = st.radio(
            "Origen del componente",
            ["Del catálogo de variantes", "Introducir manualmente"],
            horizontal=True,
            label_visibility="collapsed",
            key="bom_comp_source",
        )
        if comp_source == "Del catálogo de variantes":
            sel_comp_label = st.selectbox(
                "Componente del catálogo",
                options=variant_label_list,
                label_visibility="collapsed",
                key="bom_sel_comp",
            )
            comp_ean = variant_label_to_bc.get(sel_comp_label, "")
            comp_display_name = sel_comp_label
        else:
            comp_ean = st.text_input(
                "EAN / código del componente",
                key="bom_comp_ean",
                placeholder="ej. 8412345678901",
            )
            comp_display_name = st.text_input(
                "Nombre del componente",
                key="bom_comp_name",
                placeholder="ej. Tejido principal",
            )

    with fc3:
        st.markdown("**Cantidad por unidad**")
        comp_qty = st.number_input(
            "Cantidad",
            min_value=0.001,
            step=0.1,
            value=1.0,
            key="bom_qty",
            label_visibility="collapsed",
            format="%.3f",
        )
        st.write("")  # vertical spacer
        add_clicked = st.button(
            "Añadir", type="primary", use_container_width=True, key="bom_add"
        )

    if add_clicked:
        if not str(comp_ean).strip():
            st.error("El EAN del componente no puede estar vacío.")
        else:
            already = any(
                e["Cod Barras Variante"] == sel_var_bc
                and e["EAN Componente"] == str(comp_ean).strip()
                for e in st.session_state["bom_draft"]
            )
            if already:
                st.warning(
                    "Este componente ya está asignado a esta variante. "
                    "Elimina la entrada existente si quieres cambiar la cantidad."
                )
            else:
                st.session_state["bom_draft"].append({
                    "Cod Barras Variante": sel_var_bc,
                    "EAN Componente": str(comp_ean).strip(),
                    "Cantidad": comp_qty,
                    "_nombre_variante": sel_var_label,
                    "_nombre_componente": comp_display_name or str(comp_ean).strip(),
                })
                st.rerun()

    st.divider()

    # ── Bulk component assignment ─────────────────────────────────────────────
    with st.expander("Asignación masiva de componentes", expanded=False):
        st.markdown(
            "Aplica el mismo componente a múltiples prendas a la vez. "
            "Filtra por **prefijo EAN** (colección) o por referencia, nombre, color o talla."
        )

        # ── Step 1: target selection ──────────────────────────────────────────
        st.markdown("**Paso 1 — Seleccionar prendas objetivo**")

        # Primary: EAN prefix (collection selector)
        _pfx_c1, _pfx_c2 = st.columns([2, 3])
        with _pfx_c1:
            bulk_ean_prefix = st.text_input(
                "Prefijo de referencia interna (colección)",
                placeholder="ej. 261…",
                key="bulk_ean_prefix",
                help=(
                    "Los primeros dígitos de la referencia interna identifican la colección. "
                    "Ej.: 261 → V&G 2026."
                ),
            )
        with _pfx_c2:
            if bulk_ean_prefix:
                _pfx_preview = all_variants[
                    all_variants["Referencia interna"]
                    .astype(str)
                    .str.startswith(bulk_ean_prefix.strip())
                ]
                st.caption(
                    f"{len(_pfx_preview)} variantes cuya referencia comienza por **{bulk_ean_prefix.strip()}**"
                )

        st.markdown("— o filtra por atributos —")
        bf1, bf2 = st.columns(2)
        with bf1:
            bulk_refs = st.multiselect(
                "Referencia interna",
                sorted(all_variants["Referencia interna"].dropna().unique().tolist()),
                placeholder="Todas las referencias",
                key="bulk_refs",
            )
            bulk_q = st.text_input(
                "Buscar por nombre",
                placeholder="ej. Vestido Goya…",
                key="bulk_q",
            )
        with bf2:
            bulk_colors = st.multiselect(
                "Color",
                sorted(all_variants["Color"].dropna().unique().tolist()),
                placeholder="Todos los colores",
                key="bulk_colors",
            )
            bulk_tallas = st.multiselect(
                "Talla",
                sorted(all_variants["Talla"].dropna().unique().tolist(), key=_sz_key),
                placeholder="Todas las tallas",
                key="bulk_tallas",
            )

        # Apply filters
        bulk_target = all_variants.copy()
        if bulk_ean_prefix:
            bulk_target = bulk_target[
                bulk_target["Referencia interna"]
                .astype(str)
                .str.startswith(bulk_ean_prefix.strip())
            ]
        if bulk_refs:
            bulk_target = bulk_target[bulk_target["Referencia interna"].isin(bulk_refs)]
        if bulk_q:
            bulk_target = bulk_target[
                bulk_target["Nombre"].str.contains(bulk_q, case=False, na=False)
            ]
        if bulk_colors:
            bulk_target = bulk_target[bulk_target["Color"].isin(bulk_colors)]
        if bulk_tallas:
            bulk_target = bulk_target[bulk_target["Talla"].isin(bulk_tallas)]

        n_target = len(bulk_target)
        st.caption(f"**{n_target}** prendas seleccionadas")
        if n_target > 0 and n_target <= 30:
            st.dataframe(
                bulk_target[["Referencia interna", "Nombre", "Color", "Talla"]]
                .reset_index(drop=True),
                hide_index=True,
                use_container_width=True,
                height=min(38 * n_target + 38, 300),
            )
        elif n_target > 30:
            st.caption(
                f"Demasiadas filas para previsualizar — "
                f"se asignará el componente a las {n_target} prendas filtradas."
            )

        st.markdown("**Paso 2 — Componente a asignar**")
        bc1, bc2, bc3, bc4 = st.columns([2.5, 2.5, 1.2, 1.5])

        with bc1:
            bulk_comp_source = st.radio(
                "Origen",
                ["Del catálogo", "EAN manual"],
                horizontal=True,
                label_visibility="collapsed",
                key="bulk_comp_source",
            )
            if bulk_comp_source == "Del catálogo":
                bulk_sel_comp = st.selectbox(
                    "Componente del catálogo",
                    options=variant_label_list,
                    label_visibility="collapsed",
                    key="bulk_sel_comp",
                )
                bulk_comp_ean = variant_label_to_bc.get(bulk_sel_comp, "")
                bulk_comp_name = bulk_sel_comp
            else:
                bulk_comp_ean = st.text_input(
                    "EAN del componente",
                    placeholder="ej. 8412345678901",
                    key="bulk_comp_ean",
                )
                bulk_comp_name = st.text_input(
                    "Nombre del componente",
                    placeholder="ej. Tejido principal",
                    key="bulk_comp_name",
                )

        with bc2:
            st.markdown("Cantidad por unidad")
            bulk_qty = st.number_input(
                "Cantidad",
                min_value=0.001,
                step=0.1,
                value=1.0,
                key="bulk_qty",
                label_visibility="collapsed",
                format="%.3f",
            )

        with bc3:
            st.markdown("Si ya existe")
            bulk_dup = st.radio(
                "Si ya existe",
                ["Omitir", "Sobreescribir"],
                key="bulk_dup",
                label_visibility="collapsed",
            )

        with bc4:
            st.write("")
            st.write("")
            bulk_assign = st.button(
                f"Asignar a {n_target} prenda{'s' if n_target != 1 else ''}",
                type="primary",
                use_container_width=True,
                key="bulk_assign",
                disabled=(n_target == 0),
            )

        if bulk_assign:
            if not str(bulk_comp_ean).strip():
                st.error("El EAN del componente no puede estar vacío.")
            else:
                added = 0
                updated = 0
                skipped = 0
                comp_ean_clean = str(bulk_comp_ean).strip()
                comp_name_final = bulk_comp_name or comp_ean_clean

                for _, var_row in bulk_target.iterrows():
                    bc_var = var_row["Código de barras principal"]
                    existing_idx = next(
                        (
                            i for i, e in enumerate(st.session_state["bom_draft"])
                            if e["Cod Barras Variante"] == bc_var
                            and e["EAN Componente"] == comp_ean_clean
                        ),
                        None,
                    )
                    if existing_idx is not None:
                        if bulk_dup == "Sobreescribir":
                            st.session_state["bom_draft"][existing_idx]["Cantidad"] = bulk_qty
                            updated += 1
                        else:
                            skipped += 1
                    else:
                        st.session_state["bom_draft"].append({
                            "Cod Barras Variante": bc_var,
                            "EAN Componente": comp_ean_clean,
                            "Cantidad": bulk_qty,
                            "_nombre_variante": var_row["_label"],
                            "_nombre_componente": comp_name_final,
                        })
                        added += 1

                parts = []
                if added:
                    parts.append(f"{added} añadidas")
                if updated:
                    parts.append(f"{updated} actualizadas")
                if skipped:
                    parts.append(f"{skipped} omitidas (ya existían)")
                st.success(f"Asignación completada: {', '.join(parts)}.")
                st.rerun()

    st.divider()

    # ── BOM table ────────────────────────────────────────────────────────────
    st.subheader("BOM en construcción")

    if not st.session_state["bom_draft"]:
        st.info(
            "Aún no hay entradas. Usa el formulario de arriba para añadir componentes, "
            "o carga una BOM existente como base de partida."
        )
    else:
        draft_df = pd.DataFrame(st.session_state["bom_draft"])
        draft_df["_idx"] = range(len(draft_df))

        to_delete = None

        for bc_var, group in draft_df.groupby("Cod Barras Variante", sort=False):
            var_name = group.iloc[0]["_nombre_variante"]
            st.markdown(
                f'<div class="ref-card"><span class="ref-title">{var_name}</span></div>',
                unsafe_allow_html=True,
            )

            for _, row in group.iterrows():
                idx = int(row["_idx"])
                rc1, rc2, rc3, rc4 = st.columns([4, 2.5, 1.5, 0.6])
                rc1.write(row["_nombre_componente"])
                rc2.caption(f"EAN: {row['EAN Componente']}")
                _bom_qty_key = f"bom_qty_{row['Cod Barras Variante']}_{row['EAN Componente']}"
                rc3.number_input(
                    "Cantidad",
                    min_value=0.001,
                    step=0.1,
                    value=float(row["Cantidad"]),
                    key=_bom_qty_key,
                    label_visibility="collapsed",
                    format="%.3f",
                )
                if rc4.button("X", key=f"del_bom_{idx}", help="Eliminar esta entrada"):
                    to_delete = idx

        # Sync edited quantities back to bom_draft
        for _entry in st.session_state["bom_draft"]:
            _sk = f"bom_qty_{_entry['Cod Barras Variante']}_{_entry['EAN Componente']}"
            if _sk in st.session_state:
                _entry["Cantidad"] = float(st.session_state[_sk])

        if to_delete is not None:
            st.session_state["bom_draft"].pop(to_delete)
            st.rerun()

        n_prendas = draft_df["Cod Barras Variante"].nunique()
        n_entradas = len(draft_df)
        st.caption(f"{n_prendas} prendas — {n_entradas} entradas en la BOM")

        st.divider()

        # ── Actions ──────────────────────────────────────────────────────────
        st.subheader("Acciones")
        ac1, ac2, ac3 = st.columns(3)

        with ac1:
            if st.button("Limpiar todo", use_container_width=True):
                st.session_state["bom_draft"] = []
                if "custom_bom" in st.session_state:
                    del st.session_state["custom_bom"]
                st.rerun()

        with ac2:
            export_rows = [
                {
                    "Cod Barras Variante": e["Cod Barras Variante"],
                    "EAN Componente": e["EAN Componente"],
                    "Cantidad": e["Cantidad"],
                }
                for e in st.session_state["bom_draft"]
            ]
            export_df = pd.DataFrame(export_rows)
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                export_df.to_excel(writer, sheet_name="Fichero ejemplo", index=False)
            st.download_button(
                "Descargar BOM (Excel)",
                buf.getvalue(),
                "lista_materiales.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with ac3:
            if st.button(
                "Usar esta BOM para calcular",
                type="primary",
                use_container_width=True,
            ):
                custom_bom_df = pd.DataFrame([
                    {
                        "Cod Barras Variante": e["Cod Barras Variante"],
                        "EAN Componente": e["EAN Componente"],
                        "Cantidad": float(e["Cantidad"]),
                    }
                    for e in st.session_state["bom_draft"]
                ])
                st.session_state["custom_bom"] = custom_bom_df
                st.success(
                    "BOM personalizada activada. "
                    "Ve a la pestaña 'Plan de producción' y pulsa 'Calcular Consumos'."
                )
