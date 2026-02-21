import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import os
import json
import time
from pathlib import Path
import groq as _groq_module
from groq import Groq
import pdfplumber

# File used to persist the BOM draft across browser sessions
BOM_SESSION_FILE = Path(__file__).parent / "bom_session.json"


def _save_bom_draft():
    """Write bom_draft to disk so it survives page reloads."""
    try:
        data = [
            {**e, "Cantidad": float(e["Cantidad"])}
            for e in st.session_state.get("bom_draft", [])
        ]
        BOM_SESSION_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass

st.set_page_config(
    page_title="Planificación de Consumos",
    layout="wide",
    initial_sidebar_state="expanded",
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


# ── SIDEBAR — Asistente IA ────────────────────────────────────────────────────
with st.sidebar:
    st.header("Asistente IA")
    st.caption("Pregunta o da órdenes: plan de producción, BOM, consumos…")

    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []

    # ── Variant / component catalogs for BOM tools ────────────────────────────
    _sb_all = variantes.dropna(subset=["Código de barras principal"]).copy()
    _sb_is_fin = (
        _sb_all["Código de barras principal"].str.len().eq(13)
        & _sb_all["Código de barras principal"].str.startswith("8445790")
    )
    _sb_prendas = _sb_all[_sb_is_fin].copy()   # all finished-product variants
    _sb_comp    = _sb_all[~_sb_is_fin].copy()   # all component variants

    def _match_prendas(filtro: dict):
        """Return rows of _sb_prendas matching a garment filter dict."""
        mask = pd.Series([True] * len(_sb_prendas), index=_sb_prendas.index)
        if filtro.get("referencia"):
            mask &= _sb_prendas["Referencia interna"].str.contains(filtro["referencia"], case=False, na=False)
        if filtro.get("nombre"):
            mask &= _sb_prendas["Nombre"].str.contains(filtro["nombre"], case=False, na=False)
        if filtro.get("color"):
            mask &= _sb_prendas["Color"].str.contains(filtro["color"], case=False, na=False)
        if filtro.get("talla"):
            mask &= _sb_prendas["Talla"].str.upper() == str(filtro["talla"]).upper()
        return _sb_prendas[mask]

    def _match_comp(referencia_o_nombre: str):
        """Return component rows matching the given string in Referencia interna or Nombre."""
        m = (
            _sb_comp["Referencia interna"].str.contains(referencia_o_nombre, case=False, na=False)
            | _sb_comp["Nombre"].str.contains(referencia_o_nombre, case=False, na=False)
        )
        return _sb_comp[m]

    # ── Tool definitions ──────────────────────────────────────────────────────
    _ai_tools = [
        {
            "type": "function",
            "function": {
                "name": "set_production_quantities",
                "description": (
                    "Establece la cantidad a producir para una o varias variantes del plan "
                    "de producción. Filtra por referencia, nombre, color y/o talla. "
                    "Si no se especifica algún campo, no se filtra por él."
                ),
                "parameters": {
                    "type": "object",
                    "properties": {
                        "items": {
                            "type": "array",
                            "description": "Lista de asignaciones a realizar",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "referencia": {"type": "string", "description": "Referencia interna (parcial)"},
                                    "nombre": {"type": "string", "description": "Nombre del producto (parcial)"},
                                    "color": {"type": "string", "description": "Color (parcial)"},
                                    "talla": {"type": "string", "description": "Talla exacta"},
                                    "cantidad": {"type": "integer", "description": "Unidades a producir"},
                                },
                                "required": ["cantidad"],
                            },
                        }
                    },
                    "required": ["items"],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "clear_production_plan",
                "description": "Pone a cero todas las cantidades del plan de producción.",
                "parameters": {"type": "object", "properties": {}},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "add_bom_entries",
                "description": (
                    "Añade uno o varios componentes a la BOM en construcción para las prendas "
                    "que coincidan con el filtro. Si la entrada ya existe puede omitirse o sobreescribirse."
                ),
                "parameters": {
                    "type": "object",
                    "properties": {
                        "prendas": {
                            "type": "object",
                            "description": "Filtro para seleccionar prendas (todos los campos son opcionales)",
                            "properties": {
                                "referencia": {"type": "string"},
                                "nombre":     {"type": "string"},
                                "color":      {"type": "string"},
                                "talla":      {"type": "string"},
                            },
                        },
                        "componentes": {
                            "type": "array",
                            "description": "Componentes a añadir",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "referencia_o_nombre": {
                                        "type": "string",
                                        "description": "Referencia interna o nombre del componente",
                                    },
                                    "cantidad":      {"type": "number"},
                                    "sobreescribir": {"type": "boolean"},
                                },
                                "required": ["referencia_o_nombre", "cantidad"],
                            },
                        },
                    },
                    "required": ["prendas", "componentes"],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "remove_bom_entries",
                "description": (
                    "Elimina entradas de la BOM en construcción para las prendas que "
                    "coincidan con el filtro. Si se indica un componente solo se elimina ese; "
                    "si no se indica ninguno se eliminan todos los componentes de las prendas filtradas."
                ),
                "parameters": {
                    "type": "object",
                    "properties": {
                        "prendas": {
                            "type": "object",
                            "description": "Filtro para seleccionar prendas",
                            "properties": {
                                "referencia": {"type": "string"},
                                "nombre":     {"type": "string"},
                                "color":      {"type": "string"},
                                "talla":      {"type": "string"},
                            },
                        },
                        "componente": {
                            "type": "string",
                            "description": "Nombre o referencia del componente a eliminar (opcional)",
                        },
                    },
                    "required": ["prendas"],
                },
            },
        },
    ]

    def _exec_set_quantities(items):
        """Apply set_production_quantities tool call, return human-readable result."""
        lines = []
        for item in items:
            mask = pd.Series([True] * len(finished), index=finished.index)
            if item.get("referencia"):
                mask &= finished["Referencia interna"].str.contains(
                    item["referencia"], case=False, na=False
                )
            if item.get("nombre"):
                mask &= finished["Nombre"].str.contains(
                    item["nombre"], case=False, na=False
                )
            if item.get("color"):
                mask &= finished["Color"].str.contains(
                    item["color"], case=False, na=False
                )
            if item.get("talla"):
                mask &= finished["Talla"].str.upper() == str(item["talla"]).upper()
            matches = finished[mask]
            if matches.empty:
                lines.append(f"Sin coincidencias para: {item}")
            else:
                for _, row in matches.iterrows():
                    bc = row["Código de barras principal"]
                    st.session_state[f"qty_{bc}"] = int(item["cantidad"])
                    lines.append(
                        f"✓ {row['Referencia interna']} | {row['Nombre']} | "
                        f"{row['Color']} | {row['Talla']} → {item['cantidad']} ud."
                    )
        return "\n".join(lines) if lines else "No se realizaron cambios."

    def _exec_clear_plan():
        for _k in list(st.session_state.keys()):
            if _k.startswith("qty_"):
                del st.session_state[_k]
        st.session_state.pop("results_df", None)
        st.session_state.pop("plan_df", None)
        return "Plan de producción vaciado."

    def _exec_add_bom_entries(prendas: dict, componentes: list):
        lines = []
        prenda_matches = _match_prendas(prendas)
        if prenda_matches.empty:
            return f"Sin prendas coincidentes con el filtro: {prendas}"
        for comp_spec in componentes:
            ron = comp_spec.get("referencia_o_nombre", "")
            cantidad = float(comp_spec.get("cantidad", 1))
            sobreescribir = bool(comp_spec.get("sobreescribir", False))
            comp_matches = _match_comp(ron)
            if comp_matches.empty:
                lines.append(f"⚠ Componente no encontrado: '{ron}'")
                continue
            # Use the first (best) match
            comp_row = comp_matches.iloc[0]
            comp_bc  = comp_row["Código de barras principal"]
            comp_label = f"{comp_row['Referencia interna']} | {comp_row['Nombre']}"
            for _, p_row in prenda_matches.iterrows():
                p_bc = p_row["Código de barras principal"]
                existing = next(
                    (i for i, e in enumerate(st.session_state["bom_draft"])
                     if e["Cod Barras Variante"] == p_bc and e["EAN Componente"] == comp_bc),
                    None,
                )
                p_label = f"{p_row['Referencia interna']} | {p_row['Color']} | {p_row['Talla']}"
                if existing is not None:
                    if sobreescribir:
                        st.session_state["bom_draft"][existing]["Cantidad"] = cantidad
                        lines.append(f"↺ {p_label} ← {comp_label}: {cantidad} (actualizado)")
                    else:
                        lines.append(f"· {p_label} ← {comp_label}: ya existía, omitido")
                else:
                    st.session_state["bom_draft"].append({
                        "Cod Barras Variante":  p_bc,
                        "EAN Componente":       comp_bc,
                        "Cantidad":             cantidad,
                        "_nombre_variante":     f"{p_row['Referencia interna']} | {p_row['Nombre']} | {p_row['Color']} | {p_row['Talla']}",
                        "_nombre_componente":   f"{comp_row['Referencia interna']} | {comp_row['Nombre']} | {comp_row['Color']} | {comp_row['Talla']}",
                    })
                    lines.append(f"✓ {p_label} ← {comp_label}: {cantidad}")
        return "\n".join(lines) if lines else "No se realizaron cambios."

    def _exec_remove_bom_entries(prendas: dict, componente: str = ""):
        prenda_matches = _match_prendas(prendas)
        if prenda_matches.empty:
            return f"Sin prendas coincidentes con el filtro: {prendas}"
        p_bcs = set(prenda_matches["Código de barras principal"].tolist())
        comp_bcs: set = set()
        if componente:
            cm = _match_comp(componente)
            if cm.empty:
                return f"Componente no encontrado: '{componente}'"
            comp_bcs = set(cm["Código de barras principal"].tolist())
        before = len(st.session_state["bom_draft"])
        st.session_state["bom_draft"] = [
            e for e in st.session_state["bom_draft"]
            if not (
                e["Cod Barras Variante"] in p_bcs
                and (not comp_bcs or e["EAN Componente"] in comp_bcs)
            )
        ]
        removed = before - len(st.session_state["bom_draft"])
        return f"Eliminadas {removed} entradas de la BOM en construcción."

    # ── Chat display ──────────────────────────────────────────────────────────
    for msg in st.session_state["chat_history"]:
        if msg["role"] in ("user", "assistant"):
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    # ── User input ────────────────────────────────────────────────────────────
    user_input = st.chat_input("Escribe tu pregunta u orden…")

    if user_input:
        st.session_state["chat_history"].append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)

        # ── Build context ─────────────────────────────────────────────────────
        _ctx_parts = []

        _active_bom_ctx = st.session_state.get("custom_bom", bom)
        _ctx_parts.append(
            f"BOM activa: {len(_active_bom_ctx)} entradas, "
            f"{_active_bom_ctx['Cod Barras Variante'].nunique()} variantes, "
            f"{_active_bom_ctx['EAN Componente'].nunique()} componentes distintos."
        )

        _draft = st.session_state.get("bom_draft", [])
        if _draft:
            _draft_df_ctx = pd.DataFrame(_draft)
            _ctx_parts.append(
                f"BOM en construcción: {len(_draft)} entradas, "
                f"{_draft_df_ctx['Cod Barras Variante'].nunique()} prendas."
            )

        _plan_qtys = {
            k[4:]: int(v)
            for k, v in st.session_state.items()
            if k.startswith("qty_") and isinstance(v, (int, float)) and v > 0
        }
        if _plan_qtys:
            _ctx_parts.append(
                f"Plan de producción: {len(_plan_qtys)} variantes, "
                f"{sum(_plan_qtys.values())} unidades totales."
            )
        else:
            _ctx_parts.append("Plan de producción: vacío.")

        if "results_df" in st.session_state:
            _res_ctx = st.session_state["results_df"]
            _top5 = _res_ctx.nlargest(5, "Cantidad necesaria")[["Nombre", "Cantidad necesaria"]].to_dict("records")
            _top5_str = "; ".join(f"{r['Nombre']} ({r['Cantidad necesaria']})" for r in _top5)
            _ctx_parts.append(
                f"Consumos calculados: {len(_res_ctx)} componentes. Top 5: {_top5_str}."
            )

        # Catalog summary for the model to match user intent
        _refs_list   = sorted(_sb_prendas["Referencia interna"].dropna().unique().tolist())
        _colors_list = sorted(_sb_prendas["Color"].dropna().unique().tolist())
        _tallas_list = sorted(_sb_prendas["Talla"].dropna().unique().tolist(), key=_sz_key)
        _comp_labels = sorted(
            (_sb_comp["Referencia interna"].fillna("") + " | " + _sb_comp["Nombre"].fillna(""))
            .str.strip(" |").unique().tolist()
        )
        _ctx_parts.append(f"Referencias de prenda disponibles: {', '.join(_refs_list[:40])}{'…' if len(_refs_list)>40 else ''}.")
        _ctx_parts.append(f"Colores disponibles: {', '.join(_colors_list[:30])}{'…' if len(_colors_list)>30 else ''}.")
        _ctx_parts.append(f"Tallas disponibles: {', '.join(_tallas_list)}.")
        _ctx_parts.append(f"Componentes del catálogo: {', '.join(_comp_labels[:50])}{'…' if len(_comp_labels)>50 else ''}.")

        # BOM draft detail
        _draft_detail = st.session_state.get("bom_draft", [])
        if _draft_detail:
            _draft_by_prenda = {}
            for _e in _draft_detail:
                _draft_by_prenda.setdefault(_e["_nombre_variante"], []).append(_e["_nombre_componente"])
            _draft_summary = "; ".join(
                f"{k[:30]} ({len(v)} comp.)" for k, v in list(_draft_by_prenda.items())[:10]
            )
            _ctx_parts.append(f"BOM en construcción detalle: {_draft_summary}{'…' if len(_draft_by_prenda)>10 else ''}.")

        _context_block = "\n".join(f"- {p}" for p in _ctx_parts)

        _system_prompt = f"""Eres un asistente especializado en planificación de producción \
y gestión de materiales para una empresa de moda. Puedes responder preguntas Y ejecutar \
acciones sobre la aplicación usando las funciones disponibles.

Estado actual de la aplicación:
{_context_block}

Responde siempre en español, de forma concisa.
- Para establecer cantidades de producción → set_production_quantities
- Para vaciar el plan → clear_production_plan
- Para añadir componentes a la BOM en construcción → add_bom_entries
- Para eliminar componentes de la BOM en construcción → remove_bom_entries
- Para preguntas informativas → responde directamente sin funciones"""

        _messages_api = [{"role": "system", "content": _system_prompt}]
        _messages_api += [
            {"role": m["role"], "content": m["content"]}
            for m in st.session_state["chat_history"]
            if m["role"] in ("user", "assistant")
        ]

        try:
            _groq_client = Groq(
                api_key=st.secrets.get("GROQ_API_KEY", os.environ.get("GROQ_API_KEY", "")),
                timeout=30.0,
            )

            # First call — model may request a tool
            _resp1 = _groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=_messages_api,
                tools=_ai_tools,
                tool_choice="auto",
                max_tokens=512,
                temperature=0.2,
            )
            _msg1 = _resp1.choices[0].message

            if _msg1.tool_calls:
                # Execute each tool call
                _tool_results = []
                for _tc in _msg1.tool_calls:
                    _args = json.loads(_tc.function.arguments)
                    if _tc.function.name == "set_production_quantities":
                        _tr = _exec_set_quantities(_args.get("items", []))
                    elif _tc.function.name == "clear_production_plan":
                        _tr = _exec_clear_plan()
                    elif _tc.function.name == "add_bom_entries":
                        _tr = _exec_add_bom_entries(
                            _args.get("prendas", {}),
                            _args.get("componentes", []),
                        )
                    elif _tc.function.name == "remove_bom_entries":
                        _tr = _exec_remove_bom_entries(
                            _args.get("prendas", {}),
                            _args.get("componente", ""),
                        )
                    else:
                        _tr = "Función desconocida."
                    _tool_results.append({"tool_call_id": _tc.id, "result": _tr})

                # Build tool result messages
                _messages_api.append(_msg1)
                for _tr_item in _tool_results:
                    _messages_api.append({
                        "role": "tool",
                        "tool_call_id": _tr_item["tool_call_id"],
                        "content": _tr_item["result"],
                    })

                # Second call — model summarises what it did
                _resp2 = _groq_client.chat.completions.create(
                    model="llama-3.3-70b-versatile",
                    messages=_messages_api,
                    max_tokens=256,
                    temperature=0.2,
                )
                _answer = _resp2.choices[0].message.content
                # Store tool results in history so context is preserved
                st.session_state["chat_history"].append({
                    "role": "tool_summary",
                    "content": "\n".join(r["result"] for r in _tool_results),
                })
            else:
                _answer = _msg1.content or "Sin respuesta."

        except Exception as _e:
            _answer = f"Error al conectar con el asistente: {_e}"

        st.session_state["chat_history"].append({"role": "assistant", "content": _answer})
        with st.chat_message("assistant"):
            st.markdown(_answer)

        # Rerun so the main area (matrix, plan) reflects any quantity changes
        if any(m.get("role") == "tool_summary" for m in st.session_state["chat_history"][-3:]):
            st.rerun()

    if st.session_state["chat_history"] and st.button(
        "Limpiar conversación", use_container_width=True
    ):
        st.session_state["chat_history"] = []
        st.rerun()


# ── TABS ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["Plan de producción", "Crear / Editar BOM", "Simulador de escenarios"])


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
        _fresh_with_bom = sum(1 for bc in fresh_qtys if bc in _bom_covered)
        _bom_ok = _fresh_with_bom == len(fresh_qtys)
        _bom_status = (
            f" · BOM: **{_fresh_with_bom}/{len(fresh_qtys)}** variantes cubiertas"
            + ("" if _bom_ok else " ⚠️")
        )
        summary_slot.info(
            f"**{len(fresh_qtys)}** variantes con cantidad asignada — "
            f"**{sum(fresh_qtys.values())}** unidades totales{_bom_status}"
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

        # ── Consumos por colección ────────────────────────────────────────────
        with st.expander("Consumos por colección", expanded=False):
            _active_bom_coll = st.session_state.get("custom_bom", bom)
            _coll_map: dict[tuple, float] = {}
            for _, _pr in plan_df.iterrows():
                _bc = _pr["Código de barras"]
                _qty = int(_pr["Unidades"])
                _coll = str(_pr.get("Referencia", "")).strip()[:3] or "—"
                for _, _br in _active_bom_coll[
                    _active_bom_coll["Cod Barras Variante"] == _bc
                ].iterrows():
                    _comp = str(_br["EAN Componente"])
                    _cname = barcode_lookup.get(_comp, {}).get("Nombre", _comp)
                    _key = (_coll, _comp, _cname)
                    _coll_map[_key] = _coll_map.get(_key, 0.0) + float(_br["Cantidad"]) * _qty
            if _coll_map:
                _coll_df = pd.DataFrame([
                    {
                        "Colección": k[0],
                        "Componente": k[2],
                        "EAN": k[1],
                        "Cantidad total": round(v, 4),
                    }
                    for k, v in _coll_map.items()
                ]).sort_values(["Colección", "Componente"]).reset_index(drop=True)
                st.dataframe(_coll_df, hide_index=True, use_container_width=True)
                _coll_buf = BytesIO()
                with pd.ExcelWriter(_coll_buf, engine="openpyxl") as _w:
                    _coll_df.to_excel(_w, index=False, sheet_name="Por colección")
                st.download_button(
                    "Descargar por colección (Excel)",
                    _coll_buf.getvalue(),
                    "consumos_por_coleccion.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.info("No hay datos de colección disponibles (comprueba que la BOM esté activa).")

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

    # Initialize BOM draft — load from disk if first run of this session
    if "bom_draft" not in st.session_state:
        if BOM_SESSION_FILE.exists():
            try:
                st.session_state["bom_draft"] = json.loads(
                    BOM_SESSION_FILE.read_text("utf-8")
                )
                if st.session_state["bom_draft"]:
                    st.info(
                        f"Se ha restaurado la BOM de la sesión anterior "
                        f"({len(st.session_state['bom_draft'])} entradas).",
                        icon="💾",
                    )
            except Exception:
                st.session_state["bom_draft"] = []
        else:
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
    # 4th character of Referencia interna identifies garment type
    all_variants["_tipo_prenda"] = (
        all_variants["Referencia interna"].fillna("").str[3:4].replace("", "—")
    )
    variant_label_list = sorted(all_variants["_label"].unique().tolist())
    variant_label_to_bc = dict(
        zip(all_variants["_label"], all_variants["Código de barras principal"])
    )

    _bc_col = "Código de barras principal"
    _is_finished = (
        all_variants[_bc_col].str.len().eq(13)
        & all_variants[_bc_col].str.startswith("8445790")
    )

    # Finished products only (prenda selectors)
    prenda_variants = all_variants[_is_finished].copy()
    prenda_label_list = sorted(prenda_variants["_label"].unique().tolist())
    prenda_label_to_bc = dict(zip(prenda_variants["_label"], prenda_variants[_bc_col]))
    prenda_label_to_tipo = dict(
        zip(prenda_variants["_label"], prenda_variants["_tipo_prenda"])
    )

    # Components: everything that is NOT a finished product
    comp_variants = all_variants[~_is_finished].copy()
    comp_label_list = sorted(comp_variants["_label"].unique().tolist())
    comp_label_to_bc = dict(zip(comp_variants["_label"], comp_variants[_bc_col]))

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

    # Row 1: selectors (full width so dropdowns aren't clipped)
    st.markdown("**Prenda (variante de producto)**")
    sel_var_label = st.selectbox(
        "Variante",
        options=prenda_label_list,
        label_visibility="collapsed",
        key="bom_sel_variant",
    )
    sel_var_bc = prenda_label_to_bc.get(sel_var_label, "")

    st.markdown("**Componente / Material**")
    _bom_cq = st.text_input(
        "Buscar componente por referencia o nombre",
        placeholder="ej. Tejido, entretela, forro…",
        key="bom_comp_q",
        label_visibility="collapsed",
    )
    _bom_cq_strip = _bom_cq.strip()
    if _bom_cq_strip:
        _bom_comp_filtered = comp_variants[
            comp_variants["Referencia interna"].str.contains(_bom_cq_strip, case=False, na=False)
            | comp_variants["Nombre"].str.contains(_bom_cq_strip, case=False, na=False)
        ]
    else:
        _bom_comp_filtered = comp_variants
    _bom_comp_opts = sorted(_bom_comp_filtered["_label"].unique().tolist())
    if _bom_comp_opts:
        sel_comp_label = st.selectbox(
            "Componente",
            options=_bom_comp_opts,
            label_visibility="collapsed",
            key="bom_sel_comp",
        )
        comp_ean = comp_label_to_bc.get(sel_comp_label, "")
        comp_display_name = sel_comp_label
    else:
        st.warning("Sin resultados. Prueba con otro término.")
        comp_ean = ""
        comp_display_name = ""
    if st.checkbox("No está en el catálogo (introducir EAN)", key="bom_comp_manual"):
        _bom_ean_c1, _bom_ean_c2 = st.columns(2)
        with _bom_ean_c1:
            comp_ean = st.text_input(
                "EAN del componente",
                key="bom_comp_ean",
                placeholder="ej. 8412345678901",
            )
        with _bom_ean_c2:
            comp_display_name = st.text_input(
                "Nombre del componente",
                key="bom_comp_name",
                placeholder="ej. Tejido principal",
            )

    # Row 2: quantity + action (narrow controls)
    fc_qty, fc_btn = st.columns([2, 5])
    with fc_qty:
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
    with fc_btn:
        st.write("")
        st.write("")
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
                _pfx_preview = prenda_variants[
                    prenda_variants["Referencia interna"]
                    .astype(str)
                    .str.startswith(bulk_ean_prefix.strip())
                ]
                st.caption(
                    f"{len(_pfx_preview)} variantes cuya referencia comienza por **{bulk_ean_prefix.strip()}**"
                )

        st.markdown("— o filtra por atributos —")
        bf1, bf2, bf3 = st.columns(3)
        with bf1:
            bulk_refs = st.multiselect(
                "Referencia interna",
                sorted(prenda_variants["Referencia interna"].dropna().unique().tolist()),
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
                sorted(prenda_variants["Color"].dropna().unique().tolist()),
                placeholder="Todos los colores",
                key="bulk_colors",
            )
            bulk_tallas = st.multiselect(
                "Talla",
                sorted(prenda_variants["Talla"].dropna().unique().tolist(), key=_sz_key),
                placeholder="Todas las tallas",
                key="bulk_tallas",
            )
        with bf3:
            _tipos_disp = sorted(
                [
                    t for t in prenda_variants["_tipo_prenda"].unique()
                    if t and t != "—"
                ],
                key=lambda x: (len(x), x),
            )
            bulk_tipos = st.multiselect(
                "Tipo de prenda (4º dígito de ref.)",
                _tipos_disp,
                placeholder="Todos los tipos",
                key="bulk_tipos",
                help=(
                    "El 4º carácter de la referencia interna identifica el tipo de prenda. "
                    "Ej.: ref. '2611x' → tipo '1'."
                ),
            )

        # Apply filters — base is finished products only
        bulk_target = prenda_variants.copy()
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
        if bulk_tipos:
            bulk_target = bulk_target[bulk_target["_tipo_prenda"].isin(bulk_tipos)]

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

        # Selector a ancho completo para que el desplegable no se corte
        _bulk_cq = st.text_input(
            "Buscar componente por referencia o nombre",
            placeholder="ej. Tejido, entretela, forro…",
            key="bulk_comp_q",
            label_visibility="collapsed",
        )
        _bulk_cq_strip = _bulk_cq.strip()
        if _bulk_cq_strip:
            _bulk_comp_filtered = comp_variants[
                comp_variants["Referencia interna"].str.contains(_bulk_cq_strip, case=False, na=False)
                | comp_variants["Nombre"].str.contains(_bulk_cq_strip, case=False, na=False)
            ]
        else:
            _bulk_comp_filtered = comp_variants
        _bulk_comp_opts = sorted(_bulk_comp_filtered["_label"].unique().tolist())
        if _bulk_comp_opts:
            bulk_sel_comp = st.selectbox(
                "Componente",
                options=_bulk_comp_opts,
                label_visibility="collapsed",
                key="bulk_sel_comp",
            )
            bulk_comp_ean = comp_label_to_bc.get(bulk_sel_comp, "")
            bulk_comp_name = bulk_sel_comp
        else:
            st.warning("Sin resultados.")
            bulk_comp_ean = ""
            bulk_comp_name = ""
        if st.checkbox("No está en el catálogo (introducir EAN)", key="bulk_comp_manual"):
            _bulk_ean_c1, _bulk_ean_c2 = st.columns(2)
            with _bulk_ean_c1:
                bulk_comp_ean = st.text_input(
                    "EAN del componente",
                    placeholder="ej. 8412345678901",
                    key="bulk_comp_ean",
                )
            with _bulk_ean_c2:
                bulk_comp_name = st.text_input(
                    "Nombre del componente",
                    placeholder="ej. Tejido principal",
                    key="bulk_comp_name",
                )

        # Controles pequeños en fila inferior
        bc2, bc3, bc4 = st.columns([2, 2, 3])

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

    # ── Import BOM from PDF technical sheet ───────────────────────────────────
    with st.expander("Importar BOM desde ficha técnica (IA)", expanded=False):
        st.markdown(
            "Selecciona una **referencia** (sin color ni talla) y sube su ficha técnica estándar. "
            "La IA extraerá los materiales y asignará automáticamente el componente correcto "
            "a cada variante de color y talla."
        )

        # Step 1 — select reference (no color/size)
        st.markdown("**1. Referencia del producto**")
        _pdf_ref_opts = (
            prenda_variants[["Referencia interna", "Nombre"]]
            .drop_duplicates()
            .sort_values("Referencia interna")
        )
        _pdf_ref_opts["_rl"] = (
            _pdf_ref_opts["Referencia interna"].str.strip()
            + " | "
            + _pdf_ref_opts["Nombre"].str.strip()
        )
        _pdf_ref_label_list = _pdf_ref_opts["_rl"].tolist()

        pdf_sel_ref = st.selectbox(
            "Referencia",
            options=_pdf_ref_label_list,
            placeholder="Selecciona una referencia…",
            key="pdf_sel_ref",
            label_visibility="collapsed",
        )

        # Step 2 — upload PDF
        st.markdown("**2. Sube la ficha técnica estándar (PDF)**")
        pdf_file = st.file_uploader(
            "Ficha técnica PDF",
            type=["pdf"],
            key="pdf_upload",
            label_visibility="collapsed",
        )

        if pdf_file and pdf_sel_ref:
            _sel_ref_id = pdf_sel_ref.split(" | ")[0].strip()
            _ref_vars = prenda_variants[
                prenda_variants["Referencia interna"] == _sel_ref_id
            ].copy()
            _ref_colors = sorted(_ref_vars["Color"].unique().tolist())
            _ref_tallas = sorted(_ref_vars["Talla"].unique().tolist())
            st.caption(
                f"Se generará BOM para {len(_ref_vars)} variantes: "
                f"{len(_ref_colors)} color(es) × {len(_ref_tallas)} talla(s)"
            )

            if st.button(
                "Analizar ficha técnica",
                type="primary",
                use_container_width=True,
                key="pdf_analyze",
            ):
                # Extract text and tables from PDF
                with st.spinner("Extrayendo contenido del PDF…"):
                    try:
                        _text_parts = []
                        with pdfplumber.open(pdf_file) as _pdf:
                            for _page in _pdf.pages:
                                _t = _page.extract_text()
                                if _t:
                                    _text_parts.append(_t)
                                for _tbl in (_page.extract_tables() or []):
                                    for _row in _tbl:
                                        if _row:
                                            _text_parts.append(
                                                " | ".join(str(c) for c in _row if c)
                                            )
                        _pdf_text = "\n".join(_text_parts)[:8000]
                    except Exception as _pe:
                        _pdf_text = ""
                        st.error(f"Error al leer el PDF: {_pe}")

                if _pdf_text:
                    with st.spinner("La IA está analizando los materiales…"):
                        _pdf_comps = []
                        _groq_api_key = st.secrets.get(
                            "GROQ_API_KEY", os.environ.get("GROQ_API_KEY", "")
                        )
                        if not _groq_api_key:
                            st.error(
                                "No se encontró GROQ_API_KEY. Configúrala en "
                                ".streamlit/secrets.toml o en las variables de entorno."
                            )
                        else:
                            for _attempt in range(3):
                                try:
                                    _pdf_groq = Groq(
                                        api_key=_groq_api_key,
                                        timeout=30.0,
                                    )
                                    _pdf_resp = _pdf_groq.chat.completions.create(
                                        model="llama-3.3-70b-versatile",
                                        messages=[
                                            {
                                                "role": "system",
                                                "content": (
                                                    "Eres un experto en fichas técnicas de prendas de moda. "
                                                    "Extrae TODOS los materiales, tejidos, forros, entretelas, "
                                                    "accesorios y componentes. Para cada componente detecta su "
                                                    "color si está indicado en la ficha. "
                                                    'Devuelve ÚNICAMENTE un JSON con la clave "componentes", '
                                                    "array de objetos con los campos: "
                                                    '"nombre" (string, tipo de material sin incluir el color), '
                                                    '"color" (string, color del componente o "" si no se especifica), '
                                                    '"cantidad" (número), '
                                                    '"unidad" (string: m, ud, kg, cm, etc.). '
                                                    "Si no hay cantidad usa 1. Si no hay unidad usa 'ud'."
                                                ),
                                            },
                                            {
                                                "role": "user",
                                                "content": f"Ficha técnica:\n\n{_pdf_text}",
                                            },
                                        ],
                                        response_format={"type": "json_object"},
                                        max_tokens=1024,
                                        temperature=0,
                                    )
                                    _pdf_comps = json.loads(
                                        _pdf_resp.choices[0].message.content
                                    ).get("componentes", [])
                                    break  # success — exit retry loop
                                except (
                                    _groq_module.APIConnectionError,
                                    _groq_module.APITimeoutError,
                                ) as _ae:
                                    if _attempt < 2:
                                        time.sleep(2 ** _attempt)
                                        continue
                                    st.error(
                                        f"Error de conexión con la IA tras 3 intentos: {_ae}. "
                                        "Comprueba la conectividad de red o el estado de la API de Groq "
                                        "(https://status.groq.com)."
                                    )
                                except _groq_module.AuthenticationError:
                                    st.error(
                                        "La clave GROQ_API_KEY no es válida. "
                                        "Revísala en .streamlit/secrets.toml."
                                    )
                                    break
                                except Exception as _ae:
                                    st.error(f"Error en el análisis IA: {_ae}")
                                    break

                    if _pdf_comps:
                        # Smart color-aware matching: for each finished variant, find
                        # the catalog component that best matches the variant's color.
                        _pdf_var_results = []
                        for _, _vrow in _ref_vars.iterrows():
                            _v_color = _vrow["Color"]
                            _v_talla = _vrow["Talla"]
                            _v_bc    = _vrow["Código de barras principal"]
                            _v_label = _vrow["_label"]
                            _v_comps = []
                            for _pc in _pdf_comps:
                                _nombre    = str(_pc.get("nombre", "")).strip()
                                _pc_color  = str(_pc.get("color", "")).strip()
                                _cantidad  = float(_pc.get("cantidad", 1))
                                _unidad    = str(_pc.get("unidad", "ud"))
                                _cm_all    = _match_comp(_nombre)
                                _comp_match = None
                                if not _cm_all.empty:
                                    # Priority 1: catalog component whose color matches
                                    # the finished product variant color
                                    _cm_vc = _cm_all[
                                        _cm_all["Color"].str.contains(
                                            _v_color, case=False, na=False
                                        )
                                    ]
                                    if not _cm_vc.empty:
                                        _comp_match = _cm_vc.iloc[0]
                                    else:
                                        # Priority 2: catalog component whose color
                                        # matches the color the AI detected in the sheet
                                        if _pc_color:
                                            _cm_pc = _cm_all[
                                                _cm_all["Color"].str.contains(
                                                    _pc_color, case=False, na=False
                                                )
                                            ]
                                            if not _cm_pc.empty:
                                                _comp_match = _cm_pc.iloc[0]
                                        # Priority 3: UNICA / talla única
                                        if _comp_match is None:
                                            _cm_unica = _cm_all[
                                                _cm_all["Talla"].str.upper().isin(
                                                    ["UNICA", "ÚNICA", "UNICO", "U"]
                                                )
                                            ]
                                            _comp_match = (
                                                _cm_unica.iloc[0]
                                                if not _cm_unica.empty
                                                else _cm_all.iloc[0]
                                            )
                                if _comp_match is not None:
                                    _v_comps.append({
                                        "nombre_extraido": _nombre,
                                        "color_extraido":  _pc_color,
                                        "cantidad":        _cantidad,
                                        "unidad":          _unidad,
                                        "comp_bc":         _comp_match["Código de barras principal"],
                                        "comp_ref":        _comp_match["Referencia interna"],
                                        "comp_nombre":     _comp_match["Nombre"],
                                        "comp_color":      _comp_match["Color"],
                                        "comp_talla":      _comp_match["Talla"],
                                        "status":          "✓ Encontrado",
                                    })
                                else:
                                    _v_comps.append({
                                        "nombre_extraido": _nombre,
                                        "color_extraido":  _pc_color,
                                        "cantidad":        _cantidad,
                                        "unidad":          _unidad,
                                        "comp_bc":         None,
                                        "comp_ref":        None,
                                        "comp_nombre":     None,
                                        "comp_color":      None,
                                        "comp_talla":      None,
                                        "status":          "⚠ Sin coincidencia",
                                    })
                            _pdf_var_results.append({
                                "label":       _v_label,
                                "bc":          _v_bc,
                                "color":       _v_color,
                                "talla":       _v_talla,
                                "componentes": _v_comps,
                            })
                        st.session_state["pdf_analysis"] = {
                            "ref_label": pdf_sel_ref,
                            "variantes": _pdf_var_results,
                        }
                    else:
                        st.warning("No se identificaron componentes en la ficha técnica.")
                elif not _pdf_text:
                    st.warning(
                        "No se pudo extraer texto del PDF. "
                        "Comprueba que no esté protegido o sea solo imágenes."
                    )

        # Step 3 — review and confirm
        _pa = st.session_state.get("pdf_analysis")
        if _pa and _pa.get("ref_label") == pdf_sel_ref:
            _vars = _pa["variantes"]
            _total_ok  = sum(1 for v in _vars for c in v["componentes"] if c["comp_bc"])
            _total_nok = sum(1 for v in _vars for c in v["componentes"] if not c["comp_bc"])

            st.markdown("**3. Revisa la asignación por variante y confirma**")
            st.caption(
                f"{_total_ok} asignaciones correctas"
                + (f" — {_total_nok} sin coincidencia" if _total_nok else "")
            )

            # Flat display table: one row per (variant × component)
            _display_rows = []
            for _v in _vars:
                for _c in _v["componentes"]:
                    _display_rows.append({
                        "Variante":            f"{_v['color']} / {_v['talla']}",
                        "Material extraído":   _c["nombre_extraido"],
                        "Cantidad":            _c["cantidad"],
                        "Ud.":                 _c["unidad"],
                        "Componente asignado": (
                            f"{_c['comp_ref']} | {_c['comp_nombre']} | {_c['comp_color']}"
                            if _c["comp_bc"]
                            else "— Sin coincidencia"
                        ),
                        "Estado": _c["status"],
                    })
            st.dataframe(
                pd.DataFrame(_display_rows),
                hide_index=True,
                use_container_width=True,
            )

            if _total_nok:
                st.info(
                    "Los materiales sin coincidencia no se añadirán. "
                    "Puedes asignarlos manualmente desde el formulario."
                )

            if _total_ok:
                if st.button(
                    f"Añadir {_total_ok} asignaciones a la BOM",
                    type="primary",
                    use_container_width=True,
                    key="pdf_add_bom",
                ):
                    _added = 0
                    for _v in _vars:
                        for _c in _v["componentes"]:
                            if not _c["comp_bc"]:
                                continue
                            _exists = any(
                                e["Cod Barras Variante"] == _v["bc"]
                                and e["EAN Componente"] == _c["comp_bc"]
                                for e in st.session_state["bom_draft"]
                            )
                            if not _exists:
                                st.session_state["bom_draft"].append({
                                    "Cod Barras Variante": _v["bc"],
                                    "EAN Componente":      _c["comp_bc"],
                                    "Cantidad":            _c["cantidad"],
                                    "_nombre_variante":    _v["label"],
                                    "_nombre_componente":  (
                                        f"{_c['comp_ref']} | {_c['comp_nombre']} | {_c['comp_color']}"
                                    ),
                                })
                                _added += 1
                    del st.session_state["pdf_analysis"]
                    st.success(f"Añadidas {_added} entradas a la BOM en construcción.")
                    st.rerun()

    # ── Copy BOM from one variant to others ───────────────────────────────────
    with st.expander("Copiar BOM a otras variantes", expanded=False):
        st.markdown(
            "Copia todos los componentes de una variante origen a una o más variantes destino. "
            "Útil cuando varias tallas o colores comparten la misma estructura de materiales."
        )
        _cp1, _cp2 = st.columns(2)

        with _cp1:
            st.markdown("**Variante origen**")
            copy_src_label = st.selectbox(
                "Origen",
                prenda_label_list,
                label_visibility="collapsed",
                key="copy_src",
            )
            copy_src_bc = prenda_label_to_bc.get(copy_src_label, "")
            _src_entries = [
                e for e in st.session_state["bom_draft"]
                if e["Cod Barras Variante"] == copy_src_bc
            ]
            if _src_entries:
                st.caption(f"{len(_src_entries)} componente(s) en la BOM de origen")
                for _se in _src_entries:
                    st.caption(f"  · {_se['_nombre_componente']} — x{_se['Cantidad']:.3g}")
            else:
                st.warning("Esta variante no tiene BOM definida aún.")

        with _cp2:
            st.markdown("**Variantes destino**")
            _dst_tipo_opts = sorted(
                [t for t in set(prenda_label_to_tipo.values()) if t and t != "—"],
                key=lambda x: (len(x), x),
            )
            _filter_dst_tipos = st.multiselect(
                "Filtrar destinos por tipo de prenda (4º dígito ref.)",
                _dst_tipo_opts,
                placeholder="Todos los tipos",
                key="copy_dst_tipos",
                help=(
                    "Reduce la lista de destinos mostrando solo variantes cuya "
                    "referencia tenga ese 4º dígito. Ej.: '1' → tipo 1."
                ),
            )
            _filter_dst_color = st.multiselect(
                "Filtrar destinos por color",
                sorted(prenda_variants["Color"].dropna().unique().tolist()),
                placeholder="Todos los colores",
                key="copy_dst_color",
            )
            _filter_dst_talla = st.multiselect(
                "Filtrar destinos por talla",
                sorted(prenda_variants["Talla"].dropna().unique().tolist(), key=_sz_key),
                placeholder="Todas las tallas",
                key="copy_dst_talla",
            )
            # Build filtered destination option list (finished products only)
            _dst_options = [l for l in prenda_label_list if l != copy_src_label]
            if _filter_dst_tipos:
                _dst_options = [
                    l for l in _dst_options
                    if prenda_label_to_tipo.get(l, "—") in _filter_dst_tipos
                ]
            if _filter_dst_color:
                _color_map = dict(zip(prenda_variants["_label"], prenda_variants["Color"]))
                _dst_options = [
                    l for l in _dst_options
                    if _color_map.get(l, "") in _filter_dst_color
                ]
            if _filter_dst_talla:
                _talla_map = dict(zip(prenda_variants["_label"], prenda_variants["Talla"]))
                _dst_options = [
                    l for l in _dst_options
                    if _talla_map.get(l, "") in _filter_dst_talla
                ]
            copy_dst_labels = st.multiselect(
                "Destino",
                _dst_options,
                label_visibility="collapsed",
                key="copy_dst",
                placeholder="Selecciona una o más variantes…",
            )
            copy_dup_mode = st.radio(
                "Si el componente ya existe en destino",
                ["Omitir", "Sobreescribir"],
                horizontal=True,
                key="copy_dup_mode",
            )
            _copy_btn = st.button(
                f"Copiar BOM a {len(copy_dst_labels)} variante(s)",
                type="primary",
                use_container_width=True,
                key="btn_copy_bom",
                disabled=(not _src_entries or not copy_dst_labels),
            )

        if _copy_btn:
            _cp_added = _cp_updated = _cp_skipped = 0
            for _dst_label in copy_dst_labels:
                _dst_bc = prenda_label_to_bc.get(_dst_label, "")
                for _se in _src_entries:
                    _ex_idx = next(
                        (
                            i for i, e in enumerate(st.session_state["bom_draft"])
                            if e["Cod Barras Variante"] == _dst_bc
                            and e["EAN Componente"] == _se["EAN Componente"]
                        ),
                        None,
                    )
                    if _ex_idx is not None:
                        if copy_dup_mode == "Sobreescribir":
                            st.session_state["bom_draft"][_ex_idx]["Cantidad"] = _se["Cantidad"]
                            _cp_updated += 1
                        else:
                            _cp_skipped += 1
                    else:
                        st.session_state["bom_draft"].append({
                            "Cod Barras Variante": _dst_bc,
                            "EAN Componente": _se["EAN Componente"],
                            "Cantidad": _se["Cantidad"],
                            "_nombre_variante": _dst_label,
                            "_nombre_componente": _se["_nombre_componente"],
                        })
                        _cp_added += 1
            _parts = []
            if _cp_added:
                _parts.append(f"{_cp_added} añadidas")
            if _cp_updated:
                _parts.append(f"{_cp_updated} actualizadas")
            if _cp_skipped:
                _parts.append(f"{_cp_skipped} omitidas")
            st.success(f"Copia completada: {', '.join(_parts)}.")
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

        # Enrich with variant attributes for filtering
        _var_attrs = prenda_variants[
            ["Código de barras principal", "Referencia interna", "Color", "Talla", "_tipo_prenda"]
        ].rename(columns={"Código de barras principal": "Cod Barras Variante"})
        draft_df = draft_df.merge(_var_attrs, on="Cod Barras Variante", how="left")

        # ── Controls row ─────────────────────────────────────────────────────
        _bom_ctrl_q, _bom_ctrl_sort, _bom_ctrl_exp, _bom_ctrl_col = st.columns([4, 2.5, 1.3, 1.3])
        with _bom_ctrl_q:
            _bom_q = st.text_input(
                "Buscar en la BOM",
                placeholder="Nombre de variante o componente…",
                key="bom_table_q",
                label_visibility="collapsed",
            )
        with _bom_ctrl_sort:
            _bom_sort = st.selectbox(
                "Ordenar componentes",
                ["Sin ordenar", "Nombre A→Z", "Cantidad ↓", "Cantidad ↑"],
                key="bom_sort",
                label_visibility="collapsed",
            )
        if _bom_ctrl_exp.button("Expandir todo", use_container_width=True, key="bom_expand_all"):
            st.session_state["bom_all_expanded"] = True
            st.rerun()
        if _bom_ctrl_col.button("Contraer todo", use_container_width=True, key="bom_collapse_all"):
            st.session_state["bom_all_expanded"] = False
            st.rerun()

        # ── Filters ──────────────────────────────────────────────────────────
        _n_active_filters = sum(
            bool(st.session_state.get(k, []))
            for k in ["bom_f_refs", "bom_f_colors", "bom_f_tallas", "bom_f_tipos", "bom_f_comps"]
        )
        _filt_label = f"Filtros  ·  {_n_active_filters} activos" if _n_active_filters else "Filtros"
        with st.expander(_filt_label, expanded=False):
            _fa, _fb, _fc, _fd, _fe = st.columns(5)
            _f_refs = _fa.multiselect(
                "Referencia",
                sorted(draft_df["Referencia interna"].dropna().unique().tolist()),
                key="bom_f_refs",
                placeholder="Todas",
            )
            _f_colors = _fb.multiselect(
                "Color",
                sorted(draft_df["Color"].dropna().unique().tolist()),
                key="bom_f_colors",
                placeholder="Todos",
            )
            _f_tallas = _fc.multiselect(
                "Talla",
                sorted(draft_df["Talla"].dropna().unique().tolist()),
                key="bom_f_tallas",
                placeholder="Todas",
            )
            _f_tipos = _fd.multiselect(
                "Tipo (4º dígito)",
                sorted(draft_df["_tipo_prenda"].dropna().unique().tolist()),
                key="bom_f_tipos",
                placeholder="Todos",
            )
            _f_comps = _fe.multiselect(
                "Componente",
                sorted(draft_df["_nombre_componente"].dropna().unique().tolist()),
                key="bom_f_comps",
                placeholder="Todos",
            )

        # Apply text search
        if _bom_q:
            _mask = (
                draft_df["_nombre_variante"].str.contains(_bom_q, case=False, na=False)
                | draft_df["_nombre_componente"].str.contains(_bom_q, case=False, na=False)
            )
            draft_df = draft_df[_mask]

        # Apply multiselect filters
        if _f_refs:
            draft_df = draft_df[draft_df["Referencia interna"].isin(_f_refs)]
        if _f_colors:
            draft_df = draft_df[draft_df["Color"].isin(_f_colors)]
        if _f_tallas:
            draft_df = draft_df[draft_df["Talla"].isin(_f_tallas)]
        if _f_tipos:
            draft_df = draft_df[draft_df["_tipo_prenda"].isin(_f_tipos)]
        if _f_comps:
            draft_df = draft_df[draft_df["_nombre_componente"].isin(_f_comps)]

        _all_exp = st.session_state.get("bom_all_expanded", True)
        to_delete = None

        for bc_var, group in draft_df.groupby("Cod Barras Variante", sort=False):
            if _bom_sort == "Nombre A→Z":
                group = group.sort_values("_nombre_componente")
            elif _bom_sort == "Cantidad ↓":
                group = group.sort_values("Cantidad", ascending=False)
            elif _bom_sort == "Cantidad ↑":
                group = group.sort_values("Cantidad", ascending=True)
            var_name = group.iloc[0]["_nombre_variante"]
            n_comp = len(group)
            label = f"{var_name}  ·  {n_comp} componente{'s' if n_comp != 1 else ''}"
            with st.expander(label, expanded=_all_exp):
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
        ac1, ac_dup, ac2, ac3 = st.columns(4)

        with ac1:
            if st.button("Limpiar todo", use_container_width=True):
                st.session_state["bom_draft"] = []
                if "custom_bom" in st.session_state:
                    del st.session_state["custom_bom"]
                st.rerun()

        with ac_dup:
            _dup_pairs = [
                (e["Cod Barras Variante"], e["EAN Componente"])
                for e in st.session_state["bom_draft"]
            ]
            _n_dups = len(_dup_pairs) - len(set(_dup_pairs))
            _dup_label = f"Eliminar duplicados ({_n_dups})" if _n_dups else "Sin duplicados"
            if st.button(_dup_label, use_container_width=True, disabled=(_n_dups == 0)):
                _seen: dict = {}
                for entry in st.session_state["bom_draft"]:
                    key = (entry["Cod Barras Variante"], entry["EAN Componente"])
                    if key in _seen:
                        _seen[key]["Cantidad"] += float(entry["Cantidad"])
                    else:
                        _seen[key] = {**entry, "Cantidad": float(entry["Cantidad"])}
                st.session_state["bom_draft"] = list(_seen.values())
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

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Simulador de escenarios
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("Simulador de escenarios")
    st.markdown(
        "Introduce cantidades hipotéticas para calcular consumos sin modificar el plan real. "
        "Usa la BOM activa en ese momento."
    )

    active_bom_sc = st.session_state.get("custom_bom", bom)

    # ── Filters ──────────────────────────────────────────────────────────────
    sc_c1, sc_c2, sc_c3, sc_c4 = st.columns([3, 2, 2, 2])
    with sc_c1:
        sc_q = st.text_input("Buscar por nombre", placeholder="ej. Vestido Goya…", key="sc_q")
    with sc_c2:
        sc_color = st.selectbox(
            "Color", ["Todos"] + sorted(finished["Color"].dropna().unique().tolist()), key="sc_color"
        )
    with sc_c3:
        sc_talla = st.selectbox(
            "Talla",
            ["Todas"] + sorted(finished["Talla"].dropna().unique().tolist(), key=_sz_key),
            key="sc_talla",
        )
    with sc_c4:
        sc_ref = st.selectbox(
            "Referencia",
            ["Todas"] + sorted(finished["Referencia interna"].dropna().unique().tolist()),
            key="sc_ref",
        )

    sc_df = finished.copy()
    if sc_q:
        sc_df = sc_df[sc_df["Nombre"].str.contains(sc_q, case=False, na=False)]
    if sc_color != "Todos":
        sc_df = sc_df[sc_df["Color"] == sc_color]
    if sc_talla != "Todas":
        sc_df = sc_df[sc_df["Talla"] == sc_talla]
    if sc_ref != "Todas":
        sc_df = sc_df[sc_df["Referencia interna"] == sc_ref]

    st.caption(f"{len(sc_df)} variantes mostradas")

    # ── Bulk assignment ───────────────────────────────────────────────────────
    def _sc_bulk_apply(target_df, delta):
        for bc in target_df["Código de barras principal"]:
            sk = f"sc_qty_{bc}"
            cur = st.session_state.get(sk, 0)
            st.session_state[sk] = 0 if delta is None else max(0, cur + delta)

    with st.expander("Asignación masiva", expanded=False):
        _sm1, _sm2 = st.columns(2)
        sc_bulk_tallas = _sm1.multiselect(
            "Filtrar por talla",
            sorted(sc_df["Talla"].dropna().unique().tolist(), key=_sz_key),
            placeholder="Todas las tallas visibles",
            key="sc_bulk_tallas",
        )
        sc_bulk_colors = _sm2.multiselect(
            "Filtrar por color",
            sorted(sc_df["Color"].dropna().unique().tolist()),
            placeholder="Todos los colores visibles",
            key="sc_bulk_colors",
        )
        sc_sub_bulk = sc_df.copy()
        if sc_bulk_tallas:
            sc_sub_bulk = sc_sub_bulk[sc_sub_bulk["Talla"].isin(sc_bulk_tallas)]
        if sc_bulk_colors:
            sc_sub_bulk = sc_sub_bulk[sc_sub_bulk["Color"].isin(sc_bulk_colors)]
        st.caption(f"Afecta a **{len(sc_sub_bulk)}** de {len(sc_df)} variantes visibles")
        _sb1, _sb2, _sb3, _ss, _sb4, _sb5, _sb6, _ss2, _sb7 = st.columns(
            [1, 1, 1, 0.3, 1, 1, 1, 0.3, 1.4]
        )
        if _sb1.button("− 10", use_container_width=True, key="sc_m10"):
            _sc_bulk_apply(sc_sub_bulk, -10)
        if _sb2.button("− 5", use_container_width=True, key="sc_m5"):
            _sc_bulk_apply(sc_sub_bulk, -5)
        if _sb3.button("− 1", use_container_width=True, key="sc_m1"):
            _sc_bulk_apply(sc_sub_bulk, -1)
        if _sb4.button("+ 1", use_container_width=True, key="sc_p1"):
            _sc_bulk_apply(sc_sub_bulk, 1)
        if _sb5.button("+ 5", use_container_width=True, key="sc_p5"):
            _sc_bulk_apply(sc_sub_bulk, 5)
        if _sb6.button("+ 10", use_container_width=True, key="sc_p10"):
            _sc_bulk_apply(sc_sub_bulk, 10)
        if _sb7.button("Poner a 0", use_container_width=True, key="sc_reset"):
            _sc_bulk_apply(sc_sub_bulk, None)

    # ── Quantity matrix ───────────────────────────────────────────────────────
    _bom_covered_sc = set(active_bom_sc["Cod Barras Variante"].unique())

    for (ref, nombre), grp in sc_df.groupby(["Referencia interna", "Nombre"], sort=True):
        colors = sorted(grp["Color"].dropna().unique().tolist())
        tallas = sorted(grp["Talla"].dropna().unique().tolist(), key=_sz_key)
        bc_map = {
            (r["Color"], r["Talla"]): r["Código de barras principal"]
            for _, r in grp.iterrows()
        }
        _missing_sc = sum(1 for bc in bc_map.values() if bc not in _bom_covered_sc)
        _warn_html = (
            f'<span class="bom-warn">(sin BOM: {_missing_sc} var.)</span>'
            if _missing_sc else ""
        )
        st.markdown(
            f'<div class="ref-card"><span class="ref-title">{ref} &nbsp;·&nbsp; {nombre}</span>'
            f'{_warn_html}</div>',
            unsafe_allow_html=True,
        )
        ratios = [0.8] + [2.5] * len(colors) + [1.0]
        hcols = st.columns(ratios)
        hcols[0].write("")
        for ci, color in enumerate(colors):
            hcols[ci + 1].markdown(f'<div class="col-hdr">{color}</div>', unsafe_allow_html=True)
        hcols[-1].markdown('<div class="col-hdr-tot">Total</div>', unsafe_allow_html=True)

        col_totals = [0] * len(colors)
        for talla in tallas:
            rcols = st.columns(ratios)
            rcols[0].markdown(f'<div class="sz-lbl">{talla}</div>', unsafe_allow_html=True)
            row_total = 0
            for ci, color in enumerate(colors):
                bc = bc_map.get((color, talla))
                if bc is None:
                    continue
                sk = f"sc_qty_{bc}"
                if sk not in st.session_state:
                    st.session_state[sk] = 0
                with rcols[ci + 1]:
                    v = st.number_input(
                        talla, min_value=0, step=1, key=sk, label_visibility="collapsed"
                    )
                row_total += int(v)
                col_totals[ci] += int(v)
            cls = "tot-val" if row_total > 0 else "tot-zero"
            rcols[-1].markdown(
                f'<div class="{cls}">{row_total if row_total > 0 else "—"}</div>',
                unsafe_allow_html=True,
            )

    # ── Calculate ─────────────────────────────────────────────────────────────
    st.divider()
    sc_btn_c1, sc_btn_c2 = st.columns([2, 5])
    _sc_calcular = sc_btn_c1.button(
        "Calcular escenario", type="primary", use_container_width=True, key="sc_calcular"
    )
    if sc_btn_c2.button("Limpiar escenario", use_container_width=True, key="sc_limpiar"):
        for _k in list(st.session_state.keys()):
            if _k.startswith("sc_qty_"):
                del st.session_state[_k]
        st.session_state.pop("sc_results", None)
        st.rerun()

    if _sc_calcular:
        sc_qtys = {
            k[7:]: int(v)
            for k, v in st.session_state.items()
            if k.startswith("sc_qty_") and isinstance(v, (int, float)) and v > 0
        }
        if not sc_qtys:
            st.warning("Introduce al menos una cantidad en el escenario.")
        else:
            sc_comp_totals: dict[str, float] = {}
            for bc, qty in sc_qtys.items():
                for _, brow in active_bom_sc[active_bom_sc["Cod Barras Variante"] == bc].iterrows():
                    comp = str(brow["EAN Componente"])
                    sc_comp_totals[comp] = sc_comp_totals.get(comp, 0.0) + float(brow["Cantidad"]) * qty

            sc_results = []
            for comp_bc, total_qty in sc_comp_totals.items():
                info = barcode_lookup.get(comp_bc, {})
                col_ = str(info.get("Color", "")).strip()
                tal_ = str(info.get("Talla", "")).strip()
                sc_results.append({
                    "Referencia": info.get("Referencia interna", ""),
                    "Nombre": info.get("Nombre", f"Componente {comp_bc}"),
                    "Color": col_ if col_ not in ("", "nan") else "—",
                    "Talla": tal_ if tal_ not in ("", "nan") else "—",
                    "Código de barras": comp_bc,
                    "Cantidad escenario": round(total_qty, 4),
                })

            sc_results_df = (
                pd.DataFrame(sc_results)
                .sort_values(["Nombre", "Color", "Talla"])
                .reset_index(drop=True)
            )

            # Compare with active plan if results exist
            if "results_df" in st.session_state:
                real_df = st.session_state["results_df"][
                    ["Código de barras", "Cantidad necesaria"]
                ].rename(columns={"Cantidad necesaria": "Cantidad plan real"})
                sc_results_df = sc_results_df.merge(real_df, on="Código de barras", how="left")
                sc_results_df["Diferencia"] = (
                    sc_results_df["Cantidad escenario"]
                    - sc_results_df["Cantidad plan real"].fillna(0)
                ).round(4)

            st.session_state["sc_results"] = sc_results_df

    if "sc_results" in st.session_state:
        sc_res = st.session_state["sc_results"]
        st.subheader("Necesidades de componentes — escenario")
        st.caption(f"{len(sc_res)} componentes distintos")
        st.dataframe(sc_res, hide_index=True, use_container_width=True)


# ── Auto-save BOM draft to disk on every render ───────────────────────────────
_save_bom_draft()
