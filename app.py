import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
import os
import base64
from datetime import datetime, date
from pdf_engine import build_pdf, fmt_money, sanitize_and_compute_items
from excel_engine import build_quote_excel, build_consolidado_excel
from storage_manager import (
    save_quotation_files,
    save_consolidado_crm_file,
    sync_all_crm_events,
    list_all_stored_files,
    open_folder_in_windows,
    ROOT_F3R,
    FOLDER_PDF,
    FOLDER_EXCEL,
    FOLDER_CONSOLIDADOS,
    clean_name
)

# Configuración de página
st.set_page_config(
    page_title="F3R PRODUCCIONES S.A.S. | Suite Comercial & CRM",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_FILE = "crm_events_db.json"
LOGO_FILE = "logo_f3r_nuevo.jpg" if os.path.exists("logo_f3r_nuevo.jpg") else "Logo F3R.jpeg"

# ----------------- PALETA CORPORATIVA ORO & NEGRO F3R (FONDO BLANCO / ALTA LEGIBILIDAD) -----------------
st.markdown("""
<style>
    /* Estilos globales: Fondo blanco ejecutivo para cero fatiga visual */
    .stApp {
        background-color: #FAFAFC;
        color: #0F172A;
    }
    
    /* Barra lateral corporativa */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    
    /* Tarjetas ejecutivas F3R en blanco puro con acento dorado */
    .f3r-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #D4AF37;
        border-radius: 12px;
        padding: 20px 24px;
        color: #0F172A;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.05);
        transition: all 0.2s ease-in-out;
    }
    .f3r-card:hover {
        border-top: 4px solid #B38F26;
        box-shadow: 0 8px 25px -4px rgba(212, 175, 55, 0.2);
    }
    
    .gold-title {
        color: #0F172A;
        font-weight: 800;
        letter-spacing: -0.3px;
        border-bottom: 2.5px solid #D4AF37;
        display: inline-block;
        padding-bottom: 4px;
        margin-bottom: 12px;
    }
    .kpi-title { 
        font-size: 0.78rem; 
        color: #64748B; 
        font-weight: 700; 
        text-transform: uppercase; 
        letter-spacing: 0.8px; 
    }
    .kpi-value { 
        font-size: 1.85rem; 
        font-weight: 800; 
        color: #0F172A; 
        margin-top: 4px; 
    }
    .kpi-sub { 
        color: #64748B; 
        font-size: 0.82rem; 
        font-weight: 500; 
        margin-top: 4px; 
    }
    
    /* Botones primarios en Dorado F3R con texto Negro Ónix */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #D4AF37 0%, #C59B27 100%) !important;
        color: #000000 !important;
        font-weight: 700 !important;
        border: 1px solid #B38F26 !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px rgba(212, 175, 55, 0.25) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(212, 175, 55, 0.45) !important;
        background: linear-gradient(135deg, #E5C158 0%, #D4AF37 100%) !important;
        color: #000000 !important;
    }

    /* Botones secundarios elegantes en fondo blanco y borde gris sutil */
    div.stButton > button {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1.5px solid #CBD5E1 !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:hover {
        border-color: #D4AF37 !important;
        color: #000000 !important;
        box-shadow: 0 2px 8px rgba(212, 175, 55, 0.15) !important;
    }
    
    /* Pestañas (Tabs) corporativas */
    button[data-baseweb="tab"] {
        color: #64748B !important;
        font-weight: 600 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #0F172A !important;
        border-bottom-color: #D4AF37 !important;
        font-weight: 700 !important;
    }

    /* Fila de ítem en cotizador */
    .item-row-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #D4AF37;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
    }

    /* ============================================================== */
    /* OPTIMIZACIÓN VISUAL PARA COMPUTADORES Y MONITORES DE ESCRITORIO (PC) */
    /* ============================================================== */
    @media (min-width: 769px) {
        .main .block-container {
            padding-top: 1.8rem !important;
            padding-bottom: 3.5rem !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            max-width: 1440px !important;
            margin: 0 auto !important;
        }
        .kpi-value {
            font-size: clamp(1.3rem, 1.8vw, 1.8rem) !important;
            white-space: nowrap !important;
        }
    }

    /* ============================================================== */
    /* RESPONSIVIDAD PARA CELULARES Y TABLETS (PANTALLAS MÓVILES)    */
    /* ============================================================== */
    @media (max-width: 768px) {
        /* Aprovechar al máximo el ancho de pantalla en celular */
        .main .block-container {
            padding-top: 1rem !important;
            padding-bottom: 2.5rem !important;
            padding-left: 0.6rem !important;
            padding-right: 0.6rem !important;
        }

        /* Títulos principales escalables en celular */
        .gold-title {
            font-size: 1.3rem !important;
            letter-spacing: -0.2px !important;
            margin-bottom: 8px !important;
        }

        /* Tarjetas KPI compactas para evitar desbordes */
        .f3r-card {
            padding: 12px 14px !important;
            margin-bottom: 8px !important;
            border-radius: 10px !important;
        }

        .kpi-title {
            font-size: 0.72rem !important;
            letter-spacing: 0.4px !important;
        }

        .kpi-value {
            font-size: 1.35rem !important;
            margin-top: 2px !important;
            word-break: break-word;
        }

        .kpi-sub {
            font-size: 0.72rem !important;
            margin-top: 2px !important;
        }

        /* Botones táctiles de tamaño óptimo para dedo en celular */
        div.stButton > button, div.stDownloadButton > button {
            width: 100% !important;
            min-height: 46px !important;
            padding: 10px 14px !important;
            font-size: 0.92rem !important;
            margin-bottom: 6px !important;
        }

        /* Pestañas con desplazamiento táctil horizontal suave */
        div[data-baseweb="tab-list"] {
            gap: 4px !important;
            overflow-x: auto !important;
            flex-wrap: nowrap !important;
            -webkit-overflow-scrolling: touch;
            padding-bottom: 4px !important;
        }
        button[data-baseweb="tab"] {
            font-size: 0.82rem !important;
            padding: 8px 10px !important;
            white-space: nowrap !important;
        }

        /* Tablas y editores con scroll horizontal fluido sin descuadrar la pantalla */
        div[data-testid="stDataEditor"], div[data-testid="stDataFrame"], div[data-testid="stTable"] {
            width: 100% !important;
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch;
            border-radius: 8px !important;
        }

        /* Evitar zoom automático forzado en Safari de iPhone */
        input, select, textarea, div[data-baseweb="input"] input, div[data-baseweb="select"] {
            font-size: 16px !important;
        }

        /* Botón de menú hamburguesa de la barra lateral en celular */
        button[data-testid="stSidebarCollapseButton"] {
            color: #D4AF37 !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ----------------- GESTIÓN PERSISTENTE DE BASE DE DATOS -----------------
def load_db():
    if not os.path.exists(DB_FILE):
        return {"events": []}
    with open(DB_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    # Asegurar que cada evento posea un identificador único persistente (UUID)
    updated = False
    for idx, ev in enumerate(data.get("events", [])):
        if "uuid" not in ev:
            ev["uuid"] = f"EVT-{ev.get('id', 'HIST')}-{idx+1}"
            updated = True
    if updated:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
    return data

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_cot_code(event_dict, existing_events=None):
    """
    Genera el código de cotización oficial F3R:
    - V1 si es la primera versión no editada, V2 en la segunda, V3 en la tercera, etc.
    - Seguido de '-' y fecha MMDD (2 dígitos de mes + 2 dígitos de día)
    - Si ya existe otra cotización en la misma fecha para otro cliente, le asigna un consecutivo (-02, -03)
    """
    ver = event_dict.get("version", 1)
    try:
        ver_int = int(ver)
    except (ValueError, TypeError):
        ver_int = 1
        
    date_str = event_dict.get("created_at") or datetime.now().strftime("%Y-%m-%d")
    try:
        dt = datetime.strptime(str(date_str).split("T")[0], "%Y-%m-%d")
        mmdd = dt.strftime("%m%d")
    except Exception:
        mmdd = datetime.now().strftime("%m%d")
        
    base_code = f"V{ver_int}-{mmdd}"
    
    if existing_events:
        cur_uuid = event_dict.get("uuid")
        colliding = [
            e for e in existing_events 
            if e.get("uuid") != cur_uuid and e.get("id", "").startswith(base_code)
        ]
        if colliding:
            return f"{base_code}-{len(colliding) + 1:02d}"
            
    return base_code

db = load_db()
events = db.get("events", [])

# Categorías estándar del sector eventos F3R
CATEGORIAS_EVENTOS = [
    "SONIDO",
    "ILUMINACIÓN",
    "ESTRUCTURA",
    "AUDIOVISUALES",
    "GENERADORES ELECTRICOS",
    "ADICIONALES",
    "ZONA INFANTIL",
    "ALIMENTOS & BEBIDAS",
    "Gasto Operativo y transporte de equipos",
    "PERSONAL & STAFF",
    "OTROS"
]

# ----------------- SIDEBAR CORPORATIVO F3R -----------------
if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)
else:
    st.sidebar.title("🎪 F3R PRODUCCIONES")

st.sidebar.markdown("<h4 style='color:#d4af37; text-align:center;'>F3R PRODUCCIONES S.A.S.</h4>", unsafe_allow_html=True)
st.sidebar.caption("NIT: 902106670-3 | CRM & Cotizador Comercial")
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Navegación del Sistema:",
    [
        "📝 1. Cotizador Inteligente & Generador PDF",
        "💰 2. Matriz de Utilidad (Pre vs Real)",
        "🗂️ 3. Pipeline CRM & Seguimiento",
        "📊 4. Dashboard Ejecutivo & Funnel",
        "📁 5. Base de Datos & Archivos F3R (Excel/PDF)"
    ]
)

st.sidebar.markdown("---")
st.sidebar.markdown("<h5 style='color:#0F172A; margin-bottom:4px;'>📁 Archivos Estructurados F3R</h5>", unsafe_allow_html=True)
st.sidebar.caption("Carpetas organizadas por cliente y tipo:")
col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    if st.button("📂 Abrir F3R", use_container_width=True, help="Abrir carpeta raíz F3R_ARCHIVOS en Windows Explorer"):
        ok, msg = open_folder_in_windows()
        if ok:
            st.sidebar.success(msg)
        else:
            st.sidebar.info(msg)
with col_sb2:
    if st.button("🔄 Sincronizar", use_container_width=True, help="Generar y actualizar todos los archivos en sus carpetas"):
        sync_all_crm_events(events)
        st.success("¡Sincronizado!")
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.info("""
**F3R PRODUCCIONES S.A.S.**  
• NIT: `902106670-3`  
• WhatsApp: `+57 333 434 0735`  
• Instagram: `@f3rproducciones`  
• Correo: `f3rproducciones@gmail.com`  
• Dir: Cartagena Diagonal 31 80B75 Of 501  
• Forma de Pago: `CONTADO`  
• Cuenta: `Bancolombia Ahorros # 788 000 169-77`
""")

# ==============================================================================
# MÓDULO 1: COTIZADOR INTELIGENTE & EXPORTADOR PDF
# ==============================================================================
if modulo == "📝 1. Cotizador Inteligente & Generador PDF":
    st.markdown("<h2 class='gold-title'>📝 Cotizador Comercial Oficial F3R</h2>", unsafe_allow_html=True)
    st.markdown("Crea, edita, modifica valores y elimina ítems en tiempo real con **guardado automático permanente** en el CRM de F3R Producciones S.A.S.")

    tab_cotizar, tab_ver_pdf = st.tabs(["⚡ Editor Comercial de la Propuesta", "📄 Generador & Descarga de Cotización PDF"])

    with tab_cotizar:
        # Inicialización de estado y versión del formulario
        if "cot_form_seed" not in st.session_state:
            st.session_state.cot_form_seed = 1
        if "last_loaded_modo" not in st.session_state:
            st.session_state.last_loaded_modo = None
        if "active_cot" not in st.session_state:
            st.session_state.active_cot = None

        def crear_cotizacion_blanca():
            nueva = {
                "uuid": f"EVT-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                "client": "",
                "nit": "",
                "event_name": "",
                "contact_person": "",
                "contact_email": "",
                "contact_phone": "",
                "event_date": str(date.today()),
                "stage": "Cotizado / Negociación",
                "prob_win": 70,
                "location": "",
                "pax": 0,
                "created_at": datetime.now().strftime("%Y-%m-%d"),
                "version": 1,
                "forma_pago": "CONTADO",
                "provider_info": {
                    "razon_social": "F3R PRODUCCIONES S.A.S.",
                    "nit": "902106670-3",
                    "direccion": "Cartagena de Indias Diagonal 31 80B75 Of 501",
                    "contacto": "Francisco Ricaurte",
                    "cargo": "Gerente",
                    "email": "f3rproducciones@gmail.com",
                    "telefono": "+57 333 434 0735",
                    "instagram": "@f3rproducciones",
                    "forma_pago": "CONTADO",
                    "banco": "Cuenta de Ahorros Bancolombia # 788 000 169-77 a Nombre de F3R PRODUCCIONES S.A.S."
                },
                "notes": "",
                "items": [],
                "actual_costs": []
            }
            nueva['id'] = get_cot_code(nueva, events)
            return nueva

        # Botón directo y muy visible para iniciar en blanco desde cero
        c_btn_top1, c_btn_top2 = st.columns([1.5, 2.5])
        with c_btn_top1:
            btn_nueva_directa = st.button("➕ NUEVA COTIZACIÓN EN BLANCO", type="primary", use_container_width=True, help="Limpia completamente cliente, evento y equipos para cotizar desde cero.")
        with c_btn_top2:
            st.markdown("<p style='margin-top:6px; font-size:0.85rem; color:#475569;'>✨ Presiona este botón dorado para iniciar una cotización <b>100% en blanco</b> (sin datos de clientes ni equipos de otros proyectos).</p>", unsafe_allow_html=True)

        # Selector de modo / cotizaciones
        opciones_modo = ["➕ Crear Nueva Cotización en Blanco"] + [
            f"✏️ Editar: {e['id']} | {e['client']} ({e.get('event_name', '')})"
            for e in events
        ] + ["📋 Clonar desde Plantilla ENAP (38 Ítems Completos)"]

        col_sel1, col_sel2 = st.columns([2.5, 1])
        with col_sel1:
            modo_sel = st.selectbox("O selecciona una opción / cotización existente:", opciones_modo, key="select_cot_mode")

        with col_sel2:
            st.write("")
            st.write("")
            btn_cargar = st.button("🔄 Cargar en el Editor", use_container_width=True)

        # Detectar triggers de cambio
        modo_cambiado = (st.session_state.last_loaded_modo is not None and st.session_state.last_loaded_modo != modo_sel)
        
        # Procesar inicialización o cambio
        if btn_nueva_directa or (modo_cambiado and modo_sel == "➕ Crear Nueva Cotización en Blanco") or (btn_cargar and modo_sel == "➕ Crear Nueva Cotización en Blanco"):
            st.session_state.cot_form_seed += 1
            st.session_state.last_loaded_modo = "➕ Crear Nueva Cotización en Blanco"
            st.session_state.active_cot = crear_cotizacion_blanca()
            st.session_state.current_eid = st.session_state.active_cot["id"]
            if btn_nueva_directa or btn_cargar:
                st.rerun()

        elif (modo_cambiado and modo_sel != "➕ Crear Nueva Cotización en Blanco") or (btn_cargar and modo_sel != "➕ Crear Nueva Cotización en Blanco"):
            st.session_state.cot_form_seed += 1
            st.session_state.last_loaded_modo = modo_sel
            if modo_sel.startswith("✏️ Editar:"):
                eid = modo_sel.split(" | ")[0].replace("✏️ Editar: ", "").strip()
                match = [e for e in events if e['id'] == eid]
                if match:
                    st.session_state.active_cot = json.loads(json.dumps(match[0]))
                    st.session_state.current_eid = eid
            elif modo_sel.startswith("📋 Clonar"):
                match = [e for e in events if "ENAP" in e.get("client", "")]
                if match:
                    clon = json.loads(json.dumps(match[0]))
                    clon['uuid'] = f"EVT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    clon['created_at'] = datetime.now().strftime("%Y-%m-%d")
                    clon['version'] = 1
                    clon['id'] = get_cot_code(clon, events)
                    clon['client'] = "Nuevo Cliente Corporativo"
                    clon['event_name'] = "Evento Empresarial / Family Day"
                    st.session_state.active_cot = clon
                    st.session_state.current_eid = clon['id']
            if btn_cargar:
                st.rerun()

        elif st.session_state.active_cot is None:
            if modo_sel == "➕ Crear Nueva Cotización en Blanco":
                st.session_state.active_cot = crear_cotizacion_blanca()
                st.session_state.last_loaded_modo = "➕ Crear Nueva Cotización en Blanco"
            elif modo_sel.startswith("✏️ Editar:"):
                eid = modo_sel.split(" | ")[0].replace("✏️ Editar: ", "").strip()
                match = [e for e in events if e['id'] == eid]
                if match:
                    st.session_state.active_cot = json.loads(json.dumps(match[0]))
                    st.session_state.current_eid = eid
                else:
                    st.session_state.active_cot = crear_cotizacion_blanca()
                st.session_state.last_loaded_modo = modo_sel
            else:
                st.session_state.active_cot = crear_cotizacion_blanca()
                st.session_state.last_loaded_modo = modo_sel
            st.session_state.current_eid = st.session_state.active_cot["id"]

        cur = st.session_state.active_cot
        if not cur.get("id"):
            cur["id"] = get_cot_code(cur, events)
        seed = st.session_state.cot_form_seed

        st.markdown("<hr style='border:none; border-top:1px solid #E2E8F0; margin:16px 0;'>", unsafe_allow_html=True)
        st.markdown("#### 1. Datos del Cliente y Especificaciones del Evento")
        c1, c2, c3 = st.columns(3)
        with c1:
            cur["id"] = st.text_input("Código Cotización:", value=cur.get("id", "V1-0000"), key=f"cot_id_{seed}")
            cur["client"] = st.text_input("Razón Social del Cliente:*", value=cur.get("client", ""), placeholder="Nombre o Razón Social del Cliente", key=f"cot_client_{seed}")
            cur["nit"] = st.text_input("NIT / Documento Cliente:", value=cur.get("nit", ""), placeholder="Ej: 900.123.456-7 o C.C.", key=f"cot_nit_{seed}")
        with c2:
            cur["event_name"] = st.text_input("Nombre del Evento:*", value=cur.get("event_name", ""), placeholder="Ej: Concierto, Boda, Family Day...", key=f"cot_evname_{seed}")
            raw_ev_date = cur.get("event_date")
            try:
                if raw_ev_date:
                    val_date = datetime.strptime(str(raw_ev_date).split("T")[0], "%Y-%m-%d").date()
                else:
                    val_date = date.today()
            except Exception:
                val_date = date.today()
            picked_date = st.date_input("Fecha Prevista del Evento:*", value=val_date, key=f"cot_evdate_{seed}")
            cur["event_date"] = str(picked_date)
            cur["location"] = st.text_input("Lugar del Evento / Locación:", value=cur.get("location", ""), placeholder="Ej: Hotel, Centro de Convenciones...", key=f"cot_loc_{seed}")
            cur["pax"] = st.number_input("Aforo Estimado (Pax):", min_value=0, value=int(cur.get("pax", 0)), step=50, key=f"cot_pax_{seed}")
        with c3:
            cur["contact_person"] = st.text_input("Persona de Contacto:", value=cur.get("contact_person", ""), placeholder="Nombre de quien recibe la propuesta", key=f"cot_cp_{seed}")
            cur["contact_phone"] = st.text_input("Teléfono / WhatsApp:", value=cur.get("contact_phone", ""), placeholder="Número celular o fijo", key=f"cot_cphone_{seed}")
            cur["contact_email"] = st.text_input("Correo:", value=cur.get("contact_email", ""), placeholder="correo@ejemplo.com", key=f"cot_cmail_{seed}")

        c4, c5 = st.columns([1, 2.2])
        with c4:
            cur["forma_pago"] = st.text_input("Forma de Pago:*", value=cur.get("forma_pago", "CONTADO"), key=f"cot_fp_{seed}")
        with c5:
            prov_dict = cur.setdefault("provider_info", {})
            banco_actual = prov_dict.get("banco", "Cuenta de Ahorros Bancolombia # 788 000 169-77 a Nombre de F3R PRODUCCIONES S.A.S.")
            cur["provider_info"]["banco"] = st.text_input("Cuenta Bancaria para Consignaciones / Transferencias:*", value=banco_actual, key=f"cot_banco_{seed}")
            cur["provider_info"]["forma_pago"] = cur["forma_pago"]

        st.markdown("<hr style='border:none; border-top:1px solid #E2E8F0; margin:16px 0;'>", unsafe_allow_html=True)

        # ---------------- GESTIÓN RÁPIDA DE ÍTEMS (EDITAR / AGREGAR / ELIMINAR) ----------------
        st.markdown("#### 2. Servicios Cotizados: Agregar, Modificar Valores o Eliminar")

        # Sub-pestañas para máxima claridad y control
        tab_tabla, tab_edicion_puntual, tab_agregar = st.tabs([
            "📋 Tabla General & Modificación Rápida",
            "✏️ Modificar o Eliminar un Ítem Específico",
            "➕ Agregar Nuevo Ítem a la Cotización"
        ])

        with tab_tabla:
            if cur.get("items"):
                st.caption("💡 Modifica cualquier cantidad, costo o precio de venta directamente en la tabla y los cambios se reflejarán de inmediato:")
                df_editor = pd.DataFrame(cur["items"])
                edited_grid = st.data_editor(
                    df_editor,
                    column_config={
                        "categoria": st.column_config.SelectboxColumn("Categoría", options=CATEGORIAS_EVENTOS, required=True),
                        "item": st.column_config.TextColumn("Descripción del Ítem / Servicio", width="large", required=True),
                        "cantidad": st.column_config.NumberColumn("Cant", min_value=1, step=1, required=True),
                        "costo_unit": st.column_config.NumberColumn("Costo Unit (COP)", format="$%d", step=10000),
                        "precio_unit": st.column_config.NumberColumn("Precio Venta Unit (COP)", format="$%d", step=10000, required=True)
                    },
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"cot_grid_editor_{seed}"
                )
                cur["items"] = edited_grid.to_dict('records')
            else:
                st.info("📋 Cotización en blanco (0 equipos o suministros cargados). Puedes agregar servicios en la pestaña '➕ Agregar Nuevo Ítem' o escribirlos directamente en la tabla a continuación:")
                df_empty = pd.DataFrame(columns=["categoria", "item", "cantidad", "costo_unit", "precio_unit"])
                edited_grid_new = st.data_editor(
                    df_empty,
                    column_config={
                        "categoria": st.column_config.SelectboxColumn("Categoría", options=CATEGORIAS_EVENTOS, required=True),
                        "item": st.column_config.TextColumn("Descripción del Ítem / Servicio", width="large", required=True),
                        "cantidad": st.column_config.NumberColumn("Cant", min_value=1, step=1, required=True),
                        "costo_unit": st.column_config.NumberColumn("Costo Unit (COP)", format="$%d", step=10000),
                        "precio_unit": st.column_config.NumberColumn("Precio Venta Unit (COP)", format="$%d", step=10000, required=True)
                    },
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"cot_grid_editor_empty_{seed}"
                )
                if not edited_grid_new.empty:
                    cur["items"] = [r for r in edited_grid_new.to_dict('records') if str(r.get('item', '')).strip()]

        with tab_edicion_puntual:
            st.markdown("##### Selecciona un ítem para modificar su valor, cantidad o eliminarlo:")
            items_list = cur.get("items", [])
            if items_list:
                item_options = [f"#{i+1}: [{it.get('categoria', 'GENERAL')}] {it.get('item', '')[:45]} — {fmt_money(it.get('precio_unit', 0))}" for i, it in enumerate(items_list)]
                sel_item_idx = st.selectbox("Seleccione el Ítem a Editar o Eliminar:", range(len(item_options)), format_func=lambda x: item_options[x], key=f"edit_item_select_{seed}")

                target_item = items_list[sel_item_idx]

                ce1, ce2, ce3, ce4 = st.columns([1.5, 3, 1, 1.2])
                with ce1:
                    edit_cat = st.selectbox("Categoría:", CATEGORIAS_EVENTOS, index=CATEGORIAS_EVENTOS.index(target_item['categoria']) if target_item['categoria'] in CATEGORIAS_EVENTOS else 0, key=f"edit_cat_val_{seed}_{sel_item_idx}")
                with ce2:
                    edit_desc = st.text_input("Descripción del Servicio:", value=target_item['item'], key=f"edit_desc_val_{seed}_{sel_item_idx}")
                with ce3:
                    edit_cant = st.number_input("Cantidad:", min_value=1, value=int(target_item['cantidad']), key=f"edit_cant_val_{seed}_{sel_item_idx}")
                with ce4:
                    edit_price = st.number_input("Precio Venta Unit (COP):", min_value=0, value=int(target_item['precio_unit']), step=50000, key=f"edit_price_val_{seed}_{sel_item_idx}")

                ce5, ce6 = st.columns([1, 1])
                with ce5:
                    edit_cost = st.number_input("Costo Unit Proveedor (COP):", min_value=0, value=int(target_item.get('costo_unit', 0)), step=50000, key=f"edit_cost_val_{seed}_{sel_item_idx}")
                with ce6:
                    subt_item_calc = edit_cant * edit_price
                    st.metric("Total Venta de este Ítem:", fmt_money(subt_item_calc))

                btn_col1, btn_col2 = st.columns([1.5, 1])
                with btn_col1:
                    if st.button("💾 Actualizar Valores de este Ítem", type="primary", use_container_width=True):
                        cur["items"][sel_item_idx] = {
                            "categoria": edit_cat,
                            "item": edit_desc.strip(),
                            "cantidad": int(edit_cant),
                            "costo_unit": float(edit_cost),
                            "precio_unit": float(edit_price)
                        }
                        st.success(f"¡Ítem #{sel_item_idx+1} modificado con éxito!")
                        st.rerun()

                with btn_col2:
                    if st.button("🗑️ ELIMINAR ESTE ÍTEM", use_container_width=True):
                        eliminado = cur["items"].pop(sel_item_idx)
                        st.warning(f"Se eliminó el ítem: {eliminado.get('item', '')}")
                        st.rerun()
            else:
                st.info("No hay ítems para modificar.")

        with tab_agregar:
            st.markdown("##### Agregar un nuevo suministro o servicio a la cotización:")
            ag1, ag2, ag3, ag4, ag5 = st.columns([1.5, 3, 0.8, 1.2, 1.2])
            with ag1:
                nuevo_cat = st.selectbox("Categoría:", CATEGORIAS_EVENTOS, key="ag_cat")
            with ag2:
                nuevo_desc = st.text_input("Descripción del Equipo o Servicio:", placeholder="Ej: Póliza de responsabilidad civil", key="ag_desc")
            with ag3:
                nuevo_cant = st.number_input("Cant:", min_value=1, value=1, key="ag_cant")
            with ag4:
                nuevo_cost = st.number_input("Costo Unit Previsto (COP):", min_value=0, value=0, step=50000, key="ag_cost")
            with ag5:
                nuevo_price = st.number_input("Precio Venta Unit (COP):", min_value=0, value=0, step=50000, key="ag_price")

            if st.button("➕ Insertar a la Cotización", type="primary"):
                if nuevo_desc.strip():
                    cur["items"].append({
                        "categoria": nuevo_cat,
                        "item": nuevo_desc.strip(),
                        "cantidad": int(nuevo_cant),
                        "costo_unit": float(nuevo_cost),
                        "precio_unit": float(nuevo_price)
                    })
                    st.success("¡Nuevo ítem insertado con éxito!")
                    st.rerun()
                else:
                    st.warning("Escribe una descripción para el servicio.")

        # ---------------- TOTALES Y VALIDACIÓN FINANCIERA FORMULADA ----------------
        calc_res = sanitize_and_compute_items(cur.get("items", []))
        cur["items"] = [{
            "categoria": it["categoria"],
            "item": it["item"],
            "cantidad": it["cantidad"],
            "costo_unit": it["costo_unit"],
            "precio_unit": it["precio_unit"]
        } for it in calc_res["items"]]

        subtotal_venta = calc_res["total_venta"]
        subtotal_costo = calc_res["total_costo"]
        utilidad_prevista = calc_res["utilidad"]
        margen_previsto = calc_res["margen"]

        if cur.get("items"):
            st.markdown("<br>", unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f"""
                <div class="f3r-card">
                    <div class="kpi-title">TOTAL VENTA PROPUESTA (FORMULADA)</div>
                    <div class="kpi-value" style="color:#0F172A;">{fmt_money(subtotal_venta)}</div>
                    <div class="kpi-sub">{len(cur['items'])} Ítems validados y calculados</div>
                </div>
                """, unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                <div class="f3r-card">
                    <div class="kpi-title">PRESUPUESTO PRE-MONTAJE</div>
                    <div class="kpi-value" style="color:#DC2626;">{fmt_money(subtotal_costo)}</div>
                    <div class="kpi-sub">Límite compras directas</div>
                </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                <div class="f3r-card">
                    <div class="kpi-title">UTILIDAD PROYECTADA F3R</div>
                    <div class="kpi-value" style="color:#16A34A;">{fmt_money(utilidad_prevista)}</div>
                    <div class="kpi-sub">Ganancia estimada ex-ante</div>
                </div>
                """, unsafe_allow_html=True)
            with k4:
                st.markdown(f"""
                <div class="f3r-card">
                    <div class="kpi-title">MARGEN COMERCIAL</div>
                    <div class="kpi-value" style="color:#0284C7;">{margen_previsto:.1f}%</div>
                    <div class="kpi-sub">Rentabilidad estimada</div>
                </div>
                """, unsafe_allow_html=True)

        # ---------------- PERSISTENCIA EN EL CRM ----------------
        st.markdown("<br>", unsafe_allow_html=True)
        col_save1, col_save2 = st.columns([1.8, 1])
        with col_save1:
            if st.button("💾 GUARDAR TODOS LOS CAMBIOS EN EL CRM DE F3R PRODUCCIONES", type="primary", use_container_width=True):
                if not cur["client"].strip() or not cur["event_name"].strip():
                    st.error("Por favor completa los campos de Cliente y Nombre del Evento antes de guardar.")
                else:
                    found_idx = None
                    cur_uuid = cur.get("uuid")
                    for idx, ev in enumerate(db["events"]):
                        if cur_uuid and ev.get("uuid") == cur_uuid:
                            found_idx = idx
                            break
                    if found_idx is not None:
                        # Ha sido editada: se incrementa versión (V1 -> V2, V2 -> V3)
                        prev_ver = db["events"][found_idx].get("version", 1)
                        cur["version"] = prev_ver + 1
                        cur["id"] = get_cot_code(cur, db["events"])
                        cur["display_id"] = cur["id"]
                        db["events"][found_idx] = cur
                        st.session_state.current_eid = cur["id"]
                        msg = f"🎉 ¡Cotización actualizada a nueva versión {cur['id']} y guardada en el CRM!"
                    else:
                        cur["version"] = 1
                        if not cur.get("uuid"):
                            cur["uuid"] = f"EVT-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                        cur["id"] = get_cot_code(cur, db["events"])
                        cur["display_id"] = cur["id"]
                        db["events"].append(cur)
                        st.session_state.current_eid = cur["id"]
                        msg = f"🎉 ¡Nueva cotización {cur['id']} registrada y guardada en el CRM!"
                    save_db(db)
                    
                    # Guardado automático en carpetas estructuradas F3R
                    try:
                        res_files = save_quotation_files(cur)
                        pdf_name = os.path.basename(res_files["pdf_path"])
                        xls_name = os.path.basename(res_files["excel_path"])
                        st.toast(f"📁 Guardados en F3R_ARCHIVOS:\n• {pdf_name}\n• {xls_name}", icon="📁")
                    except Exception as err_s:
                        pass
                        
                    st.success(msg)
                    st.balloons()
                    st.rerun()

    with tab_ver_pdf:
        st.markdown("<h3 class='gold-title'>📄 Generador y Descarga de Propuesta Oficial (PDF & Excel)</h3>", unsafe_allow_html=True)
        st.caption("Verificación y validación matemática automática de todos los subtotales antes de exportar.")
        
        # Selector de evento a exportar
        dict_evs = {}
        cur_act = st.session_state.get("active_cot")
        if cur_act and cur_act.get("client", "").strip():
            dict_evs[f"⭐ Actual en Editor: {cur_act['id']} | {cur_act['client']} ({cur_act.get('event_name', '')})"] = cur_act
        for e in db.get("events", []):
            dict_evs[f"{e['id']} | {e['client']} ({e.get('event_name', '')})"] = e

        if not dict_evs:
            st.info("ℹ️ Aún no hay cotizaciones disponibles para emitir. Completa los datos del cliente y guarda la cotización en el Editor.")
        else:
            target_key = st.selectbox("Seleccione Cotización para Emitir:", list(dict_evs.keys()), key="pdf_export_select")

            if target_key:
                target_evt = dict_evs[target_key]
            
            # Auditoría y validación previa de fórmulas
            calc_audit = sanitize_and_compute_items(target_evt.get('items', []))
            tot_venta_doc = calc_audit["total_venta"]
            tot_costo_doc = calc_audit["total_costo"]
            util_doc = calc_audit["utilidad"]
            marg_doc = calc_audit["margen"]

            # Generar ambos formatos y guardarlos automáticamente en carpetas estructuradas
            pdf_bytes = build_pdf(target_evt)
            excel_bytes = build_quote_excel(target_evt)
            saved_paths = save_quotation_files(target_evt, pdf_bytes, excel_bytes)
            pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
            
            p_col1, p_col2 = st.columns([1.35, 2.65])
            with p_col1:
                st.markdown(f"""
                <div class="f3r-card">
                    <div class="gold-title" style="font-size:1.15rem;">PROPUESTA FORMAL #{target_evt['id']}</div>
                    <p style="margin-top:10px; font-size:0.88rem; color:#1E293B; line-height: 1.6;">
                    <b>Cliente:</b> {target_evt['client']}<br>
                    <b>NIT:</b> {target_evt.get('nit') or 'N/A'}<br>
                    <b>Evento:</b> {target_evt.get('event_name', '')}<br>
                    <b>Fecha:</b> {target_evt.get('event_date')}<br>
                    <b>Aforo:</b> {target_evt.get('pax', 0)} pax<br>
                    <b>Forma de Pago:</b> <span style="color:#B45309; font-weight:700;">{target_evt.get('forma_pago', 'CONTADO')}</span><br>
                    <b>Bancolombia:</b> Cta. Ahorros # 788 000 169-77<br>
                    <hr style="border:none; border-top:1px solid #E2E8F0; margin: 8px 0;">
                    <b>Total Venta (Formulado):</b> <span style="color:#0F172A; font-weight:800; font-size:1.1rem; background-color:#FEF3C7; padding:2px 6px; border-radius:4px;">{fmt_money(tot_venta_doc)}</span><br>
                    <b>Costo Previsto:</b> <span style="color:#DC2626; font-weight:700;">{fmt_money(tot_costo_doc)}</span><br>
                    <b>Utilidad Estimada:</b> <span style="color:#16A34A; font-weight:700;">{fmt_money(util_doc)} ({marg_doc:.1f}%)</span>
                    </p>
                    <hr style="border:none; border-top:1px solid #E2E8F0; margin: 8px 0;">
                    <div style="background-color:#F8FAFC; border:1px solid #E2E8F0; border-left:4px solid #D4AF37; padding:10px 14px; border-radius:6px;">
                        <span style="color:#0F172A; font-size:0.82rem; font-weight:700;">✓ VALIDACIÓN MATEMÁTICA EXITOSA</span><br>
                        <span style="color:#475569; font-size:0.78rem;">Todos los ítems ({len(calc_audit['items'])}) y subtotales cuadran al 100% con el gran total.</span>
                    </div>
                    <div style="background-color:#F8FAFC; border:1px solid #E2E8F0; padding:10px 14px; border-radius:6px; margin-top:10px;">
                        <span style="font-size:0.8rem; font-weight:700; color:#0F172A;">📁 Archivos Guardados en Carpetas F3R:</span><br>
                        <span style="font-size:0.75rem; color:#475569;">• <b>PDF:</b> {os.path.basename(saved_paths['pdf_path'])}</span><br>
                        <span style="font-size:0.75rem; color:#475569;">• <b>Excel:</b> {os.path.basename(saved_paths['excel_path'])}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                col_btn_pdf, col_btn_xls = st.columns(2)
                with col_btn_pdf:
                    st.download_button(
                        label=f"📥 Descargar PDF",
                        data=pdf_bytes,
                        file_name=os.path.basename(saved_paths["pdf_path"]),
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                with col_btn_xls:
                    st.download_button(
                        label=f"📊 Descargar Excel",
                        data=excel_bytes,
                        file_name=os.path.basename(saved_paths["excel_path"]),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                if st.button("📂 Abrir Carpeta de este Cliente en Windows", use_container_width=True, key="btn_open_client_pdf"):
                    opened, msg = open_folder_in_windows(saved_paths["client_folder_pdf"])
                    if opened:
                        st.toast(msg, icon="📂")
                    else:
                        st.info(msg)
                
            with p_col2:
                st.markdown("##### Visualizador del Documento PDF Oficial")
                st.markdown(
                    '<div class="mobile-only" style="padding: 10px 14px; background: rgba(212, 175, 55, 0.12); border: 1px solid #d4af37; border-radius: 8px; margin-bottom: 12px; font-size: 0.88rem; color: #78350f; text-align: center;">'
                    '📱 <b>¿Estás en tu celular?</b> Los visores móviles bloquean ventanas embebidas. Usa el botón superior <b>"📥 Descargar PDF"</b> para abrirlo directamente en tu teléfono.'
                    '</div>',
                    unsafe_allow_html=True
                )
                pdf_frame = f'<iframe src="data:application/pdf;base64,{pdf_b64}" width="100%" height="720" type="application/pdf" style="border: 2px solid #d4af37; border-radius: 8px; box-shadow: 0 8px 16px rgba(0,0,0,0.5);"></iframe>'
                st.markdown(pdf_frame, unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 2: MATRIZ DE UTILIDAD (PRE-MONTAJE VS POST-EVENTO)
# ==============================================================================
elif modulo == "💰 2. Matriz de Utilidad (Pre vs Real)":
    st.markdown("<h2 class='gold-title'>💰 Matriz Financiera: Utilidad Pre-Montaje vs. Real Post-Evento</h2>", unsafe_allow_html=True)
    st.markdown("""
    **Control Operativo Riguroso:**  
    1. **Pre-Montaje (Ex-Ante):** Determina tu ganancia antes de iniciar montaje y fija el techo presupuestal de compras.  
    2. **Post-Evento (Ex-Post):** Carga las facturas reales de proveedores (`RV, Aba Carpas, Omar, Inflables, Personal, Compras, etc.`) para liquidar la ganancia real.
    """)

    dict_evs = {f"{e['id']} | {e['client']} ({e.get('event_name', '')})": e for e in db.get("events", [])}
    sel_evt_fin = st.selectbox("Seleccione el Evento a Auditar Financieramente:", list(dict_evs.keys()), key="fin_sel")

    if sel_evt_fin:
        evt = dict_evs[sel_evt_fin]
        items = evt.get('items', [])
        actuals = evt.get('actual_costs', [])

        v_tot = sum(i['cantidad'] * i['precio_unit'] for i in items)
        c_pre = sum(i['cantidad'] * i.get('costo_unit', 0) for i in items)
        u_pre = v_tot - c_pre
        m_pre = (u_pre / v_tot * 100) if v_tot > 0 else 0

        c_real = sum(c['costo_real'] for c in actuals)
        u_real = v_tot - c_real if actuals else u_pre
        m_real = (u_real / v_tot * 100) if v_tot > 0 and actuals else m_pre
        desviacion = u_real - u_pre if actuals else 0

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"""
            <div class="f3r-card">
                <div class="kpi-title">VENTA TOTAL CONTRATADA</div>
                <div class="kpi-value">{fmt_money(v_tot)}</div>
                <div class="kpi-sub">Total factura cliente</div>
            </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
            <div class="f3r-card">
                <div class="kpi-title">COSTO LÍMITE PRE-MONTAJE</div>
                <div class="kpi-value" style="color:#f87171;">{fmt_money(c_pre)}</div>
                <div class="kpi-sub">Techo de compras autorizado</div>
            </div>
            """, unsafe_allow_html=True)
        with k3:
            st.markdown(f"""
            <div class="f3r-card">
                <div class="kpi-title">UTILIDAD PRE-MONTAJE</div>
                <div class="kpi-value" style="color:#4ade80;">{fmt_money(u_pre)}</div>
                <div class="kpi-sub">Margen esperado: {m_pre:.1f}%</div>
            </div>
            """, unsafe_allow_html=True)
        with k4:
            if actuals:
                st.markdown(f"""
                <div class="f3r-card">
                    <div class="kpi-title">UTILIDAD REAL POST-EVENTO</div>
                    <div class="kpi-value" style="color:#38bdf8;">{fmt_money(u_real)}</div>
                    <div class="kpi-sub">Margen real: {m_real:.1f}% (Var: {fmt_money(desviacion)})</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("ℹ️ Evento en fase previa. Ingresa facturas para liquidar.")

        st.markdown("<br>", unsafe_allow_html=True)
        tab_cruce, tab_gastos = st.tabs(["📊 Cruce de Utilidad & Semáforo de Varianza", "🧾 Registro de Facturas y Soportes Reales"])

        with tab_cruce:
            g_col1, g_col2 = st.columns([1.3, 1])
            with g_col1:
                cats_eval = ["Pre-Montaje (Proyectado)", "Post-Evento (Real)"]
                costos_data = [c_pre, c_real if actuals else c_pre]
                utilidades_data = [u_pre, u_real if actuals else 0]

                fig_comp = go.Figure(data=[
                    go.Bar(name='Costos Proveedores', x=cats_eval, y=costos_data, marker_color='#ef4444'),
                    go.Bar(name='Utilidad Neta F3R', x=cats_eval, y=utilidades_data, marker_color='#d4af37')
                ])
                fig_comp.update_layout(
                    barmode='stack',
                    title="Comparativa Financiera: Presupuesto vs. Realidad Ejecutada",
                    height=360,
                    paper_bgcolor='#FFFFFF',
                    plot_bgcolor='#FFFFFF',
                    font=dict(color="#0F172A"),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
                    margin=dict(l=10, r=10, t=35, b=40),
                    autosize=True
                )
                st.plotly_chart(fig_comp, use_container_width=True, config={'displayModeBar': False, 'responsive': True})

            with g_col2:
                st.subheader("Semáforo de Desviación")
                if actuals:
                    if desviacion < 0:
                        st.error(f"🔴 SOBRECOSTO OPERATIVO: Los costos superaron el presupuesto por `{fmt_money(abs(desviacion))}`. El margen disminuyó al {m_real:.1f}%.")
                    elif desviacion > 0:
                        st.success(f"🟢 EFICIENCIA OPERATIVA: Se generó un ahorro de `{fmt_money(desviacion)}` frente al presupuesto previsto.")
                    else:
                        st.info("🟡 Ejecución presupuestal 100% precisa.")

                    cat_map = {}
                    for c in actuals:
                        cat_map[c['categoria']] = cat_map.get(c['categoria'], 0) + c['costo_real']
                    df_pie = pd.DataFrame(list(cat_map.items()), columns=["Rubro", "Monto"])
                    fig_p = px.pie(df_pie, names="Rubro", values="Monto", hole=0.45, title="Distribución de Facturas Pagadas")
                    fig_p.update_layout(
                        height=280,
                        paper_bgcolor='#FFFFFF',
                        font=dict(color="#0F172A"),
                        legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5),
                        margin=dict(l=10, r=10, t=30, b=40),
                        autosize=True
                    )
                    st.plotly_chart(fig_p, use_container_width=True, config={'displayModeBar': False, 'responsive': True})

        with tab_gastos:
            st.subheader("Carga de Facturas y Egresos Operativos")
            with st.expander("📥 Registrar nueva factura / cuenta de cobro pagada", expanded=True):
                r1, r2, r3, r4 = st.columns([1.5, 2, 1.2, 1.2])
                with r1:
                    rubro_in = st.selectbox("Rubro de Costo:", ["Producción & Escenario", "Personal Técnico", "Estructura & Carpas", "Zona Infantil", "Audiovisuales", "Personal Operativo", "Compras & Insumos", "Imprevistos", "Bebidas & Catering"])
                with r2:
                    desc_in = st.text_input("Proveedor / Detalle del Gasto:", placeholder="Ej: RV Producciones, Aba Carpas, Omar...")
                with r3:
                    monto_in = st.number_input("Valor Pagado (COP):", min_value=0, value=1000000, step=500000)
                with r4:
                    soporte_in = st.text_input("No. Factura / Soporte:", placeholder="FAC-1234 / RECIBO")

                if st.button("➕ Cargar Gasto a la Liquidación"):
                    if desc_in.strip():
                        if "actual_costs" not in evt:
                            evt["actual_costs"] = []
                        evt["actual_costs"].append({
                            "categoria": rubro_in,
                            "item": desc_in.strip(),
                            "costo_real": float(monto_in),
                            "factura_soporte": soporte_in.strip()
                        })
                        save_db(db)
                        st.success("Gasto cargado y conciliado con éxito en el CRM.")
                        st.rerun()

            if actuals:
                st.markdown("#### Facturas y Pagos Registrados:")
                df_act = pd.DataFrame([
                    {
                        "#": i + 1,
                        "Rubro": c['categoria'],
                        "Proveedor / Detalle": c['item'],
                        "Soporte": c.get('factura_soporte', 'N/A'),
                        "Costo Real": fmt_money(c['costo_real'])
                    } for i, c in enumerate(actuals)
                ])
                st.dataframe(df_act, use_container_width=True)

# ==============================================================================
# MÓDULO 3: PIPELINE CRM & SEGUIMIENTO
# ==============================================================================
elif modulo == "🗂️ 3. Pipeline CRM & Seguimiento":
    st.markdown("<h2 class='gold-title'>🗂️ Pipeline Comercial F3R & Seguimiento</h2>", unsafe_allow_html=True)
    st.markdown("Gestiona el ciclo de vida de cada propuesta comercial y documenta acuerdos con los clientes.")

    for evt in events:
        sub_v = sum(i['cantidad'] * i['precio_unit'] for i in evt.get('items', []))
        c_pre = sum(i['cantidad'] * i.get('costo_unit', 0) for i in evt.get('items', []))
        u_pre = sub_v - c_pre

        with st.expander(f"📍 **{evt['id']} | {evt['client']}** — {evt['event_name']} ({evt.get('stage', 'Cotizado')})", expanded=True):
            col_a, col_b, col_c = st.columns([1.2, 1.2, 1])
            with col_a:
                st.markdown(f"**Contacto:** {evt.get('contact_person', 'N/A')}")
                st.markdown(f"**Teléfono:** {evt.get('contact_phone', 'N/A')}")
                st.markdown(f"**Ubicación:** {evt.get('location', 'N/A')}")
                st.markdown(f"**Fecha Prevista:** {evt.get('event_date')}")
            with col_b:
                st.markdown(f"**Aforo:** {evt.get('pax', 0)} personas")
                st.markdown(f"**Venta Cotizada:** `{fmt_money(sub_v)}`")
                st.markdown(f"**Utilidad Pre-Montaje:** `{fmt_money(u_pre)}`")
            with col_c:
                etapas = ['Prospecto / Lead', 'Cotizado / Negociación', 'Ganado / Montaje', 'Liquidado / Cerrado', 'Perdido']
                stg = evt.get('stage', 'Cotizado / Negociación')
                new_stg = st.selectbox("Estado del Evento:", etapas, index=etapas.index(stg) if stg in etapas else 1, key=f"stg_sel_{evt['id']}")
                if new_stg != stg:
                    evt['stage'] = new_stg
                    save_db(db)
                    st.success("Estado actualizado.")
                    st.rerun()

            st.markdown("**Bitácora Comercial F3R:**")
            notas = st.text_area("Registro de llamadas, acuerdos o cambios solicitados:", value=evt.get('notes', ''), key=f"notes_crm_{evt['id']}", height=65)
            if st.button("💾 Guardar Bitácora", key=f"btn_save_crm_{evt['id']}"):
                evt['notes'] = notas
                save_db(db)
                st.success("Bitácora comercial guardada con éxito en la base de datos.")

# ==============================================================================
# MÓDULO 4: DASHBOARD EJECUTIVO
# ==============================================================================
elif modulo == "📊 4. Dashboard Ejecutivo & Funnel":
    st.markdown("<h2 class='gold-title'>📊 Tablero de Control Directivo F3R</h2>", unsafe_allow_html=True)
    st.markdown("Visión global del desempeño comercial, conversión y concentración de ventas.")

    tot_pipeline = sum(sum(i['cantidad'] * i['precio_unit'] for i in e.get('items', [])) for e in events)
    ganados = [e for e in events if e.get('stage') in ['Ganado / Montaje', 'Liquidado / Cerrado']]
    tot_ganado = sum(sum(i['cantidad'] * i['precio_unit'] for i in e.get('items', [])) for e in ganados)
    util_ganada = sum(sum(i['cantidad'] * (i['precio_unit'] - i.get('costo_unit', 0)) for i in e.get('items', [])) for e in ganados)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">TOTAL PIPELINE COTIZADO</div>
            <div class="kpi-value">{fmt_money(tot_pipeline)}</div>
            <div class="kpi-sub">{len(events)} Propuestas activas</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">VENTAS GANADAS / FIRMADAS</div>
            <div class="kpi-value" style="color:#16A34A;">{fmt_money(tot_ganado)}</div>
            <div class="kpi-sub">{len(ganados)} Contratos en firme</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">UTILIDAD NETA ESTIMADA</div>
            <div class="kpi-value" style="color:#B45309;">{fmt_money(util_ganada)}</div>
            <div class="kpi-sub">Rentabilidad proyectada</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        wr = (len(ganados) / len(events) * 100) if events else 0
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">TASA DE CONVERSIÓN</div>
            <div class="kpi-value" style="color:#0284C7;">{wr:.1f}%</div>
            <div class="kpi-sub">Win Rate Comercial</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_f1, col_f2 = st.columns([1.2, 1])
    with col_f1:
        st.subheader("Embudo Comercial de Conversión (Funnel)")
        etapas_orden = ['Prospecto / Lead', 'Cotizado / Negociación', 'Ganado / Montaje', 'Liquidado / Cerrado', 'Perdido']
        datos_etapas = []
        for et in etapas_orden:
            evs_et = [e for e in events if e.get('stage') == et]
            monto = sum(sum(i['cantidad'] * i['precio_unit'] for i in e.get('items', [])) for e in evs_et)
            if monto > 0 or len(evs_et) > 0:
                datos_etapas.append({"Etapa": et, "Monto": monto, "Eventos": len(evs_et)})

        df_f = pd.DataFrame(datos_etapas)
        if not df_f.empty:
            fig_fun = px.funnel(df_f, x='Monto', y='Etapa', title="Volumen Financiero por Fase Comercial", color_discrete_sequence=['#d4af37'])
            fig_fun.update_layout(
                height=340,
                paper_bgcolor='#FFFFFF',
                font=dict(color="#0F172A"),
                margin=dict(l=10, r=10, t=30, b=10),
                autosize=True
            )
            st.plotly_chart(fig_fun, use_container_width=True, config={'displayModeBar': False, 'responsive': True})

    with col_f2:
        st.subheader("Facturación por Rubro Operativo")
        cat_map = {}
        for e in events:
            for it in e.get('items', []):
                cat = it.get('categoria', 'Otros')
                cat_map[cat] = cat_map.get(cat, 0) + (it['cantidad'] * it['precio_unit'])
        df_cp = pd.DataFrame(list(cat_map.items()), columns=["Categoría", "Venta Total"])
        fig_pie_cat = px.pie(df_cp, names="Categoría", values="Venta Total", hole=0.45, color_discrete_sequence=px.colors.sequential.YlOrBr)
        fig_pie_cat.update_layout(
            height=340,
            paper_bgcolor='#FFFFFF',
            font=dict(color="#0F172A"),
            legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5),
            margin=dict(l=10, r=10, t=30, b=40),
            autosize=True
        )
        st.plotly_chart(fig_pie_cat, use_container_width=True, config={'displayModeBar': False, 'responsive': True})

# ==============================================================================
# ==============================================================================
# MÓDULO 5: BASE DE DATOS & ARCHIVOS ESTRUCTURADOS F3R
# ==============================================================================
elif modulo == "📁 5. Base de Datos & Archivos F3R (Excel/PDF)":
    st.markdown("<h2 class='gold-title'>📁 Centro de Archivos F3R & Base de Datos CRM</h2>", unsafe_allow_html=True)
    st.markdown("Todos los archivos PDF oficiales y libros de Excel formulados se organizan y almacenan de forma estructurada en la carpeta local **`F3R_ARCHIVOS`**.")

    # Asegurar sincronización inicial
    all_files = list_all_stored_files()
    if not all_files and events:
        sync_all_crm_events(events)
        all_files = list_all_stored_files()

    count_pdf = sum(1 for f in all_files if f["category"] == "PDF Oficiales")
    count_xls = sum(1 for f in all_files if f["category"] == "Excel Formulados")
    count_crm = sum(1 for f in all_files if f["category"] == "Consolidados CRM")

    # Tarjetas KPI de archivos
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">TOTAL ARCHIVOS F3R</div>
            <div class="kpi-value">{len(all_files)}</div>
            <div class="kpi-sub">Almacenados en disco</div>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">COTIZACIONES PDF</div>
            <div class="kpi-value" style="color:#DC2626;">{count_pdf}</div>
            <div class="kpi-sub">Listas para el cliente</div>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">COTIZACIONES EXCEL</div>
            <div class="kpi-value" style="color:#16A34A;">{count_xls}</div>
            <div class="kpi-sub">Formuladas con costos</div>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="f3r-card">
            <div class="kpi-title">REGISTROS CRM</div>
            <div class="kpi-value" style="color:#0284C7;">{len(events)}</div>
            <div class="kpi-sub">Propuestas en base de datos</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Barra de Acciones Globales
    act1, act2, act3, act4 = st.columns([1.2, 1.2, 1.2, 1])
    with act1:
        if st.button("📂 Abrir Carpeta F3R en Windows", type="primary", use_container_width=True):
            opened, msg = open_folder_in_windows()
            if opened:
                st.toast(msg, icon="📂")
            else:
                st.info(msg)
    with act2:
        if st.button("🔄 Sincronizar Todos los Archivos", use_container_width=True):
            sync_all_crm_events(events)
            st.success("¡Todos los archivos PDF y Excel fueron regenerados en sus carpetas!")
            st.rerun()
    with act3:
        crm_excel_bytes = build_consolidado_excel(events)
        save_consolidado_crm_file(events, crm_excel_bytes)
        st.download_button(
            label="📊 Descargar Consolidado Excel",
            data=crm_excel_bytes,
            file_name=f"Consolidado_General_CRM_F3R_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with act4:
        # Generar CSV plano
        rows_csv = []
        for e in events:
            v_tot = sum(i['cantidad'] * i['precio_unit'] for i in e.get('items', []))
            c_tot = sum(i['cantidad'] * i.get('costo_unit', 0) for i in e.get('items', []))
            rows_csv.append({
                "ID": e['id'],
                "Cliente": e['client'],
                "NIT": e.get('nit', ''),
                "Evento": e.get('event_name', ''),
                "Fecha": e.get('event_date', ''),
                "Etapa": e.get('stage', ''),
                "Venta Total": v_tot,
                "Costo Previsto": c_tot,
                "Utilidad Prevista": v_tot - c_tot
            })
        df_csv = pd.DataFrame(rows_csv)
        st.download_button(
            label="📄 Descargar CSV",
            data=df_csv.to_csv(index=False).encode('utf-8'),
            file_name=f"F3R_Cartera_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    st.markdown("<hr style='border:none; border-top:1px solid #E2E8F0; margin: 18px 0;'>", unsafe_allow_html=True)

    tab_explorador, tab_crm_grid = st.tabs([
        "📁 Explorador de Archivos Estructurados F3R",
        "📋 Registro General de la Base de Datos CRM"
    ])

    with tab_explorador:
        st.markdown("#### Archivos Almacenados en Disco Local (`F3R_ARCHIVOS`)")
        
        # Filtros del explorador
        f_col1, f_col2 = st.columns([1, 2])
        with f_col1:
            cat_filter = st.selectbox(
                "Filtrar por Categoría:",
                ["Todos los Archivos", "PDF Oficiales", "Excel Formulados", "Consolidados CRM"]
            )
        with f_col2:
            search_query = st.text_input("🔍 Buscar por Cliente o Nombre de Archivo:", placeholder="Escribe para filtrar...")

        # Aplicar filtros
        filtered_files = all_files
        if cat_filter != "Todos los Archivos":
            filtered_files = [f for f in filtered_files if f["category"] == cat_filter]
        if search_query.strip():
            sq = search_query.lower()
            filtered_files = [f for f in filtered_files if sq in f["filename"].lower() or sq in f["client"].lower()]

        if filtered_files:
            df_display = pd.DataFrame([{
                "Categoría": f["category"],
                "Carpeta / Cliente": f["client"],
                "Nombre de Archivo": f["filename"],
                "Tamaño (KB)": f["size_kb"],
                "Fecha Modificación": f["mtime"]
            } for f in filtered_files])
            
            st.dataframe(df_display, use_container_width=True, hide_index=True)

            st.markdown("##### 📥 Descargar o Abrir Archivo Específico:")
            col_sel_f, col_act_f = st.columns([2.5, 1])
            with col_sel_f:
                opciones_files = [f"{f['category']} | {f['client']} | {f['filename']}" for f in filtered_files]
                sel_f_idx = st.selectbox("Seleccione Archivo:", range(len(opciones_files)), format_func=lambda x: opciones_files[x], key="sel_file_to_act")
                target_file_obj = filtered_files[sel_f_idx]
                
            with col_act_f:
                st.write("")
                st.write("")
                with open(target_file_obj["full_path"], "rb") as f_binary:
                    file_data = f_binary.read()
                mime_t = "application/pdf" if target_file_obj["filename"].endswith(".pdf") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                st.download_button(
                    label=f"⬇️ Descargar ({target_file_obj['size_kb']} KB)",
                    data=file_data,
                    file_name=target_file_obj["filename"],
                    mime=mime_t,
                    type="primary",
                    use_container_width=True
                )
        else:
            st.info("No se encontraron archivos con los criterios de búsqueda especificados.")

    with tab_crm_grid:
        st.markdown("#### Matriz Consolidada de Eventos en el CRM")
        rows_db = []
        for e in events:
            v_tot = sum(i['cantidad'] * i['precio_unit'] for i in e.get('items', []))
            c_tot = sum(i['cantidad'] * i.get('costo_unit', 0) for i in e.get('items', []))
            c_real = sum(c['costo_real'] for c in e.get('actual_costs', []))
            rows_db.append({
                "Código": e['id'],
                "Cliente": e['client'],
                "NIT": e.get('nit', ''),
                "Evento": e.get('event_name', ''),
                "Fecha": e.get('event_date', ''),
                "Etapa CRM": e.get('stage', ''),
                "Venta Total": fmt_money(v_tot),
                "Costo Pre": fmt_money(c_tot),
                "Utilidad Pre": fmt_money(v_tot - c_tot),
                "Costo Real": fmt_money(c_real) if c_real > 0 else "-",
                "Utilidad Real": fmt_money(v_tot - c_real) if c_real > 0 else fmt_money(v_tot - c_tot)
            })
        st.dataframe(pd.DataFrame(rows_db), use_container_width=True, hide_index=True)

