import os
import re
import subprocess
from datetime import datetime
from pdf_engine import build_pdf
from excel_engine import build_quote_excel, build_consolidado_excel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_F3R = os.path.join(BASE_DIR, "F3R_ARCHIVOS")

FOLDER_PDF = os.path.join(ROOT_F3R, "1. Cotizaciones PDF")
FOLDER_EXCEL = os.path.join(ROOT_F3R, "2. Cotizaciones Excel")
FOLDER_LIQUIDACIONES = os.path.join(ROOT_F3R, "3. Liquidaciones y Costos Reales")
FOLDER_CONSOLIDADOS = os.path.join(ROOT_F3R, "4. Consolidados CRM")

ALL_FOLDERS = [ROOT_F3R, FOLDER_PDF, FOLDER_EXCEL, FOLDER_LIQUIDACIONES, FOLDER_CONSOLIDADOS]

def clean_name(text):
    if not text:
        return "GENERAL"
    # Reemplazar caracteres no permitidos en Windows
    cleaned = re.sub(r'[\\/*?:"<>|]', "", str(text))
    cleaned = " ".join(cleaned.split()).strip(" ._")
    cleaned = cleaned[:45].strip(" ._")
    return cleaned if cleaned else "GENERAL"

def init_storage():
    """Crea la estructura de carpetas de F3R si aún no existen."""
    for folder in ALL_FOLDERS:
        os.makedirs(folder, exist_ok=True)

def get_client_folder(base_category_folder, client_name):
    clean_client = clean_name(client_name)[:45]
    folder = os.path.join(base_category_folder, clean_client)
    os.makedirs(folder, exist_ok=True)
    return folder

def save_quotation_files(evt, pdf_bytes=None, excel_bytes=None):
    """
    Genera y guarda de manera estructurada tanto el PDF como el Excel
    de una cotización en sus respectivas carpetas por cliente.
    """
    init_storage()
    
    eid = clean_name(evt.get("id", "COT-001"))
    client = clean_name(evt.get("client", "Cliente"))
    ev_name = clean_name(evt.get("event_name", "Evento"))[:30]
    
    # Generar bytes si no se proporcionaron
    if pdf_bytes is None:
        pdf_bytes = build_pdf(evt)
    if excel_bytes is None:
        excel_bytes = build_quote_excel(evt)
        
    # Rutas estructuradas
    folder_pdf_client = get_client_folder(FOLDER_PDF, client)
    folder_excel_client = get_client_folder(FOLDER_EXCEL, client)
    
    base_filename = f"Cotizacion_{eid}_{ev_name}".replace(" ", "_")
    
    path_pdf = os.path.join(folder_pdf_client, f"{base_filename}.pdf")
    path_excel = os.path.join(folder_excel_client, f"{base_filename}.xlsx")
    
    with open(path_pdf, "wb") as f_pdf:
        f_pdf.write(pdf_bytes)
        
    with open(path_excel, "wb") as f_xls:
        f_xls.write(excel_bytes)
        
    return {
        "pdf_path": path_pdf,
        "excel_path": path_excel,
        "pdf_size": len(pdf_bytes),
        "excel_size": len(excel_bytes),
        "client_folder_pdf": folder_pdf_client,
        "client_folder_excel": folder_excel_client
    }

def save_consolidado_crm_file(events, excel_bytes=None):
    """Guarda el consolidado general de cotizaciones del CRM en Excel."""
    init_storage()
    if excel_bytes is None:
        excel_bytes = build_consolidado_excel(events)
        
    filename = f"Consolidado_General_CRM_F3R_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path = os.path.join(FOLDER_CONSOLIDADOS, filename)
    with open(path, "wb") as f:
        f.write(excel_bytes)
        
    return {
        "path": path,
        "filename": filename,
        "size": len(excel_bytes)
    }

def sync_all_crm_events(events):
    """Sincroniza y guarda todos los eventos existentes en la estructura de carpetas F3R."""
    init_storage()
    saved = []
    for e in events:
        res = save_quotation_files(e)
        saved.append(res)
    return saved

def list_all_stored_files():
    """Retorna una lista completa y estructurada de todos los archivos guardados."""
    init_storage()
    files_list = []
    
    categories = [
        ("PDF Oficiales", FOLDER_PDF, "pdf"),
        ("Excel Formulados", FOLDER_EXCEL, "xlsx"),
        ("Liquidaciones de Costos", FOLDER_LIQUIDACIONES, "xlsx"),
        ("Consolidados CRM", FOLDER_CONSOLIDADOS, "xlsx"),
    ]
    
    for cat_label, root_cat, ext in categories:
        if not os.path.exists(root_cat):
            continue
        for root, dirs, files in os.walk(root_cat):
            for file in files:
                if file.lower().endswith(f".{ext}"):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, ROOT_F3R)
                    stat = os.stat(full_path)
                    
                    # Subcarpeta de cliente
                    subfolder = os.path.basename(root)
                    if subfolder == os.path.basename(root_cat):
                        client_group = "General / Directo"
                    else:
                        client_group = subfolder
                        
                    files_list.append({
                        "category": cat_label,
                        "client": client_group,
                        "filename": file,
                        "size_kb": round(stat.st_size / 1024, 1),
                        "mtime": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                        "full_path": full_path,
                        "rel_path": rel_path
                    })
                    
    # Ordenar por fecha de modificación más reciente
    files_list.sort(key=lambda x: x["mtime"], reverse=True)
    return files_list

def open_folder_in_windows(folder_path=None):
    """Abre la carpeta estructurada directamente en el Explorador de Archivos de Windows."""
    target = folder_path if (folder_path and os.path.exists(folder_path)) else ROOT_F3R
    init_storage()
    try:
        os.startfile(target)
        return True, target
    except Exception as e:
        return False, str(e)
