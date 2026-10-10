import os
import json
import warnings
from fpdf import FPDF
from datetime import datetime

warnings.filterwarnings("ignore")

def fmt_money(val):
    if val is None:
        return "$0"
    try:
        val_f = float(val)
        return f"${val_f:,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        return "$0"

def sanitize_and_compute_items(items):
    """
    Función validadora y formuladora matemática estricta:
    Asegura que cada ítem tenga cantidad, precio_unit, costo_unit y subtotales formulados correctamente.
    """
    validated_items = []
    tot_venta = 0.0
    tot_costo = 0.0

    for idx, raw in enumerate(items, start=1):
        cat = str(raw.get("categoria", "GENERAL")).strip().upper()
        desc = str(raw.get("item", "Servicio")).strip()
        
        # Validación de cantidad
        try:
            cant = int(raw.get("cantidad", 1))
            if cant <= 0:
                cant = 1
        except (ValueError, TypeError):
            cant = 1
            
        # Validación de precio de venta unitario
        try:
            p_unit = float(raw.get("precio_unit", 0))
            if p_unit < 0:
                p_unit = 0.0
        except (ValueError, TypeError):
            p_unit = 0.0

        # Validación de costo unitario
        try:
            c_unit = float(raw.get("costo_unit", 0))
            if c_unit < 0:
                c_unit = 0.0
        except (ValueError, TypeError):
            c_unit = 0.0

        # Formulación matemática exacta
        sub_venta = cant * p_unit
        sub_costo = cant * c_unit
        
        tot_venta += sub_venta
        tot_costo += sub_costo

        validated_items.append({
            "idx": idx,
            "categoria": cat,
            "item": desc,
            "cantidad": cant,
            "costo_unit": c_unit,
            "precio_unit": p_unit,
            "subtotal_venta": sub_venta,
            "subtotal_costo": sub_costo
        })

    utilidad = tot_venta - tot_costo
    margen = (utilidad / tot_venta * 100) if tot_venta > 0 else 0.0

    return {
        "items": validated_items,
        "total_venta": tot_venta,
        "total_costo": tot_costo,
        "utilidad": utilidad,
        "margen": margen
    }

class F3RQuotationPDF(FPDF):
    def __init__(self, evt):
        super().__init__(orientation='P', unit='mm', format='Letter')
        self.evt = evt
        self.set_margins(12, 10, 12)
        self.set_auto_page_break(auto=True, margin=16)

        # Cargar fuentes Unicode del sistema
        font_dir = "C:/Windows/Fonts"
        if os.path.exists(os.path.join(font_dir, "arial.ttf")):
            self.add_font("ArialUni", "", os.path.join(font_dir, "arial.ttf"))
            self.add_font("ArialUni", "B", os.path.join(font_dir, "arialbd.ttf"))
            self.add_font("ArialUni", "I", os.path.join(font_dir, "ariali.ttf"))
            self.main_font = "ArialUni"
        else:
            self.main_font = "Helvetica"

    def header(self):
        # 1. Franja Fondo Superior Negro Ónix Corporativo F3R (#000000 / #0A0A0A)
        self.set_fill_color(0, 0, 0)
        self.rect(0, 0, 216, 29, 'F')

        # Franja dorada decorativa de acento F3R Gold (#D4AF37)
        self.set_fill_color(212, 175, 55)
        self.rect(0, 29, 216, 1.2, 'F')
        
        # Logo F3R Oficial
        logo_path = "logo_f3r_nuevo.jpg" if os.path.exists("logo_f3r_nuevo.jpg") else "Logo F3R.jpeg"
        if os.path.exists(logo_path):
            try:
                self.image(logo_path, 12, 3.5, h=22)
            except Exception:
                pass
        
        # Nombre de Empresa F3R Producciones S.A.S. en Blanco
        self.set_font(self.main_font, 'B', 15)
        self.set_text_color(255, 255, 255)
        self.set_xy(38, 5)
        self.cell(90, 6, "F3R PRODUCCIONES S.A.S.", new_x="RIGHT", new_y="TOP")
        
        # Número de cotización en Oro con versionamiento dinámico
        self.set_font(self.main_font, 'B', 11)
        self.set_text_color(212, 175, 55) # Gold
        self.set_xy(125, 5)
        cot_num = self.evt.get('display_id') or self.evt.get('id', 'V1-1009')
        self.cell(79, 5, f"COTIZACIÓN: {cot_num}", align='R', new_x="LMARGIN", new_y="NEXT")
        
        # Datos de contacto en Blanco y Gris Claro
        self.set_font(self.main_font, '', 7.6)
        self.set_text_color(226, 232, 240)
        self.set_xy(38, 11.5)
        prov = self.evt.get("provider_info", {})
        nit_prov = prov.get('nit', '902106670-3')
        tel_prov = prov.get('telefono', '+57 333 434 0735')
        ig_prov = prov.get('instagram', '@f3rproducciones')
        mail_prov = prov.get('email', 'f3rproducciones@gmail.com')
        dir_prov = prov.get('direccion', 'Cartagena de Indias Diagonal 31 80B75 Of 501')
        
        self.cell(90, 4, f"NIT: {nit_prov} | WhatsApp: {tel_prov} | IG: {ig_prov}", new_x="LMARGIN", new_y="NEXT")
        self.set_x(38)
        self.cell(90, 3.8, f"{mail_prov} | {dir_prov}", new_x="RIGHT", new_y="TOP")
        
        self.set_xy(125, 11)
        self.cell(79, 4, f"Fecha Emisión: {self.evt.get('created_at', datetime.now().strftime('%Y-%m-%d'))}", align='R', new_x="LMARGIN", new_y="NEXT")
        self.set_xy(125, 15)
        self.cell(79, 4, "Vigencia: 15 días calendario | Moneda: COP", align='R', new_x="LMARGIN", new_y="NEXT")
        self.ln(7)

    def footer(self):
        self.set_y(-12)
        # Línea dorada en footer
        self.set_fill_color(212, 175, 55)
        self.rect(12, 267, 192, 0.4, 'F')
        
        self.set_font(self.main_font, 'I', 7.5)
        self.set_text_color(100, 116, 139)
        self.cell(0, 4, f"F3R Producciones S.A.S. - Propuesta Comercial Confidencial | Página {self.page_no()}/{{nb}}", align='C')

def build_pdf(evt):
    # 0. Formular y validar rigurosamente todos los datos
    calc = sanitize_and_compute_items(evt.get("items", []))
    items_val = calc["items"]
    grand_total = calc["total_venta"]

    pdf = F3RQuotationPDF(evt)
    pdf.alias_nb_pages()
    pdf.add_page()
    
    font = pdf.main_font
    
    # 1. Bloque de Datos del Cliente y F3R Producciones S.A.S.
    pdf.set_xy(12, 33)
    pdf.set_fill_color(248, 250, 252) # Blanco Slate
    pdf.set_draw_color(212, 175, 55)   # Borde Dorado F3R
    pdf.rect(12, 33, 192, 35, 'FD')
    
    # Barra de cabecera interior en Negro y Oro
    pdf.set_fill_color(0, 0, 0) # Negro Puro
    pdf.rect(12, 33, 192, 6, 'F')
    
    pdf.set_xy(15, 34)
    pdf.set_font(font, 'B', 8)
    pdf.set_text_color(212, 175, 55) # Gold
    pdf.cell(92, 4, "DATOS DEL CLIENTE", new_x="RIGHT", new_y="TOP")
    pdf.set_xy(110, 34)
    pdf.cell(90, 4, "EMPRESA EMISORA & CONDICIONES DE PAGO", new_x="LMARGIN", new_y="NEXT")
    
    # Proveedor F3R data
    prov = evt.get("provider_info", {})
    razon_prov = prov.get('razon_social', 'F3R PRODUCCIONES S.A.S.')
    nit_prov = prov.get('nit', '902106670-3')
    tel_prov = prov.get('telefono', '+57 333 434 0735')
    ig_prov = prov.get('instagram', '@f3rproducciones')
    dir_prov = prov.get('direccion', 'Cartagena de Indias Diagonal 31 80B75 Of 501')
    forma_pago_val = evt.get('forma_pago') or prov.get('forma_pago') or 'CONTADO'
    banco_val = prov.get('banco') or 'Cuenta de Ahorros Bancolombia # 788 000 169-77 a Nombre de F3R PRODUCCIONES S.A.S.'

    # Contenido Cliente
    pdf.set_text_color(15, 23, 42)
    pdf.set_font(font, '', 7.6)
    pdf.set_xy(15, 40.5)
    pdf.cell(92, 3.6, f"A nombre de: {evt.get('client', 'Cliente')[:45]}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.cell(92, 3.6, f"NIT / C.C.: {evt.get('nit', '800.141.648-9')}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.cell(92, 3.6, f"Evento: {evt.get('event_name', 'Evento')} | Aforo: {evt.get('pax', 0)} pax", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.cell(92, 3.6, f"Lugar: {evt.get('location', 'Por definir')[:45]}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.cell(92, 3.6, f"Fecha Evento: {evt.get('event_date', 'Por definir')}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(15)
    pdf.set_font(font, 'B', 7.6)
    pdf.cell(92, 3.6, f"Forma de Pago: {forma_pago_val}", new_x="RIGHT", new_y="TOP")
    
    # Contenido Proveedor F3R
    pdf.set_xy(110, 40.5)
    pdf.set_font(font, '', 7.5)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(90, 3.6, f"Razón Social: {razon_prov} (NIT {nit_prov})", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 3.6, f"Atención: Francisco Ricaurte - WhatsApp: {tel_prov}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 3.6, f"Email: f3rproducciones@gmail.com | IG: {ig_prov}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.set_font(font, 'B', 7.5)
    pdf.cell(90, 3.6, f"Forma de Pago: {forma_pago_val}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.set_font(font, '', 7.2)
    pdf.cell(90, 3.6, "Cuenta: Cuenta de Ahorros Bancolombia # 788 000 169-77", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(110)
    pdf.cell(90, 3.6, "Titular: F3R PRODUCCIONES S.A.S.", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_xy(12, 71)
    
    # 2. Agrupación y renderizado de la tabla de servicios
    grouped = {}
    for it in items_val:
        cat = it["categoria"]
        if cat not in grouped:
            grouped[cat] = []
        grouped[cat].append(it)
        
    for cat_name, cat_items in grouped.items():
        if pdf.get_y() > 230:
            pdf.add_page()
            
        # Encabezado Categoría en Negro Ónix y Dorado F3R
        pdf.set_fill_color(0, 0, 0) # Black
        pdf.set_draw_color(212, 175, 55) # Gold border
        pdf.set_text_color(212, 175, 55) # Gold text
        pdf.set_font(font, 'B', 8.5)
        
        pdf.cell(10, 6, "Item", 1, align='C', fill=True)
        pdf.cell(112, 6, f"CATEGORÍA: {cat_name}", 1, align='L', fill=True)
        pdf.set_text_color(255, 255, 255) # White
        pdf.cell(16, 6, "Cant.", 1, align='C', fill=True)
        pdf.cell(27, 6, "Vr. Unitario", 1, align='R', fill=True)
        pdf.cell(27, 6, "Valor Total", 1, align='R', fill=True, new_x="LMARGIN", new_y="NEXT")
        
        pdf.set_font(font, '', 7.5)
        pdf.set_text_color(15, 23, 42)
        pdf.set_draw_color(226, 232, 240)
        
        subtotal_cat = 0.0
        for idx, item in enumerate(cat_items, start=1):
            sub = item["subtotal_venta"]
            subtotal_cat += sub
            
            if pdf.get_y() > 245:
                pdf.add_page()
                pdf.set_fill_color(0, 0, 0)
                pdf.set_draw_color(212, 175, 55)
                pdf.set_text_color(212, 175, 55)
                pdf.set_font(font, 'B', 8)
                pdf.cell(10, 5, "Item", 1, align='C', fill=True)
                pdf.cell(112, 5, f"{cat_name} (Continuación)", 1, align='L', fill=True)
                pdf.set_text_color(255, 255, 255)
                pdf.cell(16, 5, "Cant.", 1, align='C', fill=True)
                pdf.cell(27, 5, "Vr. Unitario", 1, align='R', fill=True)
                pdf.cell(27, 5, "Valor Total", 1, align='R', fill=True, new_x="LMARGIN", new_y="NEXT")
                pdf.set_font(font, '', 7.5)
                pdf.set_text_color(15, 23, 42)
                pdf.set_draw_color(226, 232, 240)

            fill_row = (idx % 2 == 0)
            if fill_row:
                pdf.set_fill_color(248, 250, 252) # Gris perla muy claro
            else:
                pdf.set_fill_color(255, 255, 255) # Blanco
                
            y_start = pdf.get_y()
            txt_desc = str(item['item']).replace('\n', ' ')
            pdf.set_xy(22, y_start)
            pdf.multi_cell(112, 4, txt_desc, border=0, align='L', fill=False)
            y_end = pdf.get_y()
            row_h = max(5, y_end - y_start)
            
            # Dibujar celdas con alineación precisa
            pdf.set_xy(12, y_start)
            pdf.cell(10, row_h, str(idx), 1, align='C', fill=fill_row)
            pdf.set_xy(22, y_start)
            pdf.rect(22, y_start, 112, row_h, 'D')
            pdf.set_xy(134, y_start)
            pdf.cell(16, row_h, str(item['cantidad']), 1, align='C', fill=fill_row)
            pdf.cell(27, row_h, fmt_money(item['precio_unit']), 1, align='R', fill=fill_row)
            pdf.cell(27, row_h, fmt_money(sub), 1, align='R', fill=fill_row, new_x="LMARGIN", new_y="NEXT")
            pdf.set_y(y_start + row_h)
            
        # Fila Subtotal Categoría formulada
        pdf.set_font(font, 'B', 8)
        pdf.set_fill_color(241, 245, 249)
        pdf.set_draw_color(203, 213, 225)
        pdf.cell(138, 5.5, f"Subtotal {cat_name}:", 1, align='R', fill=True)
        pdf.cell(54, 5.5, fmt_money(subtotal_cat), 1, align='R', fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2.5)

    # 3. Resumen y Total General Formulado
    if pdf.get_y() > 215:
        pdf.add_page()
        
    pdf.ln(3)
    pdf.set_draw_color(212, 175, 55)
    pdf.set_fill_color(248, 250, 252)
    pdf.set_text_color(15, 23, 42)
    pdf.set_font(font, 'B', 8.5)
    pdf.set_x(102)
    pdf.cell(58, 6.5, "SUBTOTAL SERVICIOS:", 1, align='R', fill=True)
    pdf.cell(44, 6.5, fmt_money(grand_total), 1, align='R', fill=True, new_x="LMARGIN", new_y="NEXT")
    
    # Cuadro Dorado de Total Final F3R
    pdf.set_x(102)
    pdf.set_fill_color(212, 175, 55) # Gold F3R
    pdf.set_draw_color(0, 0, 0)
    pdf.set_text_color(0, 0, 0)      # Negro puro para máximo contraste
    pdf.set_font(font, 'B', 9.5)
    pdf.cell(58, 7.5, "VALOR TOTAL PROPUESTA:", 1, align='R', fill=True)
    pdf.cell(44, 7.5, fmt_money(grand_total), 1, align='R', fill=True, new_x="LMARGIN", new_y="NEXT")
    
    # 4. Términos y Condiciones Legales Exactos F3R
    pdf.ln(3)
    if pdf.get_y() > 175:
        pdf.add_page()
        
    pdf.set_fill_color(0, 0, 0) # Black
    pdf.rect(12, pdf.get_y(), 192, 5.5, 'F')
    pdf.set_xy(15, pdf.get_y() + 0.8)
    pdf.set_text_color(212, 175, 55) # Gold
    pdf.set_font(font, 'B', 8.5)
    pdf.cell(186, 4, "CONDICIONES COMERCIALES Y DE CONTRATACIÓN F3R", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font(font, '', 6.3)
    pdf.set_text_color(51, 65, 85)
    pdf.set_draw_color(212, 175, 55) # Gold border
    
    terminos_f3r = (
        "• Vigencia: Esta propuesta tiene la vigencia indicada en la cotización y está sujeta a disponibilidad de fecha, equipos y personal.\n"
        "• Reserva: La fecha se garantiza únicamente con la aceptación de la propuesta y el pago del anticipo acordado.\n"
        "• Pagos: Los pagos deberán realizarse en las fechas establecidas. El incumplimiento podrá generar suspensión del servicio y/o cobro de intereses de mora.\n"
        "• Cancelaciones: En caso de cancelación por parte del cliente, el anticipo no será reembolsable, en razón de las reservas y compromisos adquiridos por F3R Producciones.\n"
        "• Reprogramaciones: Todo cambio de fecha estará sujeto a disponibilidad y podrá generar costos adicionales.\n"
        "• Servicios adicionales: Cualquier modificación, ampliación de horario, personal, equipos o requerimiento no incluido en la propuesta será cotizado y cobrado por separado.\n"
        "• Montaje y logística: Transporte, permisos, parqueaderos, alimentación, hospedaje y demás gastos no especificados en la propuesta serán asumidos por el cliente.\n"
        "• Equipos: El cliente será responsable por daños o pérdidas ocasionados a equipos, materiales o elementos de F3R durante la prestación del servicio por causas atribuibles al cliente o terceros bajo su responsabilidad.\n"
        "• Lugar del evento: El cliente deberá garantizar las condiciones de acceso, seguridad, espacio, energía y demás requerimientos necesarios para la ejecución del servicio.\n"
        "• Fuerza mayor: F3R Producciones no será responsable por retrasos o incumplimientos derivados de hechos imprevisibles o fuera de su control.\n"
        "• Responsabilidad: F3R responderá exclusivamente por los servicios expresamente contratados y no asumirá costos o responsabilidades derivados de servicios no incluidos en la propuesta.\n"
        "• Aceptación: La aprobación de la cotización, firma, pago del anticipo o confirmación escrita del servicio implica la aceptación de estas condiciones comerciales.\n\n"
        "F3R Producciones se reserva el derecho de modificar disponibilidad, precios y condiciones antes de la confirmación formal del servicio."
    )
    pdf.multi_cell(192, 3.2, terminos_f3r, border=1)
    
    # 5. Bloque de Firmas Formales
    pdf.ln(6)
    if pdf.get_y() > 235:
        pdf.add_page()
        
    y_firmas = pdf.get_y() + 6
    pdf.set_draw_color(212, 175, 55) # Línea dorada de firma F3R
    pdf.set_line_width(0.5)
    pdf.line(22, y_firmas, 98, y_firmas)
    pdf.line(118, y_firmas, 194, y_firmas)
    pdf.set_line_width(0.2)
    
    pdf.set_xy(22, y_firmas + 2)
    pdf.set_font(font, 'B', 9)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(76, 4.2, "Francisco Ricaurte", 0, align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(22)
    pdf.set_font(font, '', 8)
    pdf.set_text_color(71, 85, 105)
    pdf.cell(76, 3.8, "Gerente", 0, align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(22)
    pdf.cell(76, 3.8, "WhatsApp +57 333 434 0735", 0, align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(22)
    pdf.cell(76, 3.8, "f3rproducciones@gmail.com", 0, align='C')
    
    pdf.set_xy(118, y_firmas + 2)
    pdf.set_font(font, 'B', 9)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(76, 4.2, f"ACEPTADO: {evt.get('client', 'CLIENTE')[:34]}", 0, align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(118)
    pdf.set_font(font, '', 8)
    pdf.set_text_color(71, 85, 105)
    pdf.cell(76, 3.8, f"NIT / C.C.: {evt.get('nit', '800.141.648-9')}", 0, align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(118)
    pdf.cell(76, 3.8, "Firma Autorizada / Aceptación", 0, align='C')

    return bytes(pdf.output())

if __name__ == "__main__":
    with open("crm_events_db.json", encoding="utf-8") as f:
        d = json.load(f)
    pdf_bytes = build_pdf(d["events"][0])
    with open("COTIZACION_PRUEBA_ENAP.pdf", "wb") as f_out:
        f_out.write(pdf_bytes)
    print("SUCCESS! PDF regenerated and validated! Size:", len(pdf_bytes), "bytes")
