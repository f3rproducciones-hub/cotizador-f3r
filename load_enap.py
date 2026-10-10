import glob, openpyxl, json, os

f = [x for x in glob.glob('*Family Day ENAP FINAL.xlsx') if not x.startswith('~$')][0]
wb = openpyxl.load_workbook(f, data_only=True)
ws = wb.active

items = []
current_cat = None
for r in range(12, 66):
    c1 = ws.cell(r, 1).value
    c2 = ws.cell(r, 2).value
    c3 = ws.cell(r, 3).value
    c4 = ws.cell(r, 4).value
    c5 = ws.cell(r, 5).value
    
    if str(c1).strip() == 'Item' and c2:
        current_cat = str(c2).strip()
        continue
    
    if isinstance(c1, (int, float)) and c2 and c3 and c4:
        items.append({
            'categoria': current_cat or 'General',
            'item': str(c2).strip(),
            'cantidad': int(c3),
            'costo_unit': 0,
            'precio_unit': float(c4)
        })

costs = [
    {'categoria': 'Producción & Escenario', 'item': 'RV (Sonido, Luces, Estructura)', 'costo_real': 38000000, 'factura_soporte': 'RV-PROD'},
    {'categoria': 'Personal Técnico', 'item': 'Omar (Logística / Técnico)', 'costo_real': 2500000, 'factura_soporte': 'PAG-OMAR'},
    {'categoria': 'Estructura & Carpas', 'item': 'Aba Carpas (Carpas Hangar y 4x4)', 'costo_real': 8650000, 'factura_soporte': 'ABA-CARP'},
    {'categoria': 'Zona Infantil', 'item': 'Inflables y Juegos', 'costo_real': 1900000, 'factura_soporte': 'INF-RECR'},
    {'categoria': 'Audiovisuales', 'item': "TV 50' + PS5 (6 combos)", 'costo_real': 1400000, 'factura_soporte': 'GAMING-01'},
    {'categoria': 'Personal Operativo', 'item': 'Ricardo (Coordinación)', 'costo_real': 1000000, 'factura_soporte': 'PAG-RIC'},
    {'categoria': 'Compras & Insumos', 'item': 'Compras Varias', 'costo_real': 1000000, 'factura_soporte': 'DIR-COMP'},
    {'categoria': 'Imprevistos', 'item': 'Gastos Varios Operativos', 'costo_real': 1000000, 'factura_soporte': 'GAS-VAR'}
]

total_cost = sum(c['costo_real'] for c in costs)
total_sale = sum(it['cantidad'] * it['precio_unit'] for it in items)
factor = total_cost / total_sale if total_sale > 0 else 0.88

for it in items:
    it['costo_unit'] = round(it['precio_unit'] * factor)

enap_event = {
    "id": "C-V1-9001",
    "client": "ESCUELA NAVAL DE CADETES ALMIRANTE PADILLA - ENAP",
    "nit": "800.141.648-9",
    "event_name": "Family Day ENAP 2026",
    "contact_person": "Comando Base Naval / Dpto Bienestar",
    "contact_email": "bienestar@enap.edu.co",
    "contact_phone": "300 815 0085",
    "event_date": "2026-10-18",
    "stage": "Ganado / Montaje",
    "prob_win": 100,
    "location": "Campo de Paradas y Zonas Verdes ENAP - Manzanillo",
    "pax": 1200,
    "created_at": "2026-09-17",
    "provider_info": {
        "razon_social": "F3R Producciones",
        "nit": "1048935287",
        "direccion": "Conjunto Residencial verona torre 4 - 316",
        "contacto": "Francisco Ricaurte R.",
        "email": "f3rproducciones@gmail.com",
        "telefono": "300 815 0085",
        "forma_pago": "Contado / 50% anticipo",
        "banco": "Bancolombia Cuenta Ahorros 509-637939-92"
    },
    "notes": "Producción Integral Family Day ENAP: Sonido Line Array DAS AERO 40 con relevos, Iluminación Beam/Par Led/Feria 500m, Ground Support 12.8x9.8m, Pantallas Led 3.9, Generadores 100kva/75kva, Carpas Hangar y Zona Infantil inflables.",
    "items": items,
    "actual_costs": costs
}

db = {"events": [enap_event]}
with open("crm_events_db.json", "w", encoding="utf-8") as out_f:
    json.dump(db, out_f, ensure_ascii=False, indent=2)

print(f"ENAP base loaded! Items: {len(items)}, Total sale: {total_sale}, Total cost: {total_cost}, Pre-utility: {total_sale - total_cost}")
