import glob, openpyxl, json, os

candidates = (
    glob.glob('*Family Day ENAP FINAL.xlsx') + 
    glob.glob('docs/*Family Day ENAP FINAL.xlsx') + 
    glob.glob('docs/*/*Family Day ENAP FINAL.xlsx')
)
valid_candidates = [x for x in candidates if not os.path.basename(x).startswith('~$')]
if not valid_candidates:
    raise FileNotFoundError("No se encontró el archivo 'Family Day ENAP FINAL.xlsx' ni en la raíz ni en docs/")
f = valid_candidates[0]
wb = openpyxl.load_workbook(f, data_only=True)
ws = wb.active

items_enap = []
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
        items_enap.append({
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
total_sale = sum(it['cantidad'] * it['precio_unit'] for it in items_enap)
factor = total_cost / total_sale if total_sale > 0 else 0.88

for it in items_enap:
    it['costo_unit'] = round(it['precio_unit'] * factor)

new_prov = {
    'razon_social': 'F3R PRODUCCIONES S.A.S.',
    'nit': '902106670-3',
    'telefono': '+57 333 434 0735',
    'instagram': '@f3rproducciones',
    'email': 'f3rproducciones@gmail.com',
    'direccion': 'Cartagena de Indias Diagonal 31 80B75 Of 501',
    'contacto': 'Francisco Ricaurte',
    'cargo': 'Gerente',
    'forma_pago': 'Contado / 50% anticipo',
    'banco': 'Bancolombia Cuenta Ahorros 509-637939-92'
}

enap_evt = {
    'id': 'V1-1009',
    'client': 'ESCUELA NAVAL DE CADETES ALMIRANTE PADILLA - ENAP',
    'nit': '800.141.648-9',
    'event_name': 'Family Day ENAP 2026',
    'contact_person': 'Comando Base Naval / Dpto Bienestar',
    'contact_email': 'bienestar@enap.edu.co',
    'contact_phone': '+57 333 434 0735',
    'event_date': '2026-10-18',
    'stage': 'Ganado / Montaje',
    'prob_win': 100,
    'location': 'Campo de Paradas y Zonas Verdes ENAP - Manzanillo',
    'pax': 1200,
    'created_at': '2026-10-09',
    'version': 1,
    'provider_info': new_prov,
    'notes': 'Producción Integral Family Day ENAP: Sonido DAS Aero 40, Luces, Tarima Ground Support, Pantalla LED, Generadores y Carpas.',
    'items': items_enap,
    'actual_costs': costs
}

with open('crm_events_db.json', 'r', encoding='utf-8') as f:
    db = json.load(f)

for ev in db.get('events', []):
    ev['version'] = ev.get('version', 1)
    ev['provider_info'] = new_prov

if not any('ENAP' in e.get('client', '') for e in db['events']):
    db['events'].insert(0, enap_evt)

with open('crm_events_db.json', 'w', encoding='utf-8') as f:
    json.dump(db, f, ensure_ascii=False, indent=2)

print('CRM Database synced with versioning and official provider!')
