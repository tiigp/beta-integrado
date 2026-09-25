import sqlite3

DB = r"C:\Users\PC\OneDrive\Desktop\BETA INTEGRADO\Control de combustible\instance\combustible.db"
con = sqlite3.connect(DB)
cur = con.cursor()
print('CHECK_NUMERO')
rows = cur.execute("SELECT id, numero, cliente, destino, chofer, cedula, chapa, responsable, fecha FROM despacho_acero WHERE numero LIKE '%20260911%' ORDER BY id").fetchall()
print(rows)
print('LAST_DISPATCHES')
print(cur.execute("SELECT id, numero, cliente, destino, chofer, cedula, chapa, responsable, fecha FROM despacho_acero ORDER BY id DESC LIMIT 10").fetchall())
print('SCHEMA_DESPACHO')
print(cur.execute('PRAGMA table_info(despacho_acero)').fetchall())
print('SCHEMA_DETALLE')
print(cur.execute('PRAGMA table_info(detalle_despacho_acero)').fetchall())
print('SCHEMA_ETIQUETA')
print(cur.execute('PRAGMA table_info(etiqueta_calidad)').fetchall())
print('AVAILABLE_LABELS')
print(cur.execute("SELECT e.id, e.numero, e.tipo_producto, e.medida, e.longitud, e.peso, e.colada, e.no_conforme FROM etiqueta_calidad e WHERE e.no_conforme = 0 AND NOT EXISTS (SELECT 1 FROM detalle_despacho_acero d WHERE d.etiqueta_id = e.id) ORDER BY e.id DESC LIMIT 20").fetchall())
con.close()
