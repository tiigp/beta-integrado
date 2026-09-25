import sqlite3

path = "C:/Users/PC/OneDrive/Desktop/BETA INTEGRADO/Control de combustible/instance/combustible.db"
con = sqlite3.connect(path)
cur = con.cursor()
print('SCHEMA_DESPACHO_ACERO')
print(cur.execute('PRAGMA table_info(despacho_acero)').fetchall())
print('SCHEMA_DETALLE')
print(cur.execute('PRAGMA table_info(detalle_despacho_acero)').fetchall())
print('SCHEMA_ETIQUETA')
print(cur.execute('PRAGMA table_info(etiqueta_calidad)').fetchall())
print('DESPACHO_TARGET')
print(cur.execute("SELECT * FROM despacho_acero WHERE numero = 'D-20260911-0002' OR numero = 'D-20260911-002' ORDER BY id").fetchall())
print('ETIQUETAS_DISPONIBLES')
print(cur.execute("SELECT id, tipo_producto, medida, longitud, peso, colada, no_conforme FROM etiqueta_calidad WHERE no_conforme = 0 AND id NOT IN (SELECT etiqueta_id FROM detalle_despacho_acero) ORDER BY id DESC LIMIT 20").fetchall())
con.close()
