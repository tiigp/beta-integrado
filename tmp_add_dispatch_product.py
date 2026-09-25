import os
import shutil
import sqlite3

DB_PATH = r"C:\Users\PC\OneDrive\Desktop\BETA INTEGRADO\Control de combustible\instance\combustible.db"
BACKUP_PATH = r"C:\Users\PC\OneDrive\Desktop\BETA INTEGRADO\Control de combustible\instance\combustible.db.before_add_product_20260911.bak"

if not os.path.exists(BACKUP_PATH):
    shutil.copy2(DB_PATH, BACKUP_PATH)

con = sqlite3.connect(DB_PATH)
cur = con.cursor()

target = cur.execute(
    "SELECT id, numero FROM despacho_acero WHERE numero LIKE 'D-20260911%' ORDER BY id DESC LIMIT 1"
).fetchone()
if target is None:
    print('NO_TARGET_DISPATCH')
    con.close()
    raise SystemExit(0)

despacho_id, numero = target
label = cur.execute(
    "SELECT id FROM etiqueta_calidad WHERE no_conforme = 0 AND id NOT IN (SELECT etiqueta_id FROM detalle_despacho_acero) ORDER BY id DESC LIMIT 1"
).fetchone()
if label is None:
    print('NO_AVAILABLE_LABELS')
    con.close()
    raise SystemExit(0)

label_id = label[0]
count_before = cur.execute("SELECT COUNT(*) FROM detalle_despacho_acero WHERE despacho_id = ?", (despacho_id,)).fetchone()[0]
cur.execute(
    "INSERT INTO detalle_despacho_acero (despacho_id, etiqueta_id) VALUES (?, ?)",
    (despacho_id, label_id),
)
con.commit()
count_after = cur.execute("SELECT COUNT(*) FROM detalle_despacho_acero WHERE despacho_id = ?", (despacho_id,)).fetchone()[0]
print({
    'despacho_id': despacho_id,
    'numero': numero,
    'etiqueta_id': label_id,
    'count_before': count_before,
    'count_after': count_after,
    'backup': BACKUP_PATH,
})
con.close()
