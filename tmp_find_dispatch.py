import sqlite3
from pathlib import Path

root = Path('C:/Users/PC/OneDrive/Desktop/BETA INTEGRADO')
for db in sorted(root.glob('**/*.db')):
    try:
        con = sqlite3.connect(str(db))
        cur = con.cursor()
        rows = cur.execute("SELECT numero, id, cliente, destino, fecha FROM despacho_acero WHERE numero LIKE '%20260911%' ORDER BY id").fetchall()
        if rows:
            print('DB=', db)
            for r in rows:
                print('  ', r)
    except Exception as e:
        print('ERR', db, e)
    finally:
        try:
            con.close()
        except Exception:
            pass
