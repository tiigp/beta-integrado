import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(r"c:\Users\PC\OneDrive\Desktop\BETA INTEGRADO\Control de combustible")))
import app as a

conn = sqlite3.connect(r"c:\Users\PC\OneDrive\Desktop\BETA INTEGRADO\Control de combustible\instance\combustible.db")
row = conn.execute("select id, nombre_archivo from documento_rrhh where nombre_archivo like '%.xls%' limit 1").fetchone()
print('ROW', row)
conn.close()

app = a.app
app.config['TESTING'] = True
client = app.test_client()
resp = client.get(f'/rrhh/documentos/{row[0]}/archivo')
print('archive_status', resp.status_code)
print('archive_location', resp.headers.get('Location'))
if resp.headers.get('Location'):
    resp2 = client.get(resp.headers.get('Location'))
    print('edit_status', resp2.status_code)
    text = resp2.get_data(as_text=True)
    print('contains_textarea', 'textarea' in text)
    print(text[:600])
