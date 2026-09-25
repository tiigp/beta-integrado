import sqlite3

p = r"Control de combustible/instance/combustible.db"
rows = [
    ("2026-08-28", 2, 35.5, ""),
    ("2026-08-31", 2, 28.0, ""),
    ("2026-09-02", 2, 32.0, ""),
    ("2026-09-04", 2, 31.0, ""),
    ("2026-08-27", 1, 32.0, "Registro recuperado del historico diario"),
    ("2026-08-31", 1, 15.0, "Registro recuperado del historico diario"),
    ("2026-09-02", 1, 22.0, "Registro recuperado del historico diario"),
    ("2026-09-04", 1, 22.0, "Registro recuperado del historico diario"),
    ("2026-08-28", 3, 141.0, "Registro recuperado del historico diario"),
    ("2026-08-31", 3, 96.0, "Registro recuperado del historico diario"),
    ("2026-09-02", 3, 73.0, "Registro recuperado del historico diario"),
    ("2026-09-04", 3, 73.0, "Registro recuperado del historico diario"),
    ("2026-08-24", 4, 15.0, "Registro recuperado del historico diario"),
    ("2026-08-27", 4, 20.0, "Registro recuperado del historico diario"),
    ("2026-08-31", 4, 12.0, "Registro recuperado del historico diario"),
    ("2026-09-04", 4, 11.0, "Registro recuperado del historico diario"),
    ("2026-08-31", 5, 66.0, "Registro recuperado del historico diario"),
    ("2026-08-26", 6, 57.0, "Registro recuperado del historico diario"),
    ("2026-08-31", 6, 65.0, "Registro recuperado del historico diario"),
    ("2026-08-27", 7, 13.0, "Registro recuperado del historico diario"),
    ("2026-09-04", 7, 20.0, "Registro recuperado del historico diario"),
    ("2026-09-07", 6, 30.0, ""),
    ("2026-09-07", 2, 41.0, ""),
    ("2026-09-07", 1, 30.0, ""),
]

connection = sqlite3.connect(p)
existing = connection.execute("SELECT COUNT(*), COALESCE(MAX(id), 0) FROM despacho").fetchone()
assert existing == (20, 20), existing
connection.executemany(
    "INSERT INTO despacho (fecha, equipo_id, litros, responsable, nombre_chofer, hora, horometro, kilometraje, tipo_carga, observacion) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [(date, equipo_id, litros, "Edgar Nuñez", "", 0.0, 0.0, 0.0, "Media", observation) for date, equipo_id, litros, observation in rows],
)
connection.commit()
print(connection.execute("SELECT COUNT(*), MIN(id), MAX(id) FROM despacho").fetchone())
