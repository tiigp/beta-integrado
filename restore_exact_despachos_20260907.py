import sqlite3

path = r"Control de combustible/instance/combustible.db"
responsable = "Edgar Nu" + chr(241) + "ez"
diego = "Diego Gonz" + chr(225) + "lez"
rows = [
    (3, "2026-08-13", 83, "Diosnel Chamorro", 0, 439, 0, "Media", "Primera Carga"),
    (3, "2026-08-17", 85, "Diosnel Chamorro", 0, 453.6, 0, "Media", "Segunda Carga"),
    (3, "2026-08-19", 126, "Diosnel Chamorro", 471, 0, 0, "Media", ""),
    (3, "2026-08-21", 113, "Diosnel Chamorro", 0, 487.1, 0, "Media", ""),
    (3, "2026-08-24", 139, "Diosnel Chamorro", 0, 501, 0, "Media", ""),
    (3, "2026-08-28", 141, "", 0, 523.2, 0, "Media", "Registro recuperado del historico diario"),
    (3, "2026-08-31", 96, "", 0, 540.5, 0, "Media", "Registro recuperado del historico diario"),
    (3, "2026-09-02", 73, "", 0, 554.9, 0, "Media", "Registro recuperado del historico diario"),
    (3, "2026-09-04", 73, "", 0, 569.3, 0, "Media", "Registro recuperado del historico diario"),
    (5, "2026-08-17", 70, "David Barboza", 0, 0, 3278975, "Media", "Primera Carga"),
    (5, "2026-08-31", 66, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (1, "2026-08-13", 2, "Heriberto Ortiz", 4870.09, 0, 0, "Media", "Primera carga"),
    (1, "2026-08-17", 43, "Heriberto Ortiz", 0, 48908, 0, "Media", "Segunda Carga"),
    (1, "2026-08-19", 40, "Gerardo Ramirez", 0, 4910.8, 0, "Media", ""),
    (1, "2026-08-21", 34, "Gerardo Ramirez", 4926, 0, 0, "Media", ""),
    (1, "2026-08-24", 29, "Gerardo Ramirez", 0, 4947.2, 0, "Media", ""),
    (1, "2026-08-27", 32, "", 0, 4970.1, 0, "Media", "Registro recuperado del historico diario"),
    (1, "2026-08-31", 15, "", 0, 4979.4, 0, "Media", "Registro recuperado del historico diario"),
    (1, "2026-09-02", 22, "", 0, 4993.7, 0, "Media", "Registro recuperado del historico diario"),
    (1, "2026-09-04", 22, "", 0, 5009.9, 0, "Media", "Registro recuperado del historico diario"),
    (1, "2026-09-07", 30, "Gerardo Ramirez", 0, 5025.4, 0, "Media", ""),
    (2, "2026-08-13", 50, "Dario Fleitas", 4870.09, 0, 0, "Media", "Primera Carga"),
    (2, "2026-08-17", 43, "Heriberto Ortiz", 0, 685.9, 0, "Media", "Segunda Carga"),
    (2, "2026-08-19", 47.5, "Heriberto Ortiz", 720, 0, 0, "Media", ""),
    (2, "2026-08-21", 37, "Heriberto Ortiz", 748.3, 0, 0, "Media", ""),
    (2, "2026-08-28", 35.5, "Heriberto Ortiz", 0, 824, 0, "Media", ""),
    (2, "2026-08-31", 28, "Dario Fleitas", 0, 847.1, 0, "Media", ""),
    (2, "2026-09-02", 32, "Dario Fleitas", 0, 875.8, 0, "Media", ""),
    (2, "2026-09-04", 31, "Dario Fleitas", 0, 902.28, 0, "Media", ""),
    (2, "2026-09-07", 41, "Pedro Gimenez", 0, 934.8, 0, "Media", ""),
    (6, "2026-08-21", 64, "Osmar Fernandez", 0, 0, 0, "Media", ""),
    (6, "2026-08-26", 57, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (6, "2026-08-31", 65, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (6, "2026-09-07", 30, "Osmar Fernandez", 0, 0, 0, "Media", ""),
    (4, "2026-08-14", 10, "David Barboza", 21215.1, 0, 0, "Media", "Primera Carga Tractor"),
    (4, "2026-08-17", 11, "Dario Fleitas", 0, 0, 0, "Media", ""),
    (4, "2026-08-21", 10, diego, 0, 0, 0, "Media", ""),
    (4, "2026-08-24", 15, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (4, "2026-08-27", 20, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (4, "2026-08-31", 12, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (4, "2026-09-04", 11, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (7, "2026-08-25", 20, "Gerardo Ramirez", 0, 0, 0, "Baja", "Retiro para uso en taller."),
    (7, "2026-08-27", 13, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
    (7, "2026-09-04", 20, "", 0, 0, 0, "Media", "Registro recuperado del historico diario"),
]

connection = sqlite3.connect(path)
connection.execute("BEGIN")
assert connection.execute("SELECT COUNT(*) FROM despacho").fetchone()[0] == 44
for equipo_id, fecha, litros, chofer, hora, horometro, kilometraje, tipo_carga, observacion in rows:
    matches = connection.execute(
        "SELECT id FROM despacho WHERE equipo_id = ? AND date(fecha) = ? AND litros = ?",
        (equipo_id, fecha, litros),
    ).fetchall()
    if not matches and equipo_id == 6 and fecha == "2026-08-21" and litros == 64:
        matches = connection.execute(
            "SELECT id FROM despacho WHERE equipo_id = ? AND date(fecha) = ? AND litros = ?",
            (equipo_id, fecha, 64.5),
        ).fetchall()
    assert len(matches) == 1, (equipo_id, fecha, litros, matches)
    connection.execute(
        "UPDATE despacho SET responsable = ?, nombre_chofer = ?, hora = ?, horometro = ?, kilometraje = ?, tipo_carga = ?, observacion = ? WHERE id = ?",
        (responsable, chofer, hora, horometro, kilometraje, tipo_carga, observacion, matches[0][0]),
    )
connection.commit()
print("actualizados=", len(rows), "despachos=", connection.execute("SELECT COUNT(*) FROM despacho").fetchone()[0])
