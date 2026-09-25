from __future__ import annotations

import importlib.util
import pathlib
import shutil
from datetime import datetime

BASE_DIR = pathlib.Path(r"C:\Users\PC\OneDrive\Desktop\BETA INTEGRADO")
APP_PATH = BASE_DIR / "Control de combustible" / "app.py"
DB_PATH = BASE_DIR / "Control de combustible" / "instance" / "combustible.db"
BACKUP_PATH = BASE_DIR / "Control de combustible" / "instance" / f"combustible.db.before_remove_orphans_{datetime.now():%Y%m%d_%H%M%S}.bak"

spec = importlib.util.spec_from_file_location("rrhh_app", APP_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"No se pudo cargar la aplicación desde {APP_PATH}")

module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


if DB_PATH.exists():
    shutil.copy2(DB_PATH, BACKUP_PATH)
    print(f"Backup creado: {BACKUP_PATH}")
else:
    raise FileNotFoundError(f"No se encontró la base de datos: {DB_PATH}")

with module.app.app_context():
    rows = module.DocumentoRRHH.query.filter(module.DocumentoRRHH.carpeta_id.is_(None)).all()
    print(f"Documentos huérfanos encontrados: {len(rows)}")

    for row in rows:
        target = (module.RRHH_DOCUMENTS_DIR / row.nombre_archivo).resolve()
        if target.parent == module.RRHH_DOCUMENTS_DIR.resolve() and target.is_file():
            target.unlink()
            print(f"Archivo eliminado: {target}")
        else:
            print(f"Archivo no encontrado, se elimina solo el registro: {row.nombre_archivo}")

        module.db.session.delete(row)

    module.db.session.commit()
    print("Registros huérfanos eliminados de la base de datos.")
