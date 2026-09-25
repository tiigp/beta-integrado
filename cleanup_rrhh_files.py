from pathlib import Path

base = Path('C:/Users/PC/OneDrive/Desktop/BETA INTEGRADO/Control de combustible/instance/documentos_rrhh')
print('EXISTS', base.exists())
files = sorted(base.iterdir())
print('BEFORE', len(files))
for f in files:
    print('REMOVE', f.name)
    if f.is_file():
        f.unlink()
print('AFTER', len(list(base.iterdir())))
