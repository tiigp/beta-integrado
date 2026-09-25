import importlib.util
import os

project = r"C:\Users\PC\OneDrive\Desktop\BETA INTEGRADO\Control de combustible"
os.chdir(project)
spec = importlib.util.spec_from_file_location("app_module", os.path.join(project, "app.py"))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
app = mod.app
client = app.test_client()
for path in ["/rrhh", "/admin", "/despachos", "/reportes", "/rrhh/notificaciones"]:
    response = client.get(path, follow_redirects=True)
    print(path, response.status_code, response.mimetype)
    if response.status_code >= 400:
        print(response.get_data(as_text=True)[:500])
