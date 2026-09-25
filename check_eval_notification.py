import os
import sys
import tempfile

os.environ["DATABASE_URL"] = "sqlite:///" + os.path.join(tempfile.gettempdir(), "control_combustible_notification_test.sqlite").replace("\\", "/")
sys.path.insert(0, r"Control de combustible")
from app import app, db, User, Colaborador, Notificacion, generate_password_hash, init_db

with app.app_context():
    db.drop_all()
    db.create_all()
    init_db()
    rrhh = User(username="rrhh_test", password_hash=generate_password_hash("rrhh123"), full_name="RRHH Test", role="rrhh")
    supervisor = User(username="supervisor_test", password_hash=generate_password_hash("sup123"), full_name="Supervisor Test", role="supervisor", correo="sup@example.com")
    db.session.add_all([rrhh, supervisor])
    db.session.flush()
    colaborador = Colaborador(nombre="Colaborador Test", legajo="T-001", supervisor_id=supervisor.id)
    db.session.add(colaborador)
    db.session.commit()
    colaborador_id = colaborador.id
    supervisor_id = supervisor.id

with app.test_client() as client:
    login = client.post("/login", data={"username": "rrhh_test", "password": "rrhh123"}, follow_redirects=False)
    response = client.post("/rrhh/evaluaciones/asignar", data={"colaborador_id": colaborador_id, "supervisor_id": supervisor_id, "formulario": "periodica", "canal_notificacion": "correo"}, follow_redirects=False)
    with app.app_context():
        notification = Notificacion.query.filter_by(usuario_id=supervisor_id).first()
        print("login=", login.status_code, "assignment=", response.status_code, "outlook=", response.headers.get("Location", "").startswith("https://outlook.office.com/"), "notification=", bool(notification))
