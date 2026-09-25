import os
import sys
import tempfile

os.environ["DATABASE_URL"] = "sqlite:///" + os.path.join(tempfile.gettempdir(), "control_combustible_completion_test.sqlite").replace("\\", "/")
sys.path.insert(0, r"Control de combustible")
from app import app, db, User, Colaborador, AsignacionEvaluacion, Notificacion, EVALUATION_FORMS, generate_password_hash, init_db

with app.app_context():
    db.drop_all(); db.create_all(); init_db()
    rrhh = User(username="rrhh_completion", password_hash=generate_password_hash("rrhh123"), full_name="RRHH", role="rrhh")
    supervisor = User(username="supervisor_completion", password_hash=generate_password_hash("sup123"), full_name="Supervisor", role="supervisor")
    db.session.add_all([rrhh, supervisor]); db.session.flush()
    collaborator = Colaborador(nombre="Persona Evaluada", legajo="T-100", supervisor_id=None)
    db.session.add(collaborator); db.session.flush()
    assignment = AsignacionEvaluacion(colaborador_id=collaborator.id, supervisor_id=supervisor.id, formulario="periodica", creado_por_id=rrhh.id)
    db.session.add(assignment); db.session.commit()
    ids = rrhh.id, supervisor.id, collaborator.id, assignment.id

criteria_count = sum(len(items) for _, items in EVALUATION_FORMS["periodica"]["groups"])
with app.test_client() as client:
    with client.session_transaction() as session:
        session["_user_id"] = str(ids[1]); session["_fresh"] = True
    data = {"colaborador_id": str(ids[2]), "formulario": "periodica", "asignacion_id": str(ids[3]), "fecha_evaluacion": "2026-09-07", "estado": "Completada"}
    data.update({f"puntaje_{index}": "5" for index in range(criteria_count)})
    response = client.post("/rrhh/evaluaciones?legajo=T-100&formulario=periodica&asignacion_id=1&popup=1", data=data, follow_redirects=False)
    with app.app_context():
        notification = Notificacion.query.filter(Notificacion.usuario_id == ids[0], Notificacion.titulo.like("Evaluación realizada:%")).first()
        assignment = db.session.get(AsignacionEvaluacion, ids[3])
        print("status=", response.status_code, "assignment=", assignment.estado, "rrhh_message=", bool(notification))
