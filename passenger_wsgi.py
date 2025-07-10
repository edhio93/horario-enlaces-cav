import os, subprocess, sys

# Activa el virtualenv (ajusta la ruta si tuvenv no está aquí)
activate = os.path.join(os.path.dirname(__file__), "venv/bin/activate_this.py")
with open(activate) as f:
    exec(f.read(), {"__file__": activate})

# Lanza Streamlit en el puerto que Passenger asigne
cmd = [
    sys.executable, "-m", "streamlit", "run", "schedule_app.py",
    "--server.address", "0.0.0.0",
    "--server.port", os.environ.get("PASSENGER_APP_PORT", "8501")
]
subprocess.Popen(cmd)

# Dummy WSGI para Passenger
def application(environ, start_response):
    start_response("200 OK", [("Content-Type","text/plain")])
    return [b"Arrancando Streamlit..."]
