import pandas as pd
from sqlalchemy import create_engine

# Lee tu Excel
df = pd.read_excel("Recursos.xlsx", sheet_name="Reservas", dtype=str).fillna("")
# Ajusta nombres de columnas a minúsculas y sin espacios, si quieres
df = df.rename(columns={
    "Fecha": "fecha",
    "Hora inicio": "hora_inicio",
    "Hora fin": "hora_fin",
    "Profesor": "profesor",
    "Curso": "curso",
    "Recurso": "recurso",
    "Observaciones": "observaciones"
})

# Conecta a Supabase (pon aquí tu URL de Connection String)
engine = create_engine("postgres://usuario:contraseña@host:5432/nombre_bd")
# Escribe la tabla (reemplaza si ya existiera)
df.to_sql("reservas", engine, if_exists="replace", index=False)
print("Migración completada.")