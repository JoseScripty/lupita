
import requests
from requests.auth import HTTPDigestAuth
from datetime import datetime
from time import mktime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Códigos de AttendanceState según el Excel exportado desde el equipo.
ESTADOS_ASISTENCIA = {
    "0": "-",
    "1": "Ingreso",
    "2": "Escape",
    "3": "Intrusión",
    "4": "Salida",
}

ESTADOS_VALIDOS = {
    "Ingreso",
    "Salida",
    "Intrusión",
    "Escape",
}


IP = "192.168.3.1"
USUARIO = "admin"
CLAVE = "qwerty1234"

# Rango: 22 de octubre de 2026
inicio = int(mktime(datetime(2026, 10, 22, 0, 0, 0).timetuple()))
fin = int(mktime(datetime(2026, 10, 23, 0, 0, 0).timetuple()))

url = f"http://{IP}/cgi-bin/recordFinder.cgi"

params = {
    "action": "find",
    "name": "AccessControlCardRec",
    "StartTime": inicio,
    "EndTime": fin,
    "count": 100,
}

respuesta = requests.get(
    url,
    params=params,
    auth=HTTPDigestAuth(USUARIO, CLAVE),
    timeout=30,
)
respuesta.raise_for_status()

# Convierte las líneas records[0].Campo=valor
# en una lista de diccionarios.
registros = {}

for linea in respuesta.text.splitlines():
    if "=" not in linea:
        continue

    clave, valor = linea.split("=", 1)

    if clave == "found":
        continue

    if not clave.startswith("records["):
        continue

    indice = int(clave.split("[", 1)[1].split("]", 1)[0])
    campo = clave.split("].", 1)[1]

    registros.setdefault(indice, {})[campo] = valor

wb = Workbook()
ws = wb.active
ws.title = "Asistencia"

columnas = [
    "Método de Desbloqueo",
    "Mascarilla facial",
    "ID",
    "Nombre",
    "Estado",
    "Motivo de la falla",
    "Estado de Asistencia",
    "Tiempo",
]

ws.append(columnas)

escritos = 0

for registro in registros.values():
    attendance = ESTADOS_ASISTENCIA.get(
        registro.get("AttendanceState", "").strip(), ""
    )
    if attendance not in ESTADOS_VALIDOS:
        continue

    try:
        timestamp = int(
            registro.get("CreateTime", "0")
        )
        fecha = datetime.fromtimestamp(timestamp)

        metodo = registro.get("Method", "")
        mascara = registro.get("Mask", "")
        user_id = registro.get("UserID", "")
        nombre = registro.get("CardName", "")
        status = registro.get("Status", "")
        error = registro.get("ErrorCode", "")

        # Dejamos los códigos como texto hasta validar
        # su equivalencia exacta con la interfaz.
        ws.append([
            metodo,
            mascara,
            user_id,
            nombre,
            status,
            error,
            attendance,
            fecha.strftime("%d-%m-%Y %H:%M:%S"),
        ])
        escritos += 1

    except (ValueError, OSError) as e:
        print("Registro omitido:", e)

# Formato básico
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

for celda in ws[1]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )
    celda.alignment = Alignment(horizontal="center")

for columna in ws.columns:
    letra = columna[0].column_letter
    ancho = max(
        len(str(c.value or ""))
        for c in columna
    )
    ws.column_dimensions[letra].width = min(ancho + 3, 32)

archivo = "registros_dahua.xlsx"
wb.save(archivo)

print(f"Registros recibidos: {len(registros)}")
print(f"Registros en el Excel: {escritos}")
print(f"Excel generado: {archivo}")
