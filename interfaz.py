"""Interfaz para consultar las marcas de asistencia del equipo Dahua.

Ejecutar con:  py interfaz.py
"""

import calendar
import copy
import json
import os
import sys
import threading
import tkinter as tk
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from time import mktime
from tkinter import filedialog, messagebox, ttk

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from requests.auth import HTTPDigestAuth

# --- Configuración ----------------------------------------------------------

IP = "192.168.3.1"
USUARIO = "admin"
CLAVE = "qwerty1234"
MAX_REGISTROS = 100  # registros que se piden al equipo en cada consulta

# Códigos de AttendanceState según el Excel exportado desde el equipo.
ESTADOS_ASISTENCIA = {
    "0": "-",
    "1": "Ingreso",
    "2": "Escape",
    "3": "Intrusión",
    "4": "Salida",
}
MARCAS = ["Ingreso", "Escape", "Intrusión", "Salida"]  # en el orden de la jornada
ESTADOS_VALIDOS = set(MARCAS)
METODOS = {"15": "Cara"}
RESULTADOS = {"1": "Exitoso", "0": "Fallido"}

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
DIAS_CORTOS = [d[:3].capitalize() for d in DIAS]  # Lun, Mar, Mié, ...
AZUL = "#1F4E78"
AYUDA_RESUMEN = (
    "Selecciona un usuario para ver sus observaciones y sus marcas. "
    "Rojo = faltan marcas (esas horas no se cuentan). "
    "Amarillo = marcas mal hechas que se corrigieron, conviene revisarlas."
)
AYUDA_SEMANA = AYUDA_RESUMEN + " ⚠ = ese día tiene observaciones."
AYUDA_CORREGIR = (
    "Doble clic en una marca para corregirla (ej.: Escape que debía ser Ingreso). "
    "Los cambios se guardan en esta PC."
)

# Cambios hechos a mano, junto al programa (o al .exe).
CARPETA = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
ARCHIVO_CORRECCIONES = CARPETA / "correcciones.json"
FORMATO_FECHA = "%Y-%m-%d %H:%M:%S"

# Emojis de Slack para el resumen copiado (mismos colores que la tabla).
ICONOS_SLACK = {
    "Completo": ":large_green_circle:",
    "Revisar": ":large_yellow_circle:",
    "Incompleto": ":red_circle:",
}

COLUMNAS_REGISTROS = [
    ("fecha", "Fecha", 90, "center"),
    ("hora", "Hora", 90, "center"),
    ("id", "ID", 60, "center"),
    ("nombre", "Nombre", 220, "w"),
    ("marca", "Estado de Asistencia", 150, "center"),
    ("metodo", "Método", 100, "center"),
    ("resultado", "Estado", 100, "center"),
    ("correccion", "Corrección", 120, "center"),
]

COLUMNAS_RESUMEN = [
    ("id", "ID", 50, "center"),
    ("nombre", "Nombre", 170, "w"),
    ("ingreso", "Ingreso", 75, "center"),
    ("salida", "Salida", 75, "center"),
    ("pausas", "Escapes", 70, "center"),
    ("fuera", "Tiempo fuera", 95, "center"),
    ("trabajado", "Horas laboradas", 115, "center"),
    ("estado", "Estado", 90, "center"),
    ("obs", "Observaciones", 380, "w"),
]


# --- Consulta al equipo -----------------------------------------------------

def inicio_del_dia(dia):
    return int(mktime(datetime(dia.year, dia.month, dia.day).timetuple()))


def consultar_registros(dia):
    """Devuelve todas las marcas del día, ordenadas por hora."""
    respuesta = requests.get(
        f"http://{IP}/cgi-bin/recordFinder.cgi",
        params={
            "action": "find",
            "name": "AccessControlCardRec",
            "StartTime": inicio_del_dia(dia),
            "EndTime": inicio_del_dia(dia + timedelta(days=1)),
            "count": MAX_REGISTROS,
        },
        auth=HTTPDigestAuth(USUARIO, CLAVE),
        timeout=30,
    )
    respuesta.raise_for_status()

    # Convierte las líneas records[0].Campo=valor en diccionarios.
    crudos = {}
    for linea in respuesta.text.splitlines():
        clave, igual, valor = linea.partition("=")
        if not igual or not clave.startswith("records["):
            continue
        indice, _, campo = clave[len("records["):].partition("].")
        if indice.isdigit() and campo:
            crudos.setdefault(int(indice), {})[campo] = valor.strip()

    registros = []
    for crudo in crudos.values():
        try:
            fecha = datetime.fromtimestamp(int(crudo.get("CreateTime", "")))
        except (ValueError, OSError):
            continue
        metodo = crudo.get("Method", "")
        resultado = crudo.get("Status", "")
        registros.append({
            "fecha": fecha,
            "id": crudo.get("UserID", ""),
            "nombre": crudo.get("CardName", ""),
            "marca": ESTADOS_ASISTENCIA.get(crudo.get("AttendanceState", ""), "-"),
            "metodo": METODOS.get(metodo, metodo),
            "resultado": RESULTADOS.get(resultado, resultado),
        })

    registros.sort(key=lambda r: r["fecha"])
    return registros


def mensaje_error(error):
    if isinstance(error, (requests.ConnectionError, requests.Timeout)):
        return (
            f"No se pudo conectar con el equipo ({IP}).\n"
            "Revisa que esté encendido y en la misma red."
        )
    respuesta = getattr(error, "response", None)
    if respuesta is not None and respuesta.status_code == 401:
        return "El equipo rechazó el usuario o la clave."
    return f"Ocurrió un error al consultar el equipo:\n{error}"


# --- Cambios hechos a mano --------------------------------------------------
# Las marcas del equipo no se modifican. Los cambios se guardan aparte, en
# ARCHIVO_CORRECCIONES, y se aplican al calcular:
#   "correcciones": {"<ID>|<fecha y hora>": "<marca correcta>"}
#   "agregadas":    [{"id", "nombre", "fecha", "marca"}]  (ej.: trabajo virtual)
#   "usuarios":     {"<ID>": "<nombre>"}  (para elegir a quién agregarle marcas)

def clave_registro(r):
    return f"{r['id']}|{r['fecha']:{FORMATO_FECHA}}"


def clave_agregada(a):
    return f"{a.get('id')}|{a.get('fecha')}"


def cargar_cambios():
    try:
        with open(ARCHIVO_CORRECCIONES, encoding="utf-8") as archivo:
            datos = json.load(archivo)
    except FileNotFoundError:
        datos = {}
    if not isinstance(datos, dict):
        raise ValueError("el archivo no tiene el formato esperado")
    return {
        "correcciones": dict(datos.get("correcciones", {})),
        "agregadas": list(datos.get("agregadas", [])),
        "usuarios": dict(datos.get("usuarios", {})),
    }


def guardar_cambios(cambios):
    temporal = ARCHIVO_CORRECCIONES.with_suffix(".tmp")
    temporal.write_text(
        json.dumps(cambios, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    os.replace(temporal, ARCHIVO_CORRECCIONES)  # así no queda a medio escribir


def aplicar_correcciones(registros, correcciones):
    """Copia de los registros con las marcas corregidas (guarda la original)."""
    corregidos = []
    for r in registros:
        nueva = correcciones.get(clave_registro(r))
        if nueva in ESTADOS_VALIDOS and nueva != r["marca"]:
            r = {**r, "marca": nueva, "original": r["marca"]}
        corregidos.append(r)
    return corregidos


def marcas_agregadas(agregadas):
    """Convierte las marcas agregadas a mano al mismo formato que las del equipo."""
    marcas = []
    for a in agregadas:
        try:
            fecha = datetime.strptime(a["fecha"], FORMATO_FECHA)
        except (KeyError, TypeError, ValueError):
            continue
        if a.get("marca") not in ESTADOS_VALIDOS:
            continue
        marcas.append({
            "fecha": fecha,
            "id": str(a.get("id", "")),
            "nombre": a.get("nombre", ""),
            "marca": a["marca"],
            "metodo": "Manual",
            "resultado": "Exitoso",
            "manual": True,
        })
    return marcas


# --- Cálculo de horas -------------------------------------------------------

def hora(fecha):
    return fecha.strftime("%H:%M")


def duracion(tiempo):
    minutos = int(tiempo.total_seconds() // 60)
    return f"{minutos // 60}:{minutos % 60:02d}"


def calcular_jornada(marcas):
    """Calcula la jornada de un usuario a partir de sus marcas en orden.

    La jornada esperada es: un Ingreso, luego 0..n pares Escape -> Intrusión,
    y al final una Salida. Solo se suman los tramos que tienen marca de inicio
    y de fin; si falta una marca, ese tramo no se cuenta y se indica.
    """
    trabajado = fuera = timedelta()
    estado = "sin_ingreso"  # sin_ingreso | dentro | pausa | terminado
    desde = None            # hora en que empezó el tramo actual
    ingreso = salida = None
    pausas = 0
    faltas, revisar, sin_estado = [], [], []

    for marca in marcas:
        t, tipo, h = marca["fecha"], marca["marca"], hora(marca["fecha"])

        if tipo not in ESTADOS_VALIDOS:
            sin_estado.append(h)
            continue

        if estado == "terminado":
            if tipo in ("Escape", "Salida"):
                revisar.append(f"{tipo} {h} después de la Salida, no se tomó en cuenta")
                continue
            # Volvió a entrar: esa Salida en realidad fue un Escape.
            revisar.append(f"Salida {hora(salida)} seguida de más marcas, se tomó como Escape")
            estado, desde, salida = "pausa", salida, None
            pausas += 1

        if tipo == "Ingreso":
            if estado == "sin_ingreso":
                ingreso, desde, estado = t, t, "dentro"
                continue
            if estado == "dentro":
                revisar.append(f"Ingreso {h} de más, no se tomó en cuenta")
                continue
            revisar.append(f"Ingreso {h} después de un Escape, se tomó como Intrusión")
            tipo = "Intrusión"

        if tipo == "Intrusión":
            if estado == "pausa":
                fuera += t - desde
            elif estado == "dentro":
                faltas.append(
                    f"Falta Escape antes de la Intrusión {h} "
                    f"(tramo {hora(desde)}–{h} sin contar)"
                )
            else:
                faltas.append(f"Falta Ingreso (primera marca: Intrusión {h})")
            desde, estado = t, "dentro"

        elif tipo == "Escape":
            if estado == "dentro":
                trabajado += t - desde
            elif estado == "pausa":
                faltas.append(f"Falta Intrusión entre los Escapes de {hora(desde)} y {h}")
            else:
                faltas.append(f"Falta Ingreso (primera marca: Escape {h})")
            desde, estado = t, "pausa"
            pausas += 1

        else:  # Salida
            if estado == "dentro":
                trabajado += t - desde
            elif estado == "pausa":
                faltas.append(f"Falta Intrusión después del Escape {hora(desde)}")
            else:
                faltas.append(f"Falta Ingreso (primera marca: Salida {h})")
            salida, estado = t, "terminado"

    if estado == "dentro":
        faltas.append(f"Falta Salida (tramo desde {hora(desde)} sin contar)")
    elif estado == "pausa":
        faltas.append(f"Falta Intrusión y Salida después del Escape {hora(desde)}")
    elif estado == "sin_ingreso":
        faltas.append("Sin marcas de asistencia")

    observaciones = faltas + revisar
    if sin_estado:
        observaciones.append(f"Marcas sin estado (no se cuentan): {', '.join(sin_estado)}")

    return {
        "ingreso": ingreso,
        "salida": salida,
        "pausas": pausas,
        "fuera": fuera,
        "trabajado": trabajado,
        "estado": "Incompleto" if faltas else "Revisar" if revisar else "Completo",
        "observaciones": observaciones,
    }


def resumir(registros):
    """Una fila por usuario con su jornada calculada (solo marcas exitosas)."""
    por_usuario = defaultdict(list)
    for r in registros:
        if r["resultado"] == "Exitoso":
            por_usuario[r["id"]].append(r)

    filas = []
    for uid, marcas in por_usuario.items():
        fila = calcular_jornada(marcas)  # los registros ya vienen ordenados
        fila["id"] = uid
        fila["nombre"] = next((m["nombre"] for m in marcas if m["nombre"]), "")
        filas.append(fila)

    filas.sort(key=lambda f: f["nombre"].lower())
    return filas


def resumir_semana(dias):
    """Una fila por usuario con su jornada de cada día y los totales de la semana.

    `dias` es una lista de (fecha, registros de ese día).
    """
    usuarios = {}
    for dia, registros in dias:
        for jornada in resumir(registros):
            u = usuarios.setdefault(jornada["id"], {"id": jornada["id"], "nombre": "", "dias": {}})
            u["nombre"] = u["nombre"] or jornada["nombre"]
            u["dias"][dia] = jornada

    filas = list(usuarios.values())
    for u in filas:
        jornadas = list(u["dias"].values())
        estados = {j["estado"] for j in jornadas}
        u["trabajado"] = sum((j["trabajado"] for j in jornadas), timedelta())
        u["fuera"] = sum((j["fuera"] for j in jornadas), timedelta())
        u["estado"] = next(e for e in ("Incompleto", "Revisar", "Completo") if e in estados)
        u["observaciones"] = [
            f"{dia_corto(dia)}: {obs}"
            for dia, j in u["dias"].items()
            for obs in j["observaciones"]
        ]

    filas.sort(key=lambda f: f["nombre"].lower())
    return filas


def lunes(dia):
    return dia - timedelta(days=dia.weekday())


def dia_corto(dia):
    return f"{DIAS_CORTOS[dia.weekday()]} {dia.day}"


def texto_periodo(inicio, semanal):
    if not semanal:
        return inicio.strftime("%d-%m-%Y")
    return f"{inicio:%d-%m-%Y} al {inicio + timedelta(days=6):%d-%m-%Y}"


def columnas_semana(inicio):
    """Columnas del resumen semanal; cada día lleva su fecha en el título."""
    dias = [(f"d{i}", dia_corto(inicio + timedelta(days=i)), 58, "center") for i in range(7)]
    return [
        ("id", "ID", 45, "center"),
        ("nombre", "Nombre", 160, "w"),
        *dias,
        ("dias", "Días", 45, "center"),
        ("fuera", "Tiempo fuera", 90, "center"),
        ("trabajado", "Total horas", 85, "center"),
        ("estado", "Estado", 85, "center"),
        ("obs", "Observaciones", 240, "w"),
    ]


def fila_registro(r):
    return (
        r["fecha"].strftime("%d-%m-%Y"),
        r["fecha"].strftime("%H:%M:%S"),
        r["id"],
        r["nombre"],
        r["marca"],
        r["metodo"],
        r["resultado"],
        "Agregada a mano" if r.get("manual") else
        f"Antes: {r['original']}" if "original" in r else "",
    )


def etiquetas_registro(r):
    if r.get("manual"):
        return ("agregado",)
    return ("corregido",) if "original" in r else ()


def fila_resumen(f):
    return (
        f["id"],
        f["nombre"],
        hora(f["ingreso"]) if f["ingreso"] else "—",
        hora(f["salida"]) if f["salida"] else "—",
        f["pausas"],
        duracion(f["fuera"]),
        duracion(f["trabajado"]),
        f["estado"],
        "; ".join(f["observaciones"]),
    )


def fila_semana(f, inicio):
    horas_por_dia = []
    for i in range(7):
        jornada = f["dias"].get(inicio + timedelta(days=i))
        if jornada is None:
            horas_por_dia.append("—")
        else:
            aviso = "" if jornada["estado"] == "Completo" else " ⚠"
            horas_por_dia.append(duracion(jornada["trabajado"]) + aviso)
    return (
        f["id"],
        f["nombre"],
        *horas_por_dia,
        len(f["dias"]),
        duracion(f["fuera"]),
        duracion(f["trabajado"]),
        f["estado"],
        "; ".join(f["observaciones"]),
    )


# --- Texto para Slack -------------------------------------------------------

def plural(n, palabra):
    return f"{n} {palabra}{'' if n == 1 else 's'}"


def texto_slack(resumen, inicio, semanal):
    """Resumen con formato de Slack (*negrita* y :emojis:), listo para pegar."""
    if semanal:
        titulo = f"Asistencia de la semana del {texto_periodo(inicio, True)}"
    else:
        titulo = f"Asistencia del {DIAS[inicio.weekday()]} {inicio:%d-%m-%Y}"

    cuenta = [plural(len(resumen), "persona")]
    incompletos = sum(f["estado"] == "Incompleto" for f in resumen)
    revisar = sum(f["estado"] == "Revisar" for f in resumen)
    if incompletos:
        cuenta.append(f"{incompletos} con marcas faltantes")
    if revisar:
        cuenta.append(f"{revisar} para revisar")

    lineas = [f"*{titulo}*", " · ".join(cuenta), ""]
    for f in resumen:
        detalle = f"{duracion(f['trabajado'])} h"
        if semanal:
            detalle += f" en {plural(len(f['dias']), 'día')}"
        elif f["ingreso"] or f["salida"]:
            entrada = hora(f["ingreso"]) if f["ingreso"] else "—"
            salida = hora(f["salida"]) if f["salida"] else "—"
            detalle += f" ({entrada} a {salida})"
        nombre = f["nombre"] or f"ID {f['id']}"
        lineas.append(f"{ICONOS_SLACK[f['estado']]} *{nombre}*: {detalle}")
        lineas.extend(f"      • {obs}" for obs in f["observaciones"])

    lineas += [
        "",
        f"{ICONOS_SLACK['Completo']} completo   {ICONOS_SLACK['Revisar']} revisar   "
        f"{ICONOS_SLACK['Incompleto']} faltan marcas (esas horas no se cuentan)",
    ]
    return "\n".join(lineas)


# --- Excel ------------------------------------------------------------------

def guardar_excel(archivo, registros, columnas_resumen, filas_resumen):
    wb = Workbook()
    hojas = [
        (wb.active, "Registros", COLUMNAS_REGISTROS, [fila_registro(r) for r in registros]),
        (wb.create_sheet(), "Resumen", columnas_resumen, filas_resumen),
    ]
    for hoja, titulo, columnas, filas in hojas:
        hoja.title = titulo
        hoja.append([c[1] for c in columnas])
        for fila in filas:
            hoja.append(fila)

        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for celda in hoja[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill(fill_type="solid", fgColor=AZUL[1:])
            celda.alignment = Alignment(horizontal="center")
        for columna in hoja.columns:
            ancho = max(len(str(c.value or "")) for c in columna)
            hoja.column_dimensions[columna[0].column_letter].width = min(ancho + 3, 60)

    wb.save(archivo)


# --- Selector de fecha ------------------------------------------------------

class Calendario(tk.Toplevel):
    """Ventana pequeña con el mes para elegir un día (o la semana de ese día)."""

    def __init__(self, selector, fecha, al_elegir, semanal=False):
        super().__init__(selector)
        self.title("Elegir semana" if semanal else "Elegir fecha")
        self.resizable(False, False)
        self.transient(selector.winfo_toplevel())
        self.al_elegir = al_elegir
        # Días resaltados: el elegido, o los 7 de la semana elegida.
        self.marcados = {fecha + timedelta(days=i) for i in range(7 if semanal else 1)}
        self.anio, self.mes = fecha.year, fecha.month

        cabecera = ttk.Frame(self, padding=(8, 8, 8, 4))
        cabecera.pack(fill="x")
        ttk.Button(cabecera, text="◀", width=3, command=lambda: self.cambiar_mes(-1)).pack(side="left")
        ttk.Button(cabecera, text="▶", width=3, command=lambda: self.cambiar_mes(1)).pack(side="right")
        self.titulo = ttk.Label(cabecera, anchor="center", font=("Segoe UI", 10, "bold"))
        self.titulo.pack(fill="x", expand=True)

        self.cuadricula = ttk.Frame(self, padding=(8, 0, 8, 8))
        self.cuadricula.pack()
        self.dibujar()

        x = selector.winfo_rootx()
        y = selector.winfo_rooty() + selector.winfo_height() + 2
        self.geometry(f"+{x}+{y}")
        self.bind("<Escape>", lambda _e: self.destroy())
        self.wait_visibility()
        self.grab_set()
        self.focus_set()

    def cambiar_mes(self, delta):
        indice = self.mes - 1 + delta
        self.anio += indice // 12
        self.mes = indice % 12 + 1
        self.dibujar()

    def dibujar(self):
        for widget in self.cuadricula.winfo_children():
            widget.destroy()
        self.titulo.config(text=f"{MESES[self.mes - 1]} {self.anio}")

        for col, nombre in enumerate(["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]):
            ttk.Label(self.cuadricula, text=nombre, width=4, anchor="center").grid(row=0, column=col)

        hoy = date.today()
        for fila, semana in enumerate(calendar.monthcalendar(self.anio, self.mes), start=1):
            for col, numero in enumerate(semana):
                if not numero:
                    continue
                dia = date(self.anio, self.mes, numero)
                boton = tk.Button(
                    self.cuadricula, text=numero, width=3, relief="flat",
                    command=lambda d=dia: self.elegir(d),
                )
                if dia in self.marcados:
                    boton.config(bg=AZUL, fg="white", activebackground=AZUL, activeforeground="white")
                elif dia == hoy:
                    boton.config(fg=AZUL, font=("Segoe UI", 9, "bold"))
                if dia > hoy:
                    boton.config(state="disabled")
                boton.grid(row=fila, column=col, padx=1, pady=1)

    def elegir(self, dia):
        self.al_elegir(dia)
        self.destroy()


class SelectorFecha(ttk.Frame):
    """Campo de fecha con anterior/siguiente y calendario.

    En modo semanal el periodo es la semana (lunes a domingo) del día elegido
    y las flechas avanzan de semana en semana.
    """

    def __init__(self, padre):
        super().__init__(padre)
        self.texto = tk.StringVar()
        self.semanal = False

        ttk.Button(self, text="◀", width=3, command=lambda: self.mover(-1)).pack(side="left")
        self.campo = ttk.Entry(self, textvariable=self.texto, width=12, state="readonly", justify="center")
        self.campo.pack(side="left", padx=2)
        self.campo.bind("<Button-1>", lambda _e: self.abrir_calendario())
        self.btn_siguiente = ttk.Button(self, text="▶", width=3, command=lambda: self.mover(1))
        self.btn_siguiente.pack(side="left")
        ttk.Button(self, text="▼", width=3, command=self.abrir_calendario).pack(side="left", padx=(4, 0))

        self.poner(date.today())

    def poner(self, dia):
        hoy = date.today()
        self.dia = min(dia, hoy)  # se recuerda para volver al mismo día desde la semana
        self.fecha = lunes(self.dia) if self.semanal else self.dia  # inicio del periodo
        self.texto.set(texto_periodo(self.fecha, self.semanal))
        ultimo = lunes(hoy) if self.semanal else hoy
        self.btn_siguiente.state(["disabled"] if self.fecha >= ultimo else ["!disabled"])

    def mover(self, sentido):
        self.poner(self.dia + timedelta(days=sentido * (7 if self.semanal else 1)))

    def cambiar_modo(self, semanal):
        self.semanal = semanal
        self.campo.config(width=23 if semanal else 12)
        self.poner(self.dia)

    def abrir_calendario(self):
        Calendario(self, self.fecha, self.poner, self.semanal)


def abrir_como_dialogo(ventana, padre):
    """Centra la ventana sobre la principal y bloquea la principal mientras está abierta."""
    ventana.resizable(False, False)
    ventana.transient(padre)
    ventana.update_idletasks()
    x = padre.winfo_rootx() + (padre.winfo_width() - ventana.winfo_reqwidth()) // 2
    y = padre.winfo_rooty() + (padre.winfo_height() - ventana.winfo_reqheight()) // 3
    ventana.geometry(f"+{x}+{y}")
    ventana.bind("<Escape>", lambda _e: ventana.destroy())
    ventana.wait_visibility()
    ventana.grab_set()
    ventana.focus_set()


class CorregirMarca(tk.Toplevel):
    """Ventana para cambiar el estado de asistencia de una marca.

    Si la marca fue agregada a mano, también se puede eliminar.
    """

    def __init__(self, padre, registro, al_guardar, al_eliminar=None):
        super().__init__(padre)
        self.title("Corregir marca")
        self.al_guardar = al_guardar
        self.al_eliminar = al_eliminar
        self.marca = tk.StringVar(value=registro["marca"])
        original = None if registro.get("manual") else registro.get("original", registro["marca"])

        cuerpo = ttk.Frame(self, padding=14)
        cuerpo.pack(fill="both")
        nombre = registro["nombre"] or f"ID {registro['id']}"
        ttk.Label(cuerpo, text=nombre, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        fecha = f"{registro['fecha']:%d-%m-%Y  %H:%M:%S}"
        if registro.get("manual"):
            fecha += "  (agregada a mano)"
        ttk.Label(cuerpo, text=fecha).pack(anchor="w", pady=(0, 10))
        ttk.Label(cuerpo, text="Estado de asistencia correcto:").pack(anchor="w")
        for tipo in MARCAS:
            texto = f"{tipo}  (lo que se marcó en el equipo)" if tipo == original else tipo
            ttk.Radiobutton(cuerpo, text=texto, value=tipo, variable=self.marca).pack(anchor="w", padx=(12, 0), pady=1)

        botones = ttk.Frame(cuerpo)
        botones.pack(fill="x", pady=(14, 0))
        if al_eliminar:
            ttk.Button(botones, text="Eliminar marca", command=self.eliminar).pack(side="left", padx=(0, 18))
        ttk.Button(botones, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(botones, text="Guardar", command=self.guardar).pack(side="right", padx=(0, 6))

        self.bind("<Return>", lambda _e: self.guardar())
        abrir_como_dialogo(self, padre)

    def guardar(self):
        self.destroy()
        self.al_guardar(self.marca.get())

    def eliminar(self):
        if messagebox.askyesno("Eliminar marca", "¿Eliminar esta marca agregada a mano?", parent=self):
            self.destroy()
            self.al_eliminar()


class AgregarMarca(tk.Toplevel):
    """Ventana para agregar marcas a mano (ej.: alguien que trabajó virtual).

    Se queda abierta para poder agregar varias seguidas (Ingreso, Salida, ...).
    `al_agregar(id, nombre, fecha, marca)` devuelve un mensaje de error o None.
    """

    def __init__(self, padre, usuarios, dias, dia_inicial, al_agregar):
        super().__init__(padre)
        self.title("Agregar marca")
        self.al_agregar = al_agregar
        # Personas conocidas (nombre en minúsculas -> ID y nombre). Al elegir una
        # se usa su ID; el campo ID solo aparece para alguien que no está en la lista.
        self.conocidos = {nombre.strip().lower(): (uid, nombre) for uid, nombre in usuarios.items() if nombre}
        self.dias = {f"{DIAS[d.weekday()]} {d:%d-%m-%Y}": d for d in dias}

        cuerpo = ttk.Frame(self, padding=14)
        cuerpo.pack(fill="both")
        cuerpo.columnconfigure(1, weight=1)

        ttk.Label(cuerpo, text="Persona:").grid(row=0, column=0, sticky="w", pady=3)
        nombres = sorted((n for _, n in self.conocidos.values()), key=str.lower)
        self.nombre = ttk.Combobox(cuerpo, width=30, values=nombres)
        self.nombre.grid(row=0, column=1, sticky="we", pady=3)
        for evento in ("<<ComboboxSelected>>", "<KeyRelease>"):
            self.nombre.bind(evento, lambda _e: self.mostrar_campo_id())

        self.etiqueta_id = ttk.Label(cuerpo, text="ID:")
        self.etiqueta_id.grid(row=1, column=0, sticky="w", pady=3)
        self.fila_id = ttk.Frame(cuerpo)
        self.fila_id.grid(row=1, column=1, sticky="w", pady=3)
        self.uid = ttk.Entry(self.fila_id, width=10)
        self.uid.pack(side="left")
        ttk.Label(self.fila_id, text="persona nueva: su ID en el equipo").pack(side="left", padx=6)
        self.mostrar_campo_id()

        ttk.Label(cuerpo, text="Día:").grid(row=2, column=0, sticky="w", pady=3)
        self.dia = ttk.Combobox(cuerpo, state="readonly", width=24, values=list(self.dias))
        self.dia.set(next((t for t, d in self.dias.items() if d == dia_inicial), list(self.dias)[0]))
        self.dia.grid(row=2, column=1, sticky="w", pady=3)

        ttk.Label(cuerpo, text="Hora:").grid(row=3, column=0, sticky="w", pady=3)
        fila_hora = ttk.Frame(cuerpo)
        fila_hora.grid(row=3, column=1, sticky="w", pady=3)
        self.hora = ttk.Entry(fila_hora, width=7, justify="center")
        self.hora.pack(side="left")
        ttk.Label(fila_hora, text="ej.: 08:00 o 17:30").pack(side="left", padx=6)

        ttk.Label(cuerpo, text="Estado:").grid(row=4, column=0, sticky="nw", pady=(6, 0))
        self.marca = tk.StringVar(value=MARCAS[0])
        opciones = ttk.Frame(cuerpo)
        opciones.grid(row=4, column=1, sticky="w", pady=(6, 0))
        for tipo in MARCAS:
            ttk.Radiobutton(opciones, text=tipo, value=tipo, variable=self.marca).pack(anchor="w", pady=1)

        self.aviso = tk.StringVar(value="Puedes agregar varias marcas seguidas.")
        ttk.Label(cuerpo, textvariable=self.aviso, wraplength=320, justify="left").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))

        botones = ttk.Frame(cuerpo)
        botones.grid(row=6, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botones, text="Agregar", command=self.agregar).pack(side="left", padx=(0, 6))
        ttk.Button(botones, text="Cerrar", command=self.destroy).pack(side="left")

        self.bind("<Return>", lambda _e: self.agregar())
        abrir_como_dialogo(self, padre)
        self.nombre.focus_set()

    def persona_conocida(self):
        return self.conocidos.get(self.nombre.get().strip().lower())

    def mostrar_campo_id(self):
        """El ID solo se pide si el nombre escrito no está en la lista."""
        nuevo = self.nombre.get().strip() and not self.persona_conocida()
        for widget in (self.etiqueta_id, self.fila_id):
            if nuevo:
                widget.grid()
            else:
                widget.grid_remove()

    def agregar(self):
        conocida = self.persona_conocida()
        if conocida:
            uid, nombre = conocida
        else:
            uid, nombre = self.uid.get().strip(), self.nombre.get().strip()
        if not nombre:
            self.aviso.set("Elige a la persona de la lista.")
            self.nombre.focus_set()
            return
        if not uid:
            self.aviso.set(f"«{nombre}» no está en la lista. Elígela de la lista o, "
                           "si es una persona nueva, escribe su ID del equipo.")
            self.uid.focus_set()
            return
        try:
            hora_marca = datetime.strptime(self.hora.get().strip(), "%H:%M").time()
        except ValueError:
            self.aviso.set("Escribe la hora en formato 24 horas, por ejemplo 08:00 o 17:30.")
            self.hora.focus_set()
            return

        fecha = datetime.combine(self.dias[self.dia.get()], hora_marca)
        error = self.al_agregar(uid, nombre, fecha, self.marca.get())
        if error:
            self.aviso.set(error)
            return
        self.aviso.set(f"✓ Se agregó {self.marca.get()} {hora(fecha)} de {nombre}. "
                       "Puedes agregar otra o cerrar.")
        self.hora.delete(0, "end")
        self.hora.focus_set()


# --- Ventana principal ------------------------------------------------------

def poner_columnas(tabla, columnas):
    tabla.configure(columns=[c[0] for c in columnas])
    for clave, titulo, ancho, ancla in columnas:
        tabla.heading(clave, text=titulo, anchor=ancla)
        tabla.column(clave, width=ancho, minwidth=40, anchor=ancla,
                     stretch=clave in ("nombre", "obs"))


def crear_tabla(padre, columnas):
    marco = ttk.Frame(padre)
    marco.pack(fill="both", expand=True)
    tabla = ttk.Treeview(marco, show="headings", selectmode="browse")
    poner_columnas(tabla, columnas)
    barra = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y")
    tabla.pack(side="left", fill="both", expand=True)
    return tabla


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Asistencia - Control de acceso")
        self.geometry("1150x720")
        self.minsize(960, 550)

        # Datos cargados (cambian solo al presionar Generar).
        self.inicio = None     # primer día del periodo
        self.semanal = False
        self.dias = []         # (día, marcas tal como vienen del equipo)
        self.registros = []    # marcas con estado de asistencia válido (ya corregidas)
        self.resumen = []      # una fila por usuario (del día o de la semana)
        self.columnas_resumen = COLUMNAS_RESUMEN
        self.filas_resumen = []
        self.usuarios = {}     # texto del filtro -> ID de usuario

        estilo = ttk.Style(self)
        estilo.configure("Treeview", rowheight=24)
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        # Barra superior
        barra = ttk.Frame(self, padding=10)
        barra.pack(fill="x")
        self.modo = tk.StringVar(value="dia")
        for texto, valor in (("Día", "dia"), ("Semana", "semana")):
            ttk.Radiobutton(
                barra, text=texto, value=valor, variable=self.modo,
                command=lambda: self.selector.cambiar_modo(self.modo.get() == "semana"),
            ).pack(side="left", padx=(0, 6))
        self.selector = SelectorFecha(barra)
        self.selector.pack(side="left", padx=(6, 12))
        self.btn_generar = ttk.Button(barra, text="Generar", command=self.generar)
        self.btn_generar.pack(side="left")
        self.btn_exportar = ttk.Button(barra, text="Exportar a Excel", command=self.exportar)
        self.btn_exportar.pack(side="left", padx=(6, 0))
        self.btn_copiar = ttk.Button(barra, text="Copiar resumen", command=self.copiar_resumen)
        self.btn_copiar.pack(side="left", padx=6)
        for boton in (self.btn_exportar, self.btn_copiar):
            boton.state(["disabled"])

        ttk.Label(barra, text="Usuario:").pack(side="left", padx=(18, 6))
        self.filtro = ttk.Combobox(barra, state="readonly", width=24, values=["Todos"])
        self.filtro.set("Todos")
        self.filtro.pack(side="left")
        self.filtro.bind("<<ComboboxSelected>>", lambda _e: self.llenar_registros())

        # Barra de estado
        self.estado = tk.StringVar(value="Elige una fecha y presiona Generar.")
        ttk.Label(self, textvariable=self.estado, padding=(10, 0, 10, 8)).pack(side="bottom", fill="x")

        # Tablas
        paneles = ttk.PanedWindow(self, orient="vertical")
        paneles.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.marco_registros = ttk.LabelFrame(paneles, text="Registros", padding=6)
        pie = ttk.Frame(self.marco_registros)
        pie.pack(side="bottom", fill="x", pady=(6, 0))
        self.btn_corregir = ttk.Button(pie, text="Corregir marca…", command=self.abrir_correccion)
        self.btn_corregir.pack(side="left")
        self.btn_agregar = ttk.Button(pie, text="Agregar marca…", command=self.abrir_agregar)
        self.btn_agregar.pack(side="left", padx=(6, 0))
        for boton in (self.btn_corregir, self.btn_agregar):
            boton.state(["disabled"])
        ttk.Label(pie, text=AYUDA_CORREGIR).pack(side="left", padx=8)
        self.tabla_registros = crear_tabla(self.marco_registros, COLUMNAS_REGISTROS)
        self.tabla_registros.tag_configure("corregido", background="#DDEBF7")
        self.tabla_registros.tag_configure("agregado", background="#E2EFDA")
        self.tabla_registros.bind("<<TreeviewSelect>>", self.al_elegir_registro)
        self.tabla_registros.bind("<Double-1>", self.al_doble_clic)
        paneles.add(self.marco_registros, weight=3)

        self.marco_resumen = ttk.LabelFrame(paneles, text="Resumen por usuario", padding=6)
        self.detalle = tk.StringVar(value=AYUDA_RESUMEN)
        detalle = ttk.Label(self.marco_resumen, textvariable=self.detalle, justify="left")
        detalle.pack(side="bottom", fill="x", pady=(6, 0))
        detalle.bind("<Configure>", lambda e: detalle.config(wraplength=e.width - 10))
        self.tabla_resumen = crear_tabla(self.marco_resumen, COLUMNAS_RESUMEN)
        self.tabla_resumen.tag_configure("Incompleto", background="#F8D7DA")
        self.tabla_resumen.tag_configure("Revisar", background="#FFF3CD")
        self.tabla_resumen.bind("<<TreeviewSelect>>", self.al_elegir_resumen)
        paneles.add(self.marco_resumen, weight=2)

        try:
            self.cambios = cargar_cambios()
        except (OSError, TypeError, ValueError) as error:
            self.cambios = {"correcciones": {}, "agregadas": [], "usuarios": {}}
            messagebox.showwarning(
                "Correcciones",
                f"No se pudo leer {ARCHIVO_CORRECCIONES}:\n{error}\n\n"
                "Las correcciones y marcas agregadas antes no se van a aplicar. "
                "Si guardas un cambio nuevo, ese archivo se reemplaza.",
                parent=self,
            )

    # Consulta en segundo plano para que la ventana no se congele.
    def generar(self):
        inicio, semanal = self.selector.fecha, self.selector.semanal
        # Un día por consulta (el equipo limita la cantidad de registros);
        # los días de la semana que aún no llegan no se consultan.
        dias = [inicio + timedelta(days=i) for i in range(7 if semanal else 1)]
        dias = [d for d in dias if d <= date.today()]
        self.btn_generar.state(["disabled"])
        self.config(cursor="watch")
        self.estado.set(f"Consultando el equipo ({IP})…")
        resultado = {}

        def consultar():
            try:
                resultado["dias"] = [(d, consultar_registros(d)) for d in dias]
            except Exception as error:  # se le muestra al usuario
                resultado["error"] = error

        hilo = threading.Thread(target=consultar, daemon=True)
        hilo.start()
        self.esperar(hilo, resultado, inicio, semanal)

    def esperar(self, hilo, resultado, inicio, semanal):
        if hilo.is_alive():
            self.after(100, self.esperar, hilo, resultado, inicio, semanal)
            return
        self.btn_generar.state(["!disabled"])
        self.config(cursor="")
        if "error" in resultado:
            self.estado.set("No se pudo obtener la información.")
            messagebox.showerror("Error", mensaje_error(resultado["error"]), parent=self)
            return
        self.mostrar(inicio, semanal, resultado["dias"])

    def mostrar(self, inicio, semanal, dias):
        self.inicio, self.semanal, self.dias = inicio, semanal, dias
        self.recordar_usuarios()
        self.calcular()
        self.usuarios = {f["nombre"] or f"ID {f['id']}": f["id"] for f in self.resumen}

        self.filtro["values"] = ["Todos", *self.usuarios]
        if self.filtro.get() not in self.usuarios:
            self.filtro.set("Todos")

        periodo = texto_periodo(inicio, semanal)
        self.marco_registros.config(text=f"Registros del {periodo}")
        self.marco_resumen.config(text=f"Resumen por usuario del {periodo}")
        self.llenar_registros()
        self.llenar_resumen()
        for boton in (self.btn_exportar, self.btn_copiar):
            boton.state(["!disabled"] if self.resumen else ["disabled"])
        self.btn_agregar.state(["!disabled"])
        self.estado.set(self.texto_estado())

    def recordar_usuarios(self):
        """Guarda los nombres que vienen del equipo, para elegirlos al agregar marcas
        (así aparece alguien aunque ese día no haya marcado)."""
        conocidos = self.cambios["usuarios"]
        nuevos = {r["id"]: r["nombre"] for _, registros in self.dias
                  for r in registros if r["id"] and r["nombre"]}
        if all(conocidos.get(uid) == nombre for uid, nombre in nuevos.items()):
            return
        conocidos.update(nuevos)
        try:
            guardar_cambios(self.cambios)
        except OSError:
            pass  # solo sirve para sugerir nombres; se intenta de nuevo la próxima vez

    def calcular(self):
        """Aplica los cambios hechos a mano a las marcas del equipo y calcula el resumen."""
        agregadas = defaultdict(list)
        for r in marcas_agregadas(self.cambios["agregadas"]):
            agregadas[r["fecha"].date()].append(r)
        dias = []
        for d, registros in self.dias:
            registros = aplicar_correcciones(registros, self.cambios["correcciones"]) + agregadas[d]
            registros.sort(key=lambda r: r["fecha"])
            dias.append((d, registros))
        todos = [r for _, registros in dias for r in registros]
        self.registros = [r for r in todos if r["marca"] in ESTADOS_VALIDOS]
        if self.semanal:
            self.resumen = resumir_semana(dias)
            self.columnas_resumen = columnas_semana(self.inicio)
            self.filas_resumen = [fila_semana(f, self.inicio) for f in self.resumen]
        else:
            self.resumen = resumir(todos)
            self.columnas_resumen = COLUMNAS_RESUMEN
            self.filas_resumen = [fila_resumen(f) for f in self.resumen]

    def texto_estado(self):
        periodo = texto_periodo(self.inicio, self.semanal)
        if not self.registros and not self.resumen:
            return f"No hay marcas para el {periodo}."
        texto = f"{periodo}: {len(self.registros)} marcas de {len(self.resumen)} usuarios."
        incompletos = sum(f["estado"] == "Incompleto" for f in self.resumen)
        if incompletos:
            texto += f" {incompletos} con marcas faltantes."
        llenos = [d for d, registros in self.dias if len(registros) >= MAX_REGISTROS]
        if llenos:
            cuando = f" el {', '.join(dia_corto(d) for d in llenos)}" if self.semanal else ""
            texto += (
                f" Atención: el equipo devolvió el máximo de {MAX_REGISTROS} "
                f"registros{cuando}, puede que falten marcas."
            )
        return texto

    def llenar_registros(self):
        tabla = self.tabla_registros
        tabla.delete(*tabla.get_children())
        self.btn_corregir.state(["disabled"])
        uid = self.usuarios.get(self.filtro.get())  # None = Todos
        # El iid de cada fila es su posición en self.registros.
        for i, r in enumerate(self.registros):
            if uid is None or r["id"] == uid:
                tabla.insert("", "end", iid=str(i), values=fila_registro(r), tags=etiquetas_registro(r))

    def al_elegir_registro(self, _evento):
        self.btn_corregir.state(["!disabled"] if self.tabla_registros.selection() else ["disabled"])

    def al_doble_clic(self, evento):
        tabla = self.tabla_registros
        if tabla.identify_region(evento.x, evento.y) == "cell":
            self.abrir_correccion(tabla.identify_row(evento.y))

    def abrir_correccion(self, iid=None):
        iid = iid or next(iter(self.tabla_registros.selection()), None)
        if not iid:
            return
        registro = self.registros[int(iid)]
        eliminar = (lambda: self.eliminar(registro)) if registro.get("manual") else None
        CorregirMarca(self, registro, lambda marca: self.corregir(registro, marca), eliminar)

    def abrir_agregar(self):
        dias = [d for d, _ in self.dias]
        AgregarMarca(self, self.cambios["usuarios"], dias, self.selector.dia, self.agregar)

    def guardar(self, anteriores):
        """Guarda los cambios en el archivo; si no se puede, vuelve a `anteriores`."""
        try:
            guardar_cambios(self.cambios)
            return True
        except OSError as error:
            self.cambios = anteriores
            messagebox.showerror(
                "No se pudo guardar",
                f"No se pudo guardar el cambio en {ARCHIVO_CORRECCIONES}:\n{error}",
                parent=self.grab_current() or self,
            )
            return False

    def buscar_agregada(self, registro):
        clave = clave_registro(registro)
        return next((a for a in self.cambios["agregadas"] if clave_agregada(a) == clave), None)

    def corregir(self, registro, marca):
        actual = registro["marca"]
        original = registro.get("original", actual)
        if marca == actual:
            return
        anteriores = copy.deepcopy(self.cambios)
        if registro.get("manual"):
            agregada = self.buscar_agregada(registro)
            if agregada is None:
                return
            agregada["marca"] = marca
        elif marca == original:
            self.cambios["correcciones"].pop(clave_registro(registro), None)
        else:
            self.cambios["correcciones"][clave_registro(registro)] = marca
        if not self.guardar(anteriores):
            return

        self.calcular()
        self.refrescar_tablas()
        nombre = registro["nombre"] or f"ID {registro['id']}"
        cambio = "se dejó como en el equipo" if marca == original else f"{actual} → {marca}"
        self.estado.set(f"Marca de {nombre} a las {hora(registro['fecha'])}: {cambio}. {self.texto_estado()}")

    def agregar(self, uid, nombre, fecha, marca):
        """Agrega una marca hecha a mano. Devuelve un mensaje si no se pudo."""
        repetida = any(
            r["id"] == uid and r["fecha"].replace(second=0) == fecha
            for _, registros in self.dias for r in registros
        ) or any(r["id"] == uid and r["fecha"] == fecha for r in marcas_agregadas(self.cambios["agregadas"]))
        if repetida:
            return f"{nombre} ya tiene una marca a las {hora(fecha)} ese día."

        anteriores = copy.deepcopy(self.cambios)
        self.cambios["agregadas"].append(
            {"id": uid, "nombre": nombre, "fecha": f"{fecha:{FORMATO_FECHA}}", "marca": marca}
        )
        self.cambios["usuarios"].setdefault(uid, nombre)
        if not self.guardar(anteriores):
            return "No se pudo guardar la marca."

        self.mostrar(self.inicio, self.semanal, self.dias)
        self.elegir_usuario(uid)
        self.estado.set(f"Se agregó {marca} {hora(fecha)} de {nombre}. {self.texto_estado()}")
        return None

    def eliminar(self, registro):
        agregada = self.buscar_agregada(registro)
        if agregada is None:
            return
        anteriores = copy.deepcopy(self.cambios)
        self.cambios["agregadas"].remove(agregada)
        if not self.guardar(anteriores):
            return
        self.mostrar(self.inicio, self.semanal, self.dias)
        self.elegir_usuario(registro["id"])
        nombre = registro["nombre"] or f"ID {registro['id']}"
        self.estado.set(
            f"Se eliminó la marca {registro['marca']} {hora(registro['fecha'])} de {nombre}. "
            f"{self.texto_estado()}"
        )

    def elegir_usuario(self, uid):
        """Selecciona a la persona en el resumen (y así se filtran sus marcas)."""
        indice = next((i for i, f in enumerate(self.resumen) if f["id"] == uid), None)
        if indice is not None:
            self.tabla_resumen.selection_set(str(indice))
            self.tabla_resumen.see(str(indice))

    def refrescar_tablas(self):
        """Pone los valores recalculados sin perder el filtro ni la selección.

        Una corrección no cambia qué marcas ni qué usuarios hay, solo sus
        valores, así que las filas siguen en el mismo lugar.
        """
        for iid in self.tabla_registros.get_children():
            r = self.registros[int(iid)]
            self.tabla_registros.item(iid, values=fila_registro(r), tags=etiquetas_registro(r))
        for i, (f, valores) in enumerate(zip(self.resumen, self.filas_resumen)):
            self.tabla_resumen.item(str(i), values=valores, tags=(f["estado"],))
        seleccion = self.tabla_resumen.selection()
        if seleccion:
            self.detalle.set(self.texto_detalle(self.resumen[int(seleccion[0])]))

    def llenar_resumen(self):
        tabla = self.tabla_resumen
        tabla.delete(*tabla.get_children())
        poner_columnas(tabla, self.columnas_resumen)
        for i, (f, valores) in enumerate(zip(self.resumen, self.filas_resumen)):
            tabla.insert("", "end", iid=str(i), values=valores, tags=(f["estado"],))
        self.detalle.set(AYUDA_SEMANA if self.semanal else AYUDA_RESUMEN)

    def al_elegir_resumen(self, _evento):
        seleccion = self.tabla_resumen.selection()
        if not seleccion:
            return
        fila = self.resumen[int(seleccion[0])]
        self.detalle.set(self.texto_detalle(fila))
        etiqueta = next((k for k, v in self.usuarios.items() if v == fila["id"]), "Todos")
        self.filtro.set(etiqueta)
        self.llenar_registros()

    def texto_detalle(self, fila):
        sin_obs = f"{'Semana' if self.semanal else 'Jornada'} completa, sin observaciones."
        return f"{fila['nombre']}: {'; '.join(fila['observaciones']) or sin_obs}"

    def exportar(self):
        archivo = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar Excel",
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")],
            initialfile=f"asistencia_{texto_periodo(self.inicio, self.semanal).replace(' ', '_')}.xlsx",
        )
        if not archivo:
            return
        try:
            guardar_excel(archivo, self.registros, self.columnas_resumen, self.filas_resumen)
        except PermissionError:
            messagebox.showerror(
                "No se pudo guardar",
                "El archivo está abierto en otro programa. Ciérralo e intenta de nuevo.",
                parent=self,
            )
            return
        self.estado.set(f"Excel guardado: {archivo}")

    def copiar_resumen(self):
        self.clipboard_clear()
        self.clipboard_append(texto_slack(self.resumen, self.inicio, self.semanal))
        self.estado.set("Resumen copiado. Pégalo en Slack con Ctrl+V.")


if __name__ == "__main__":
    try:  # texto nítido en pantallas con escala en Windows
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass
    App().mainloop()
