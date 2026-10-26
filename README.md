# Asistencia - Control de acceso

Programa para consultar las marcas de asistencia del equipo de control de acceso Dahua (reconocimiento facial) y calcular las horas laboradas de cada persona en un día.

## Uso

1. Abrir `dist/Asistencia.exe` (o `py interfaz.py` si se tiene Python).
2. Elegir **Día** o **Semana** (de lunes a domingo).
3. Elegir el día o la semana con las flechas ◀ ▶ o con el calendario ▼ (en modo semana, al elegir cualquier día se toma su semana).
4. Presionar **Generar**.

La ventana muestra dos tablas:

- **Registros:** todas las marcas del periodo (fecha, hora, usuario, estado de asistencia, método). Se puede filtrar por persona con el selector **Usuario**.
- **Resumen por usuario:**
  - Por día: hora de ingreso y salida, cantidad de escapes, tiempo fuera, horas laboradas y observaciones.
  - Por semana: horas laboradas de cada día (⚠ = ese día tiene observaciones), días con marcas, tiempo fuera, total de horas y observaciones (cada una indica el día).

  Al hacer clic en una persona se filtran sus marcas en la tabla de arriba y se muestran sus observaciones completas.

Con **Exportar a Excel** se guardan ambas tablas en un archivo `.xlsx` (hojas "Registros" y "Resumen").

Con **Copiar resumen** se copia al portapapeles un mensaje listo para pegar en Slack (Ctrl+V): título con la fecha o semana, conteo de personas con problemas y una línea por persona con sus horas, un círculo de color según su estado (verde, amarillo o rojo) y sus observaciones.

En la semana actual solo se consultan los días hasta hoy.

La computadora debe estar en la misma red que el equipo.

## Cómo se calculan las horas

La jornada esperada de cada persona es:

```
Ingreso → (Escape → Intrusión) × n → Salida
```

- **Ingreso:** entrada al inicio del día (solo una).
- **Escape:** sale temporalmente (almuerzo, diligencia, etc.).
- **Intrusión:** regresa después de un Escape.
- **Salida:** salida al final del día.

**Horas laboradas** es la suma de los tramos dentro de la oficina (Ingreso/Intrusión → Escape/Salida). **Tiempo fuera** es la suma de los tramos Escape → Intrusión.

Solo se suman los tramos que tienen marca de inicio y de fin. Si falta una marca, ese tramo no se cuenta y se indica en las observaciones.

| Color | Estado | Significado |
|---|---|---|
| Sin color | Completo | Todas las marcas están en orden. |
| Amarillo | Revisar | Hubo una marca equivocada que se pudo interpretar (ej.: un segundo Ingreso después de un Escape se toma como Intrusión; una Salida seguida de más marcas se toma como Escape). Las horas se calcularon, pero conviene revisarlas. |
| Rojo | Incompleto | Falta al menos una marca (ej.: Escape sin Intrusión, sin Salida). Las horas de ese tramo no se cuentan. |

Las marcas sin estado de asistencia (aparecen como "-" en el equipo) y los intentos fallidos no se cuentan, pero las marcas sin estado se mencionan en las observaciones.

Si se consulta el día actual, quien todavía no ha marcado Salida aparecerá como Incompleto.

## Configuración

Los datos de conexión al equipo están al inicio de `interfaz.py`:

```python
IP = "192.168.3.1"
USUARIO = "admin"
CLAVE = "..."
```

Si cambian, hay que editar el archivo y volver a generar el `.exe` (ver abajo).

## Ejecutar desde el código

Requiere Python 3 (probado con 3.14) y estas librerías:

```
py -m pip install requests openpyxl
```

Luego:

```
py interfaz.py
```

## Generar el .exe

Requiere PyInstaller (`py -m pip install pyinstaller`). Desde la carpeta del proyecto:

```
py -m PyInstaller --noconfirm --onefile --windowed --name Asistencia interfaz.py
```

El ejecutable queda en `dist/Asistencia.exe` y se puede copiar a otra PC con Windows sin instalar Python. Las carpetas `build/` y el archivo `Asistencia.spec` que se crean se pueden borrar.

Al ser un ejecutable sin firma digital, Windows SmartScreen puede mostrar un aviso la primera vez: "Más información" → "Ejecutar de todas formas".

## Publicar una versión

El `.exe` no se guarda en el repositorio (`dist/` está en `.gitignore`). Se publica como archivo adjunto en la sección **Releases** de GitHub, desde donde se descarga.

Requiere [GitHub CLI](https://cli.github.com/) (`winget install GitHub.cli`) con sesión iniciada (`gh auth login`) en una cuenta con acceso al repositorio.

1. Generar el `.exe` (ver arriba).
2. Crear el tag y subirlo:

   ```
   git tag -a v1.0.0 -m "Asistencia v1.0.0"
   git push origin main --tags
   ```

3. Crear el release con el ejecutable adjunto:

   ```
   gh release create v1.0.0 dist/Asistencia.exe --title "Asistencia v1.0.0"
   ```

Para una versión nueva se cambia `v1.0.0` por el número que corresponda (ej.: `v1.1.0`).

## Archivos

| Archivo | Descripción |
|---|---|
| `interfaz.py` | Programa con la interfaz gráfica. |
| `dist/Asistencia.exe` | Ejecutable generado a partir de `interfaz.py`. |
| `main.py` | Script sin interfaz que genera `registros_dahua.xlsx` para una fecha fija escrita en el código. |
| `AttendanceRecord_*.xlsx` | Exportación hecha desde la página del equipo; se usó para identificar qué significa cada código. |

## Detalles técnicos

Los registros se obtienen con la API CGI del equipo:

```
GET http://<IP>/cgi-bin/recordFinder.cgi?action=find&name=AccessControlCardRec&StartTime=<inicio>&EndTime=<fin>&count=100
```

con autenticación Digest. El equipo devuelve códigos numéricos, que se traducen así (comparando con la exportación del equipo):

| Campo | Código → Texto |
|---|---|
| `AttendanceState` | 0 → "-", 1 → Ingreso, 2 → Escape, 3 → Intrusión, 4 → Salida |
| `Method` | 15 → Cara (otros métodos se muestran con su número) |
| `Status` | 1 → Exitoso, 0 → Fallido (el 0 todavía no aparece en los datos; se tomó de la documentación de Dahua) |

Cada consulta pide como máximo 100 registros. En modo semana se hace una consulta por día. Si un día llega a ese límite, la barra inferior muestra un aviso (indicando el día) porque podrían faltar marcas.
