#!/usr/bin/env python3
"""Base de datos en Excel: registro de colaboradores y parametros de calibracion.

Todo se guarda en .xlsx a proposito, para que se pueda abrir y corregir a mano
sin pasar por la aplicacion.
"""

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

def carpeta_base():
    """Dónde viven los Excel de datos.

    Empaquetado con PyInstaller, `__file__` apunta a una carpeta temporal que
    se borra al cerrar, así que hay que usar la del ejecutable. Si esa carpeta
    no se puede escribir (por ejemplo en Archivos de programa), se cae a una
    carpeta del usuario.
    """
    if getattr(sys, "frozen", False):
        # Empaquetado: los datos NO van junto al ejecutable. Dentro de un .app
        # quedarian escondidos en el paquete, y en Archivos de programa Windows
        # los desvia a una carpeta oculta. Una carpeta del usuario es visible,
        # se puede escribir siempre y es facil de decir por telefono.
        base = Path.home() / "Documents" / "Jornal"
    else:
        base = Path(__file__).resolve().parent
    try:
        base.mkdir(parents=True, exist_ok=True)
        prueba = base / ".permiso"
        prueba.touch()
        prueba.unlink()
        return base
    except OSError:
        alterna = Path.home() / "Jornal"
        alterna.mkdir(parents=True, exist_ok=True)
        return alterna


DIR_BASE = carpeta_base()
DIR_DATOS = DIR_BASE / "datos"
ARCHIVO_ROLES = DIR_DATOS / "roles.xlsx"
ARCHIVO_CONFIG = DIR_DATOS / "configuracion.xlsx"
ARCHIVO_PLANTILLA = DIR_DATOS / "plantilla.xlsx"
ARCHIVO_PLANTILLA_VERSION = DIR_DATOS / "plantilla.version"


def carpeta_resultados(cfg):
    """La carpeta donde caen los archivos generados. Si en la configuracion hay
    un nombre suelto, cuelga de la carpeta de la aplicacion; si hay una ruta
    completa, se respeta tal cual."""
    valor = str(cfg.get("carpeta_de_resultados") or "resultados").strip()
    ruta = Path(valor)
    if not ruta.is_absolute():
        ruta = DIR_BASE / ruta
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def hay_plantilla():
    return ARCHIVO_PLANTILLA.exists()


def version_plantilla():
    """Version de la plantilla cargada. Sube cada vez que se reemplaza, para
    saber de un vistazo si es la misma de siempre o una nueva de la jefatura."""
    if not ARCHIVO_PLANTILLA.exists():
        return ""
    try:
        v = ARCHIVO_PLANTILLA_VERSION.read_text(encoding="utf-8").strip()
        if v:
            return v
    except OSError:
        pass
    _escribir_version("1.0.0")      # la primera que hubo, sin registro previo
    return "1.0.0"


def _escribir_version(v):
    try:
        ARCHIVO_PLANTILLA_VERSION.parent.mkdir(parents=True, exist_ok=True)
        ARCHIVO_PLANTILLA_VERSION.write_text(v, encoding="utf-8")
    except OSError:
        pass


def _siguiente_version():
    """Cada plantilla nueva sube la minor: 1.0.0 -> 1.1.0 -> 1.2.0."""
    actual = version_plantilla() or "1.0.0"
    try:
        mayor, menor, parche = (int(x) for x in actual.split("."))
    except ValueError:
        return "1.0.0"
    return f"{mayor}.{menor + 1}.0"


def sembrar_plantilla_incluida():
    """Al empaquetar, la plantilla viaja dentro del ejecutable. La primera vez
    que se abre se copia a la carpeta de datos del usuario, para que la
    aplicacion funcione sin pedir nada."""
    if hay_plantilla() or not getattr(sys, "frozen", False):
        return False
    semilla = Path(getattr(sys, "_MEIPASS", "")) / "semilla" / "plantilla.xlsx"
    if not semilla.exists():
        return False
    import shutil
    ARCHIVO_PLANTILLA.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(semilla, ARCHIVO_PLANTILLA)
    _escribir_version("1.0.0")
    return True


def guardar_plantilla(origen):
    """Guarda una copia de la plantilla dentro de los datos de la aplicacion.
    Se valida antes: si el archivo no tiene la estructura esperada, no se copia."""
    import shutil
    import openpyxl as _ox
    import motor
    wb = _ox.load_workbook(origen, data_only=False)
    motor.leer_formato(wb[wb.sheetnames[0]])      # lanza si no sirve
    nueva = _siguiente_version() if ARCHIVO_PLANTILLA.exists() else "1.0.0"
    ARCHIVO_PLANTILLA.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(origen, ARCHIVO_PLANTILLA)
    _escribir_version(nueva)
    return ARCHIVO_PLANTILLA

# El listado de personal NO se guarda: sale de la hoja del formato que viene en
# el propio archivo del biometrico, que es la que cambia cada quincena. Aqui solo
# viven los pocos datos que el biometrico no puede saber.
COLUMNAS = ["cedula", "nombres", "rol", "licencia", "notas"]

# (clave, valor por defecto, tipo, descripcion para la vista de calibracion)
PARAMETROS = [
    ("jornal_dia_completo", 20, "numero",
     "Lo que se paga a un DIARIO por una jornada completa de lunes a viernes."),
    ("tarifa_hora_semana", 2.25, "numero",
     "Valor de la hora suelta entre semana, para las jornadas parciales."),
    ("tarifa_hora_finde", 3.0, "numero",
     "Valor de la hora en sábado y domingo."),
    ("ajuste_media_hora", -0.03, "numero",
     "Corrección que agosto aplica sólo cuando las horas terminan en .5"),
    ("sabado_estandar_horas", 5.0, "numero",
     "Duración del sábado normal. Con ese valor exacto se escribe el monto fijo."),
    ("sabado_estandar_valor", 15, "numero",
     "Monto fijo del sábado estándar."),
    ("carpeta_de_resultados", "resultados", "texto",
     "Dónde se guardan los archivos generados. Un nombre suelto se crea junto a la "
     "aplicación; también se puede poner una ruta completa."),
    ("horas_jornada_completa", 8.0, "numero",
     "Desde cuántas horas netas un día se paga completo en vez de por horas."),
    ("redondeo_horas", 0.5, "numero",
     "A qué fracción de hora se redondea. 0.5 = media hora."),
    ("minutos_marca_duplicada", 15, "entero",
     "Dos marcas separadas por menos de estos minutos se toman como una sola."),
    ("ventanas.fin_entrada", "11:00", "hora",
     "Antes de esta hora, una marca es ENTRADA."),
    ("ventanas.fin_alm_out", "11:45", "hora",
     "Hasta esta hora, una marca es salida al almuerzo."),
    ("ventanas.fin_mediodia", "12:15", "hora",
     "Hasta esta hora es MEDIODÍA: si no hay nada después, es salida de media jornada."),
    ("ventanas.fin_alm_in", "13:30", "hora",
     "Hasta esta hora, una marca es retorno del almuerzo. Después es SALIDA."),
    ("marca_afiliado_presente", "X", "texto",
     "Lo que se escribe a un AFILIADO presente."),
    ("marca_ausencia", "-", "texto",
     "Lo que se escribe cuando no hay registro de horas."),
    ("reconstruir_marca_faltante", True, "si_no",
     "Si falta una sola marca del día, estimarla con el horario habitual de la persona."),
    ("escribir_ausencias", True, "si_no",
     "Escribir el guión a quien no marcó un día laborable."),
    ("escribir_sabados", True, "si_no",
     "Calcular los sábados."),
    ("escribir_domingos", False, "si_no",
     "Calcular los domingos. Apagado: en agosto sólo una persona cobró domingo."),
    ("escribir_extras", False, "si_no",
     "Llenar la columna de extras. Apagado: las extras salen de la bitácora."),
    ("corregir_formula_subtotal", True, "si_no",
     "Reparar la fórmula de subtotal si quedó apuntando a un rango viejo."),
    ("ajustar_fechas_del_periodo", True, "si_no",
     "Poner en la plantilla las fechas y el título de la quincena que trae el biométrico."),
    ("respetar_lo_ya_escrito", True, "si_no",
     "No sobrescribir las celdas que el formato ya trae llenas: las puso una persona."),
    ("corregir_cedulas_en_el_formato", True, "si_no",
     "Escribir las cédulas completas, con su 0 inicial, en la columna del formato."),
    ("umbral_candidato_extras_salida", "17:00", "hora",
     "Salidas a esta hora o después se listan como posible hora extra."),
    ("umbral_candidato_extras_horas", 9.5, "numero",
     "Jornadas de estas horas o mas se listan como posible hora extra."),
]

# Parametros que la mayoria de las veces no se tocan: se ocultan tras un
# interruptor para que la pantalla de calibracion muestre solo lo esencial.
AVANZADOS = {
    "ajuste_media_hora", "minutos_marca_duplicada",
    "ventanas.fin_entrada", "ventanas.fin_alm_out", "ventanas.fin_mediodia",
    "ventanas.fin_alm_in", "marca_afiliado_presente", "marca_ausencia",
    "reconstruir_marca_faltante", "escribir_ausencias", "escribir_sabados",
    "escribir_domingos", "escribir_extras", "heredar_roles_de_quincena_anterior",
    "corregir_formula_subtotal", "ajustar_fechas_del_periodo", "respetar_lo_ya_escrito",
    "corregir_cedulas_en_el_formato",
    "umbral_candidato_extras_salida", "umbral_candidato_extras_horas",
}

ETIQUETAS = {
    "carpeta_de_resultados": "Carpeta de resultados",
    "jornal_dia_completo": "Jornal del día completo",
    "tarifa_hora_semana": "Hora entre semana",
    "tarifa_hora_finde": "Hora de fin de semana",
    "ajuste_media_hora": "Ajuste de la media hora",
    "sabado_estandar_horas": "Duración del sábado normal",
    "sabado_estandar_valor": "Monto del sábado normal",
    "horas_jornada_completa": "Horas para pagar día completo",
    "redondeo_horas": "Redondeo de horas",
    "minutos_marca_duplicada": "Minutos para marca repetida",
    "ventanas.fin_entrada": "Fin de la ventana de entrada",
    "ventanas.fin_alm_out": "Fin de la salida al almuerzo",
    "ventanas.fin_mediodia": "Fin del mediodía",
    "ventanas.fin_alm_in": "Fin del retorno del almuerzo",
    "marca_afiliado_presente": "Marca del afiliado presente",
    "marca_ausencia": "Marca de ausencia",
    "reconstruir_marca_faltante": "Reconstruir la marca que falte",
    "escribir_ausencias": "Escribir ausencias",
    "escribir_sabados": "Calcular sábados",
    "escribir_domingos": "Calcular domingos",
    "escribir_extras": "Llenar la columna de extras",
    "heredar_roles_de_quincena_anterior": "Heredar roles de la quincena anterior",
    "corregir_formula_subtotal": "Corregir la fórmula de subtotal",
    "ajustar_fechas_del_periodo": "Ajustar las fechas del período",
    "corregir_cedulas_en_el_formato": "Completar las cédulas en el formato",
    "respetar_lo_ya_escrito": "Respetar lo que el formato ya trae",
    "umbral_candidato_extras_salida": "Hora de salida que sugiere extra",
    "umbral_candidato_extras_horas": "Horas netas que sugieren extra",
}

_SI_NO = {"SI": True, "SÍ": True, "TRUE": True, "1": True, "VERDADERO": True,
          "NO": False, "FALSE": False, "0": False, "FALSO": False}


def _convertir(valor, tipo):
    if tipo == "si_no":
        if isinstance(valor, bool):
            return valor
        return _SI_NO.get(str(valor).strip().upper(), False)
    if tipo == "entero":
        return int(float(valor))
    if tipo == "numero":
        return float(valor)
    if tipo == "hora":
        import datetime as dt
        if isinstance(valor, dt.time):
            return valor.strftime("%H:%M")
        return str(valor).strip()
    return str(valor).strip()


def _estilo_cabecera(ws, anchos):
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="FF1E5F58")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for col, ancho in zip("ABCDEFGH", anchos):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A2"


# --------------------------------------------------------------------------- #
# configuracion
# --------------------------------------------------------------------------- #
def config_por_defecto():
    cfg = {}
    for clave, valor, tipo, _ in PARAMETROS:
        _asignar(cfg, clave, _convertir(valor, tipo))
    return cfg


def _asignar(cfg, clave, valor):
    if "." in clave:
        grupo, hijo = clave.split(".", 1)
        cfg.setdefault(grupo, {})[hijo] = valor
    else:
        cfg[clave] = valor


def guardar_config(cfg, ruta=ARCHIVO_CONFIG):
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "parametros"
    ws.append(["parametro", "valor", "tipo", "que controla"])
    for clave, defecto, tipo, desc in PARAMETROS:
        if "." in clave:
            grupo, hijo = clave.split(".", 1)
            valor = cfg.get(grupo, {}).get(hijo, defecto)
        else:
            valor = cfg.get(clave, defecto)
        if tipo == "si_no":
            valor = "SI" if valor else "NO"
        ws.append([clave, valor, tipo, desc])
    _estilo_cabecera(ws, (34, 14, 10, 86))
    for fila in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        fila[0].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(ruta)
    return ruta


def cargar_config(ruta=ARCHIVO_CONFIG):
    ruta = Path(ruta)
    if not ruta.exists():
        cfg = config_por_defecto()
        guardar_config(cfg, ruta)
        return cfg
    tipos = {c: t for c, _, t, _ in PARAMETROS}
    cfg = config_por_defecto()
    ws = openpyxl.load_workbook(ruta, data_only=True)["parametros"]
    for clave, valor, *_ in ws.iter_rows(min_row=2, values_only=True):
        if clave is None or clave not in tipos or valor is None:
            continue
        try:
            _asignar(cfg, str(clave).strip(), _convertir(valor, tipos[str(clave).strip()]))
        except (TypeError, ValueError):
            pass  # valor ilegible: se queda el de por defecto
    return cfg


# --------------------------------------------------------------------------- #
# colaboradores
# --------------------------------------------------------------------------- #
def guardar_roles(registros, ruta=ARCHIVO_ROLES):
    """registros: lista de dicts con las claves de COLUMNAS."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "roles"
    ws.append(["Cedula", "Personal", "Rol", "Licencia", "Notas"])
    for r in registros:
        ws.append([r.get("cedula", ""), r.get("nombres", ""),
                   r.get("rol", ""), r.get("licencia", ""), r.get("notas", "")])
    for celda in ws["A"][1:]:
        celda.number_format = "@"      # texto: protege el 0 inicial en Excel
    _estilo_cabecera(ws, (14, 38, 11, 14, 40))
    ws.auto_filter.ref = f"A1:E{max(ws.max_row, 1)}"

    ultima = max(ws.max_row, 2)
    for col, opciones in (("C", '"G,Q,TRIUNFO"'),
                          ("D", '"PM,IESS,MSP,VACACIONES"')):
        dv = DataValidation(type="list", formula1=opciones, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{ultima + 200}")

    guia = wb.create_sheet("instructivo")
    for fila in [
        ["Columna", "Para que sirve"],
        ["Cedula", "10 digitos. Si Excel le quita el 0 inicial, la aplicacion lo repone."],
        ["Personal", "Solo de referencia, para reconocer a la persona."],
        ["Rol", "G = guardia, Q = quemador, TRIUNFO = enviado a esa sucursal."],
        ["", "Con rol, la celda del dia lleva siempre la letra."],
        ["Licencia", "PM, IESS, MSP o VACACIONES. No se escribe sola: avisa en la revision."],
        ["Notas", "Texto libre. La aplicacion no lo usa."],
        ["", ""],
        ["Ojo", "Aqui NO va el listado de personal. Ese sale de la hoja del formato"],
        ["", "que viene dentro del archivo del biometrico, y cambia cada quincena."],
    ]:
        guia.append(fila)
    _estilo_cabecera(guia, (16, 96))
    for fila in guia.iter_rows(min_row=2, min_col=2, max_col=2):
        fila[0].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(ruta)
    return ruta


def cargar_roles(ruta=ARCHIVO_ROLES):
    """-> lista de dicts. Las cedulas se normalizan al leerlas."""
    import motor
    ruta = Path(ruta)
    if not ruta.exists():
        guardar_roles([], ruta)
        return []
    ws = openpyxl.load_workbook(ruta, data_only=True)["roles"]
    out = []
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if not fila or not any(str(v).strip() for v in fila[:2] if v is not None):
            continue
        v = list(fila) + [None] * (len(COLUMNAS) - len(fila))
        out.append({
            "cedula": motor.normalizar_cedula(v[0]),
            "nombres": str(v[1] or "").strip(),
            "rol": str(v[2] or "").strip().upper(),
            "licencia": str(v[3] or "").strip().upper(),
            "notas": str(v[4] or "").strip(),
        })
    return out


def indice(registros):
    """-> {cedula: ficha}, para pasarselo al motor."""
    return {r["cedula"]: r for r in registros if r["cedula"]}


def roles_desde_quincena(archivo, hoja=None):
    """Extrae los roles de una quincena ya llena: ahi las letras G, Q y TRIUNFO
    estan escritas dia a dia, asi que no hay que mantenerlas a mano."""
    import motor

    wb = openpyxl.load_workbook(archivo, data_only=False)
    hojas = [hoja] if hoja else wb.sheetnames
    mejor, mejor_n = [], 0
    for nombre in hojas:
        try:
            filas, _ = motor.leer_formato(wb[nombre])
        except Exception:
            continue
        cols = [5 + 2 * k for k in range(16)]
        encontrados = []
        for f in filas:
            if not f["cedula"]:
                continue
            vistos = {}
            for c in cols:
                v = str(wb[nombre].cell(f["fila"], c).value or "").strip().upper()
                if v in motor.ROLES + motor.LICENCIAS:
                    vistos[v] = vistos.get(v, 0) + 1
            if not vistos:
                continue
            cod, n = max(vistos.items(), key=lambda kv: kv[1])
            if n < 8:
                continue
            encontrados.append({
                "cedula": f["cedula"], "nombres": f["nombre"],
                "rol": cod if cod in motor.ROLES else "",
                "licencia": cod if cod in motor.LICENCIAS else "",
                "notas": f"tomado de «{nombre}»",
            })
        if len(encontrados) > mejor_n:
            mejor, mejor_n = encontrados, len(encontrados)
    wb.close()
    return mejor
