# -*- coding: utf-8 -*-
from __future__ import annotations

import calendar
import json
import random
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("No se encontró openpyxl. Ejecuta: pip install openpyxl")
    raise

# ============================================================
# 1. CONFIGURACION GENERAL
# ============================================================

CARPETA = Path(__file__).resolve().parent
ARCHIVO_ESTADO = CARPETA / "ESTADO_ROL_K9.json"

FECHA_BASE = date(2026, 11, 30)
GRUPOS_BASE_30_NOV = ("G3", "G2")

PATRON = [
    ("G1", "G4"),
    ("G2", "G1"),
    ("G3", "G2"),
    ("G4", "G3"),
]

GRUPOS = {
    "G1": ["RAMOS", "TOMAIRO", "FLORES", "PALOMINO", "LLONTOP", "OYOLA"],
    "G2": ["MIRANDA", "NINA", "CUEVA", "INGA", "NUÑEZ", "CABALLERO"],
    "G3": ["TARAZONA", "PRADA", "MAYORGA", "ROJAS", "OLIVER", "YUCRA"],
    "G4": ["PEREZ", "ESPINO", "MOSCOSO", "PATIÑO", "CAMACHO", "CONCHA"],
}

EXCLUIDOS = {"GUTIERREZ", "CIUDAD", "LUQUE"}
NO_COUNTER = {"PEREZ", "ROJAS", "ESPINO"}
NO_SERPOST = {"MIRANDA"}

HISTORIAL_COUNTER_INICIAL = {
    "2026-10": {
        "G1": ["RAMOS", "LLONTOP"],
        "G2": ["MIRANDA", "NUÑEZ"],
        "G3": ["PRADA", "TARAZONA"],
        "G4": ["CAMACHO", "MOSCOSO"],
    },
    "2026-11": {
        "G1": ["TOMAIRO", "OYOLA"],
        "G2": ["INGA", "NINA"],
        "G3": ["OLIVER", "YUCRA"],
        "G4": ["CONCHA", "PATIÑO"],
    },
}

ESPECIAL_INICIAL = {
    "serpost": "ROJAS",
    "acompanante": "ESPINO",
}

VACACIONES = {
    "FLORES": [(date(2026, 11, 17), date(2026, 12, 1))],
}

COLUMNAS_DIA = [
    "MARITIMA DIA 1",
    "MARITIMA DIA 2",
    "PATRULLA / VETERINARIA",
    "PATRULLA ITINERANTE",
    "PATRULLA / DHL",
    "CARGA AEREA",
    "SERPOST (DHL)",
]

COLUMNAS_COUNTER = [
    "COUNTER DIA 06-18",
    "COUNTER DIA 08-20",
    "COUNTER NOCHE 20-08",
    "COUNTER NOCHE 18-06",
]

COLUMNAS_EXTRA = [
    "COUNTER ACOMPANANTE",
    "DISPONIBLES TURNO DIA",
    "DISPONIBLES TURNO NOCHE",
]

COLUMNAS = [
    "DIA", "FECHA", "GRUPO TURNO DIA", "GRUPO TURNO NOCHE"
] + COLUMNAS_DIA + COLUMNAS_COUNTER + COLUMNAS_EXTRA

GUIAS = [
    persona
    for grupo in GRUPOS
    for persona in GRUPOS[grupo]
    if persona not in EXCLUIDOS
]


# ============================================================
# 2. FECHAS Y MEMORIA
# ============================================================

def mes_nombre(mes: int) -> str:
    nombres = [
        "", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO",
        "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE",
        "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ]
    return nombres[mes]


def clave_mes(anio: int, mes: int) -> str:
    return f"{anio:04d}-{mes:02d}"


def fecha_a_texto(fecha: date) -> str:
    return fecha.isoformat()


def texto_a_fecha(valor: str) -> date:
    return date.fromisoformat(valor)


def ultimo_dia_mes(anio: int, mes: int) -> date:
    return date(anio, mes, calendar.monthrange(anio, mes)[1])


def dias_mes(anio: int, mes: int):
    for dia in range(1, calendar.monthrange(anio, mes)[1] + 1):
        yield date(anio, mes, dia)


def nombre_dia(fecha: date) -> str:
    return [
        "LUNES", "MARTES", "MIERCOLES", "JUEVES",
        "VIERNES", "SABADO", "DOMINGO"
    ][fecha.weekday()]


def crear_estado_inicial() -> dict:
    return {
        "version": 1,
        "base": {
            "fecha": fecha_a_texto(FECHA_BASE),
            "grupo_dia": GRUPOS_BASE_30_NOV[0],
            "grupo_noche": GRUPOS_BASE_30_NOV[1],
            "descripcion": "Último día del rol K9 de noviembre 2026",
        },
        "meses": {},
        "historial_counter": HISTORIAL_COUNTER_INICIAL,
        "especial": {
            "ultimo_serpost": ESPECIAL_INICIAL["serpost"],
            "ultimo_acompanante": ESPECIAL_INICIAL["acompanante"],
        },
    }


def cargar_estado() -> dict:
    if not ARCHIVO_ESTADO.exists():
        estado = crear_estado_inicial()
        guardar_estado(estado)
        return estado

    try:
        with ARCHIVO_ESTADO.open("r", encoding="utf-8") as archivo:
            estado = json.load(archivo)
    except (json.JSONDecodeError, OSError) as exc:
        raise RuntimeError(
            f"No se pudo leer {ARCHIVO_ESTADO.name}: {exc}. "
            "No borres el archivo sin hacer una copia."
        )

    estado.setdefault("meses", {})
    estado.setdefault(
        "historial_counter",
        HISTORIAL_COUNTER_INICIAL.copy()
    )
    estado.setdefault("especial", ESPECIAL_INICIAL.copy())
    estado.setdefault("base", crear_estado_inicial()["base"])

    return estado


def guardar_estado(estado: dict) -> None:
    temporal = ARCHIVO_ESTADO.with_suffix(".tmp")

    with temporal.open("w", encoding="utf-8") as archivo:
        json.dump(estado, archivo, ensure_ascii=False, indent=2)

    temporal.replace(ARCHIVO_ESTADO)


def grupo_en_fecha(fecha: date, estado: dict) -> tuple[str, str]:
    base = texto_a_fecha(estado["base"]["fecha"])
    diferencia = (fecha - base).days

    indice_base = PATRON.index(
        (estado["base"]["grupo_dia"], estado["base"]["grupo_noche"])
    )

    indice = (indice_base + diferencia) % len(PATRON)
    return PATRON[indice]


def fase_grupo(fecha: date, grupo: str, estado: dict) -> str:
    dia, noche = grupo_en_fecha(fecha, estado)

    if grupo == dia:
        return "DIA"

    if grupo == noche:
        return "NOCHE"

    return "DESCANSO"


def esta_vacaciones(persona: str, fecha: date) -> bool:
    for inicio, fin in VACACIONES.get(persona, []):
        if inicio <= fecha <= fin:
            return True

    return False


def personas_activas(grupo: str, fecha: date) -> list[str]:
    return [
        persona
        for persona in GRUPOS[grupo]
        if persona not in EXCLUIDOS
        and not esta_vacaciones(persona, fecha)
    ]


def personal_turno(fecha: date, estado: dict) -> dict:
    grupo_dia, grupo_noche = grupo_en_fecha(fecha, estado)

    return {
        "DIA": personas_activas(grupo_dia, fecha),
        "NOCHE": personas_activas(grupo_noche, fecha),
    }


def mes_anterior(
    anio: int, mes: int, saltos: int = 1
) -> tuple[int, int]:
    indice = anio * 12 + (mes - 1) - saltos
    return indice // 12, indice % 12 + 1


# ============================================================
# 3. HISTORIAL Y ROTACION COUNTER
# ============================================================

def titulares_counter_anteriores(
    estado: dict,
    anio: int,
    mes: int,
    grupo: str,
    meses_descanso: int = 2
) -> set[str]:

    prohibidos = set()

    for salto in range(1, meses_descanso + 1):
        a, m = mes_anterior(anio, mes, salto)

        registro = estado["historial_counter"].get(
            clave_mes(a, m), {}
        )

        prohibidos.update(registro.get(grupo, []))

    return prohibidos


def ultimo_mes_counter(
    estado: dict,
    persona: str,
    grupo: str,
    anio: int,
    mes: int
) -> int:

    mejor = 999

    for clave, grupos in estado["historial_counter"].items():
        try:
            a, m = map(int, clave.split("-"))
        except ValueError:
            continue

        if (a, m) >= (anio, mes):
            continue

        if persona in grupos.get(grupo, []):
            distancia = (anio * 12 + mes) - (a * 12 + m)
            mejor = min(mejor, distancia)

    return mejor


def seleccionar_counter(
    estado: dict,
    anio: int,
    mes: int,
    rng: random.Random
) -> dict:

    seleccion = {}

    for grupo, lista in GRUPOS.items():

        elegibles = [
            persona
            for persona in lista
            if persona not in EXCLUIDOS
            and persona not in NO_COUNTER
        ]

        vetados = titulares_counter_anteriores(
            estado, anio, mes, grupo, meses_descanso=2
        )

        frescos = [
            persona for persona in elegibles
            if persona not in vetados
        ]

        # Primero se respeta el descanso de dos meses.
        # Si no hay suficientes candidatos, se recurre a los
        # que llevan más tiempo sin realizar Counter.
        candidatos = frescos if len(frescos) >= 2 else elegibles

        candidatos = sorted(
            candidatos,
            key=lambda persona: (
                -ultimo_mes_counter(
                    estado, persona, grupo, anio, mes
                ),
                rng.random()
            )
        )

        if len(candidatos) < 2:
            raise RuntimeError(
                f"No hay dos conductores Counter en {grupo}."
            )

        seleccion[grupo] = candidatos[:2]

    estado["historial_counter"][clave_mes(anio, mes)] = seleccion

    return seleccion


# ============================================================
# 4. SERPOST, CARGA AEREA Y MARITIMA
# ============================================================

def especiales_mes(estado: dict) -> tuple[str, str]:
    anterior_serpost = estado["especial"].get(
        "ultimo_serpost", "ESPINO"
    )

    anterior_acomp = estado["especial"].get(
        "ultimo_acompanante", "ROJAS"
    )

    return anterior_serpost, anterior_acomp


def actualizar_especiales(
    estado: dict,
    serpost: str,
    acompanante: str
) -> None:

    estado["especial"]["ultimo_serpost"] = serpost
    estado["especial"]["ultimo_acompanante"] = acompanante


def seleccionar_asignaciones_fijas_mes(
    titulares: dict,
    rng: random.Random
) -> dict:

    serpost = ["ROJAS", "ESPINO"]

    asignaciones = {
        "serpost": serpost,
        "carga_aerea": [],
        "maritima": {}
    }

    usados_especiales = set(serpost)
    disponibles_por_grupo = {}

    # Marítima: dos personas fijas por grupo durante todo el mes.
    for grupo, lista in GRUPOS.items():

        disponibles = [
            persona
            for persona in lista
            if persona not in EXCLUIDOS
            and persona not in titulares[grupo]
            and persona not in usados_especiales
        ]

        rng.shuffle(disponibles)

        if len(disponibles) < 2:
            raise RuntimeError(
                f"No hay dos personas disponibles para Marítima en {grupo}."
            )

        asignaciones["maritima"][grupo] = disponibles[:2]
        disponibles_por_grupo[grupo] = disponibles[2:]

    # Carga Aérea: dos personas fijas del G1.
    carga = disponibles_por_grupo["G1"]

    if len(carga) < 2:
        raise RuntimeError(
            "No hay dos personas libres en G1 para Carga Aérea."
        )

    asignaciones["carga_aerea"] = carga[:2]

    return asignaciones


FECHA_INICIO_ALTERNANCIA = date(2026, 12, 1)


def persona_de_bloque(pareja: list[str], fecha: date) -> str:
    dias_transcurridos = (
        fecha - FECHA_INICIO_ALTERNANCIA
    ).days

    indice_bloque = dias_transcurridos // 2

    return pareja[indice_bloque % 2]


# ============================================================
# 5. ASIGNADOR EQUITATIVO
# ============================================================

class Asignador:

    def __init__(self, semilla: int, conteos: Counter):
        self.rng = random.Random(semilla)
        self.conteos = conteos
        self.ultimo_puesto = {}

    def escoger(
        self,
        candidatos: list[str],
        asignados: set[str],
        puesto: str
    ) -> str:

        libres = [
            persona
            for persona in candidatos
            if persona not in asignados
        ]

        if not libres:
            return ""

        libres.sort(
            key=lambda persona: (
                self.conteos[persona],
                self.ultimo_puesto.get(persona) == puesto,
                self.rng.random(),
            )
        )

        elegido = libres[0]

        self.conteos[elegido] += 1
        self.ultimo_puesto[elegido] = puesto
        asignados.add(elegido)

        return elegido


# ============================================================
# 6. CREACION DE FILAS
# ============================================================

def inicializar_fila(fecha: date, estado: dict) -> dict:

    dia, noche = grupo_en_fecha(fecha, estado)

    fila = {
        "DIA": nombre_dia(fecha),
        "FECHA": fecha,
        "GRUPO TURNO DIA": dia,
        "GRUPO TURNO NOCHE": noche,
    }

    for columna in COLUMNAS_DIA + COLUMNAS_COUNTER + COLUMNAS_EXTRA:
        fila[columna] = ""

    return fila


# ============================================================
# 7. COUNTER: DOS DE DIA Y DOS DE NOCHE
# ============================================================

def asignar_counter_dia(
    fila: dict,
    titulares: dict,
    turnos: dict,
    asignados: set[str],
    conteos: Counter
) -> None:

    grupo_dia = fila["GRUPO TURNO DIA"]
    grupo_noche = fila["GRUPO TURNO NOCHE"]

    dia = [
        persona
        for persona in titulares[grupo_dia]
        if persona in turnos["DIA"]
    ]

    noche = [
        persona
        for persona in titulares[grupo_noche]
        if persona in turnos["NOCHE"]
    ]

    puestos = [
        ("COUNTER DIA 06-18", dia[0] if len(dia) > 0 else ""),
        ("COUNTER DIA 08-20", dia[1] if len(dia) > 1 else ""),
        ("COUNTER NOCHE 20-08", noche[0] if len(noche) > 0 else ""),
        ("COUNTER NOCHE 18-06", noche[1] if len(noche) > 1 else ""),
    ]

    for puesto, persona in puestos:
        if persona and persona not in asignados:
            fila[puesto] = persona
            asignados.add(persona)
            conteos[persona] += 1


# ============================================================
# 8. ASIGNACION DE SERVICIOS FIJOS
# ============================================================

def asignar_servicios_dia(
    fila: dict,
    fecha: date,
    turnos: dict,
    asignados: set[str],
    asignador: Asignador,
    asignaciones_fijas: dict
) -> None:

    grupo_dia = fila["GRUPO TURNO DIA"]

    # SERPOST: alternancia de dos días.
    serpost = persona_de_bloque(
        asignaciones_fijas["serpost"], fecha
    )

    if serpost not in asignados:
        fila["SERPOST (DHL)"] = serpost
        asignados.add(serpost)
        asignador.conteos[serpost] += 1

    # CARGA AEREA: alternancia de dos días.
    carga = persona_de_bloque(
        asignaciones_fijas["carga_aerea"], fecha
    )

    if carga not in asignados:
        fila["CARGA AEREA"] = carga
        asignados.add(carga)
        asignador.conteos[carga] += 1

    # MARITIMA: dos personas fijas del grupo que está de día.
    candidatos = turnos["DIA"]

    maritimos = [
        persona
        for persona in asignaciones_fijas["maritima"][grupo_dia]
        if persona in candidatos and persona not in asignados
    ]

    if len(maritimos) > 0:
        fila["MARITIMA DIA 1"] = maritimos[0]
        asignados.add(maritimos[0])
        asignador.conteos[maritimos[0]] += 1

    if len(maritimos) > 1:
        fila["MARITIMA DIA 2"] = maritimos[1]
        asignados.add(maritimos[1])
        asignador.conteos[maritimos[1]] += 1

    # CIUDAD permanece en Patrulla/Veterinaria de lunes a viernes.
    if fecha.weekday() < 5:
        fila["PATRULLA / VETERINARIA"] = "CIUDAD"


# ============================================================
# 9. PATRULLAS: SE ASIGNAN DESPUES DE LOS PUESTOS FIJOS Y COUNTER
# ============================================================

def asignar_patrullas_restantes(
    fila: dict,
    fecha: date,
    turnos: dict,
    asignados: set[str],
    asignador: Asignador
) -> None:

    candidatos = turnos["DIA"]

    fila["PATRULLA ITINERANTE"] = asignador.escoger(
        candidatos,
        asignados,
        "PATRULLA ITINERANTE"
    )

    if fecha.weekday() < 5:
        fila["PATRULLA / DHL"] = asignador.escoger(
            candidatos,
            asignados,
            "PATRULLA / DHL"
        )


# ============================================================
# 10. CONSTRUCCION DEL MES
# ============================================================

def construir_mes(
    anio: int,
    mes: int,
    estado: dict,
    semilla: int
) -> dict:

    rng = random.Random(semilla)

    conteos = Counter()

    asignador = Asignador(
        semilla ^ (anio * 100 + mes),
        conteos
    )

    titulares = seleccionar_counter(
        estado, anio, mes, rng
    )

    asignaciones_fijas = seleccionar_asignaciones_fijas_mes(
        titulares, rng
    )

    serpost_titular, acompanante = especiales_mes(estado)

    actualizar_especiales(
        estado,
        acompanante,
        serpost_titular
    )

    filas = []
    historial_turnos = []

    for fecha in dias_mes(anio, mes):

        fila = inicializar_fila(fecha, estado)

        turnos = personal_turno(fecha, estado)

        asignados = set()

        # Primero se reservan Serpost, Carga y Marítima.
        asignar_servicios_dia(
            fila,
            fecha,
            turnos,
            asignados,
            asignador,
            asignaciones_fijas
        )

        # Segundo se reservan los cuatro puestos obligatorios Counter.
        asignar_counter_dia(
            fila,
            titulares,
            turnos,
            asignados,
            conteos
        )

        # Tercero se asignan las patrullas con el personal restante.
        # Así una patrulla no puede quitarle un conductor a Counter.
        asignar_patrullas_restantes(
            fila,
            fecha,
            turnos,
            asignados,
            asignador
        )

        disponibles_dia = [
            persona
            for persona in turnos["DIA"]
            if persona not in asignados
        ]

        disponibles_noche = [
            persona
            for persona in turnos["NOCHE"]
            if persona not in asignados
        ]

        fila["DISPONIBLES TURNO DIA"] = ", ".join(
            disponibles_dia
        )

        fila["DISPONIBLES TURNO NOCHE"] = ", ".join(
            disponibles_noche
        )

        filas.append(fila)

        historial_turnos.append({
            "fecha": fecha_a_texto(fecha),
            "grupo_dia": fila["GRUPO TURNO DIA"],
            "grupo_noche": fila["GRUPO TURNO NOCHE"],
            "fase": "DIA -> NOCHE -> DESCANSO -> DESCANSO",
        })

    return {
        "anio": anio,
        "mes": mes,
        "nombre_mes": mes_nombre(mes),
        "fecha_inicio": fecha_a_texto(date(anio, mes, 1)),
        "fecha_fin": fecha_a_texto(ultimo_dia_mes(anio, mes)),
        "semilla": semilla,
        "titulares_counter": titulares,
        "asignaciones_fijas": asignaciones_fijas,
        "serpost_titular": serpost_titular,
        "acompanante": acompanante,
        "filas": [
            {
                **fila,
                "FECHA": fecha_a_texto(fila["FECHA"]),
            }
            for fila in filas
        ],
        "historial_turnos": historial_turnos,
        "conteos": dict(conteos),
    }


# ============================================================
# 11. VALIDACION DEL CICLO 2X2
# ============================================================

def validar_ciclo(
    estado: dict,
    anio: int,
    mes: int
) -> list[str]:

    errores = []

    for fecha in dias_mes(anio, mes):

        for grupo in GRUPOS:

            fases = [
                fase_grupo(
                    fecha + timedelta(days=i),
                    grupo,
                    estado
                )
                for i in range(4)
            ]

            activo = [
                fase != "DESCANSO"
                for fase in fases
            ]

            patrones_validos = (
                [True, True, False, False],
                [True, False, False, True],
                [False, False, True, True],
                [False, True, True, False],
            )

            if tuple(activo) not in [
                tuple(patron)
                for patron in patrones_validos
            ]:
                errores.append(
                    f"Ciclo 2x2 incorrecto para {grupo} desde {fecha}."
                )

    return errores


# ============================================================
# 12. VALIDACION COMPLETA
# ============================================================

def validar_mes(
    datos: dict,
    estado: dict
) -> list[str]:

    errores = []

    anio = datos["anio"]
    mes = datos["mes"]

    if len(datos["filas"]) != calendar.monthrange(anio, mes)[1]:
        errores.append("La cantidad de días del mes no coincide.")

    errores.extend(
        validar_ciclo(estado, anio, mes)
    )

    for fila in datos["filas"]:

        fecha = texto_a_fecha(fila["FECHA"])

        dia, noche = grupo_en_fecha(fecha, estado)

        if (
            fila["GRUPO TURNO DIA"] != dia
            or fila["GRUPO TURNO NOCHE"] != noche
        ):
            errores.append(
                f"Grupos incorrectos el {fecha}."
            )

        # Los cuatro Counter deben tener personal.
        for puesto in COLUMNAS_COUNTER:
            if not str(fila.get(puesto, "") or "").strip():
                errores.append(
                    f"Falta personal en {puesto} el {fecha}."
                )

        # Verificar que no haya personas duplicadas.
        ocupados = []

        for columna in COLUMNAS_DIA + COLUMNAS_COUNTER + [
            "COUNTER ACOMPANANTE"
        ]:

            valor = str(
                fila.get(columna, "") or ""
            ).strip()

            if not valor:
                continue

            persona = valor.replace("(APOYO)", "").strip()

            for nombre in persona.split("/"):
                nombre = nombre.strip()

                if nombre in GUIAS:
                    ocupados.append(nombre)

        duplicados = [
            persona
            for persona, cantidad in Counter(ocupados).items()
            if cantidad > 1
        ]

        if duplicados:
            errores.append(
                f"Un guía tiene dos puestos el {fecha}: "
                f"{', '.join(duplicados)}"
            )

        activos = set(GRUPOS[dia] + GRUPOS[noche]) - EXCLUIDOS

        personal_servicio_fijo = {
            fila.get("SERPOST (DHL)", ""),
            fila.get("CARGA AEREA", ""),
        }

        for persona in ocupados:

            if (
                persona not in activos
                and persona not in personal_servicio_fijo
            ):
                errores.append(
                    f"{persona} aparece asignado en su día de descanso ({fecha})."
                )

            if esta_vacaciones(persona, fecha):
                errores.append(
                    f"{persona} aparece trabajando durante vacaciones ({fecha})."
                )

        if "MIRANDA" in str(fila.get("SERPOST (DHL)", "")):
            errores.append(
                f"Miranda aparece en Serpost el {fecha}."
            )

        # Serpost debe respetar la alternancia cada dos días.
        fijas = datos["asignaciones_fijas"]

        serpost_esperado = persona_de_bloque(
            fijas["serpost"], fecha
        )

        serpost_real = str(
            fila.get("SERPOST (DHL)", "") or ""
        ).strip()

        if serpost_real != serpost_esperado:
            errores.append(
                f"Serpost no respeta la alternancia el {fecha}."
            )

        # Carga Aérea debe respetar la alternancia cada dos días.
        carga_esperada = persona_de_bloque(
            fijas["carga_aerea"], fecha
        )

        carga_real = str(
            fila.get("CARGA AEREA", "") or ""
        ).strip()

        if carga_real != carga_esperada:
            errores.append(
                f"Carga Aérea no respeta la alternancia el {fecha}."
            )

        # Marítima conserva los mismos dos titulares por grupo.
        grupo_dia = fila["GRUPO TURNO DIA"]

        maritimos_esperados = fijas["maritima"][grupo_dia]

        maritimos_reales = [
            str(fila.get("MARITIMA DIA 1", "") or "").strip(),
            str(fila.get("MARITIMA DIA 2", "") or "").strip(),
        ]

        if maritimos_reales != maritimos_esperados:
            errores.append(
                f"Marítima no conserva su pareja fija el {fecha}."
            )

        # CIUDAD permanece en Veterinaria de lunes a viernes.
        veterinaria = str(
            fila.get("PATRULLA / VETERINARIA", "") or ""
        ).strip()

        if fecha.weekday() < 5 and veterinaria != "CIUDAD":
            errores.append(
                f"Veterinaria debe tener a CIUDAD el {fecha}."
            )

        if fecha.weekday() >= 5 and veterinaria:
            errores.append(
                f"Veterinaria no debe asignarse el fin de semana ({fecha})."
            )

    # Dos titulares mensuales por cada grupo.
    for grupo, personas in datos["titulares_counter"].items():

        if len(personas) != 2:
            errores.append(
                f"{grupo} no tiene dos titulares Counter."
            )

        for persona in personas:
            if persona in NO_COUNTER:
                errores.append(
                    f"{persona} está restringido de Counter."
                )

    fijas = datos["asignaciones_fijas"]

    if len(fijas["serpost"]) != 2:
        errores.append(
            "Serpost debe tener exactamente dos personas al mes."
        )

    if len(fijas["carga_aerea"]) != 2:
        errores.append(
            "Carga Aérea debe tener exactamente dos personas al mes."
        )

    for grupo in GRUPOS:
        if len(fijas["maritima"].get(grupo, [])) != 2:
            errores.append(
                f"Marítima debe tener dos personas fijas en {grupo}."
            )

    return errores


# ============================================================
# 13. FORMATO EXCEL
# ============================================================

AZUL_OSCURO = "17365D"
AZUL = "4472C4"
CELESTE = "D9EAF7"
GRIS = "E7E6E6"
BLANCO = "FFFFFF"


def estilo_hoja(ws):

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "E3"
    ws.auto_filter.ref = ws.dimensions

    for fila in ws.iter_rows():

        for celda in fila:

            celda.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            celda.border = Border(
                bottom=Side(
                    style="hair",
                    color="C7D4E0"
                )
            )

    for celda in ws[1]:

        celda.fill = PatternFill(
            "solid",
            fgColor=AZUL_OSCURO
        )

        celda.font = Font(
            bold=True,
            color=BLANCO,
            size=13
        )

    if ws.max_row >= 2:

        for celda in ws[2]:

            celda.fill = PatternFill(
                "solid",
                fgColor=AZUL
            )

            celda.font = Font(
                bold=True,
                color=BLANCO
            )

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 40

    for columna in range(1, ws.max_column + 1):

        letra = get_column_letter(columna)
        ancho = 13

        for fila in range(1, ws.max_row + 1):

            valor = ws.cell(fila, columna).value

            if valor is not None:
                ancho = max(
                    ancho,
                    min(36, len(str(valor)) + 2)
                )

        ws.column_dimensions[letra].width = min(ancho, 36)


# ============================================================
# 14. CREAR EXCEL
# ============================================================

def crear_excel(datos: dict) -> Path:

    wb = Workbook()

    ws = wb.active
    ws.title = "ROL K9"

    ws.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=len(COLUMNAS)
    )

    ws.cell(
        1,
        1,
        f"ROL DE SERVICIO K9 - {datos['nombre_mes']} {datos['anio']}"
    )

    ws.append(COLUMNAS)

    for fila in datos["filas"]:
        ws.append([
            fila.get(columna, "")
            for columna in COLUMNAS
        ])

    for fila in range(3, ws.max_row + 1):

        fecha = ws.cell(fila, 2).value

        if isinstance(fecha, str):
            ws.cell(fila, 2).value = date.fromisoformat(fecha)

        ws.cell(fila, 2).number_format = "dd/mm/yyyy"

        if ws.cell(fila, 2).value.weekday() >= 5:

            for col in range(1, ws.max_column + 1):
                ws.cell(fila, col).fill = PatternFill(
                    "solid",
                    fgColor="F2F2F2"
                )

    estilo_hoja(ws)

    # RESUMEN DE SERVICIOS
    resumen = wb.create_sheet("RESUMEN")

    resumen.append([
        f"RESUMEN DE SERVICIOS - {datos['nombre_mes']} {datos['anio']}"
    ])

    resumen.append([
        "GUIA",
        "GRUPO",
        "SERVICIOS",
        "TITULAR COUNTER"
    ])

    grupo_persona = {
        persona: grupo
        for grupo, personas in GRUPOS.items()
        for persona in personas
    }

    titulares = {
        persona
        for lista in datos["titulares_counter"].values()
        for persona in lista
    }

    for persona in GUIAS:

        resumen.append([
            persona,
            grupo_persona[persona],
            datos["conteos"].get(persona, 0),
            "SI" if persona in titulares else "NO",
        ])

    estilo_hoja(resumen)
    resumen.freeze_panes = "A3"

    # HISTORIAL
    hist = wb.create_sheet("HISTORIAL")

    hist.append(["HISTORIAL DEL ROL K9"])
    hist.append([
        "FECHA",
        "GRUPO DIA",
        "GRUPO NOCHE",
        "CICLO"
    ])

    for item in datos["historial_turnos"]:

        hist.append([
            date.fromisoformat(item["fecha"]),
            item["grupo_dia"],
            item["grupo_noche"],
            item["fase"],
        ])

        hist.cell(
            hist.max_row, 1
        ).number_format = "dd/mm/yyyy"

    hist.append([])
    hist.append(["GRUPO", "TITULARES COUNTER"])

    for grupo, personas in datos["titulares_counter"].items():
        hist.append([
            grupo,
            ", ".join(personas)
        ])

    hist.append([])
    hist.append(["ASIGNACIONES FIJAS DEL MES"])

    fijas = datos["asignaciones_fijas"]

    hist.append([
        "SERPOST (2 PERSONAS)",
        " / ".join(fijas["serpost"])
    ])

    hist.append([
        "CARGA AEREA (2 PERSONAS)",
        " / ".join(fijas["carga_aerea"])
    ])

    for grupo, personas in fijas["maritima"].items():
        hist.append([
            f"MARITIMA {grupo}",
            " / ".join(personas)
        ])

    hist.append([
        "PATRULLA / VETERINARIA L-V",
        "CIUDAD (PERMANENTE)"
    ])

    hist.append([
        "ACOMPANANTE COUNTER (REFERENCIA)",
        datos["acompanante"]
    ])

    hist.append(["SEMILLA", datos["semilla"]])

    estilo_hoja(hist)
    hist.freeze_panes = "A3"

    # CONFIGURACION
    config = wb.create_sheet("CONFIGURACION")

    config.append(["CONFIGURACION DEL GENERADOR K9"])
    config.append(["ELEMENTO", "VALOR"])

    config.append([
        "Ciclo",
        "DIA -> NOCHE -> DESCANSO -> DESCANSO"
    ])

    config.append([
        "Base",
        "30/11/2026 - G3/G2"
    ])

    config.append([
        "Continuidad",
        "Se carga desde ESTADO_ROL_K9.json"
    ])

    config.append([
        "Excluidos",
        ", ".join(sorted(EXCLUIDOS))
    ])

    config.append([
        "No Counter",
        ", ".join(sorted(NO_COUNTER))
    ])

    config.append([
        "No Serpost",
        ", ".join(sorted(NO_SERPOST))
    ])

    config.append([
        "Veterinaria",
        "CIUDAD permanente de lunes a viernes"
    ])

    config.append([
        "Carga Aérea",
        "2 personas fijas; alternan cada 2 días, incluidos fines de semana"
    ])

    config.append([
        "Serpost",
        "2 personas fijas; alternan cada 2 días, incluidos fines de semana"
    ])

    config.append([
        "Marítima",
        "2 personas fijas por grupo durante todo el mes"
    ])

    config.append([])
    config.append(["GRUPO", "GUIAS"])

    for grupo, personas in GRUPOS.items():
        config.append([
            grupo,
            ", ".join(personas)
        ])

    estilo_hoja(config)
    config.freeze_panes = "A3"

    nombre = (
        f"ROL_K9_{datos['anio']}_{datos['mes']:02d}_"
        f"{datos['nombre_mes']}.xlsx"
    )

    ruta = CARPETA / nombre
    wb.save(ruta)

    return ruta


# ============================================================
# 15. GUARDADO Y CONTINUIDAD MENSUAL
# ============================================================

def guardar_mes_en_estado(
    estado: dict,
    datos: dict
) -> None:

    clave = clave_mes(
        datos["anio"],
        datos["mes"]
    )

    estado["meses"][clave] = datos
    estado["ultimo_mes_generado"] = clave
    estado["ultima_fecha_generada"] = datos["fecha_fin"]

    guardar_estado(estado)


def mes_guardado(
    estado: dict,
    anio: int,
    mes: int
) -> dict | None:

    return estado.get("meses", {}).get(
        clave_mes(anio, mes)
    )


def buscar_mes_previo(
    estado: dict,
    anio: int,
    mes: int
) -> tuple[int, int] | None:

    objetivo = anio * 12 + mes
    encontrados = []

    for clave in estado.get("meses", {}):

        try:
            a, m = map(int, clave.split("-"))
        except ValueError:
            continue

        indice = a * 12 + m

        if indice < objetivo:
            encontrados.append((indice, a, m))

    if not encontrados:
        return None

    _, a, m = max(encontrados)

    return a, m


def validar_continuidad_mes(
    estado: dict,
    anio: int,
    mes: int
) -> None:

    previo = buscar_mes_previo(
        estado, anio, mes
    )

    if previo is None:

        if (anio, mes) != (2026, 12):
            raise RuntimeError(
                "No hay un mes previo guardado. "
                "Primero genera diciembre de 2026."
            )

        return

    pa, pm = previo

    esperado = date(
        pa,
        pm,
        calendar.monthrange(pa, pm)[1]
    ) + timedelta(days=1)

    solicitado = date(anio, mes, 1)

    if solicitado != esperado:
        raise RuntimeError(
            f"No se puede saltar de {mes_nombre(pm)} {pa} "
            f"a {mes_nombre(mes)} {anio}. "
            "Primero genera el mes inmediato siguiente."
        )


def actualizar_base_al_ultimo_dia(
    estado: dict,
    datos: dict
) -> None:

    ultima_fila = datos["filas"][-1]

    estado["base"] = {
        "fecha": ultima_fila["FECHA"],
        "grupo_dia": ultima_fila["GRUPO TURNO DIA"],
        "grupo_noche": ultima_fila["GRUPO TURNO NOCHE"],
        "descripcion": "Último día del último mes generado y guardado",
    }


def reconstruir_estado_antes_de_generar(
    estado: dict,
    anio: int,
    mes: int
) -> dict:

    objetivo = anio * 12 + mes
    conservar = {}

    for clave, datos_mes in estado.get("meses", {}).items():

        a, m = map(int, clave.split("-"))

        if a * 12 + m < objetivo:
            conservar[clave] = datos_mes

    estado["meses"] = conservar

    nuevo_historial = {}

    for clave, valor in estado.get(
        "historial_counter", {}
    ).items():

        a, m = map(int, clave.split("-"))

        if a * 12 + m < objetivo:
            nuevo_historial[clave] = valor

    estado["historial_counter"] = nuevo_historial

    if conservar:

        clave_ultimo = max(
            conservar,
            key=lambda k: tuple(map(int, k.split("-")))
        )

        ultimo = conservar[clave_ultimo]
        ultima_fila = ultimo["filas"][-1]

        estado["base"] = {
            "fecha": ultima_fila["FECHA"],
            "grupo_dia": ultima_fila["GRUPO TURNO DIA"],
            "grupo_noche": ultima_fila["GRUPO TURNO NOCHE"],
            "descripcion": "Último mes previo conservado",
        }

        estado["especial"] = {
            "ultimo_serpost": ultimo["serpost_titular"],
            "ultimo_acompanante": ultimo["acompanante"],
        }

    else:

        inicial = crear_estado_inicial()

        estado["base"] = inicial["base"]
        estado["especial"] = inicial["especial"]
        estado["historial_counter"] = inicial["historial_counter"]

    return estado


# ============================================================
# 16. INGRESO DEL MES Y AÑO
# ============================================================

def preguntar_mes_anio() -> tuple[int, int]:

    print("\n" + "=" * 64)
    print("        GENERADOR AUTOMATICO DEL ROL K9")
    print("        SOLO GUIAS Y PERROS - NO OPERARIOS")
    print("=" * 64)

    texto_mes = input(
        "Mes (nombre o número): "
    ).strip().upper()

    texto_anio = input(
        "Año (por ejemplo 2026): "
    ).strip()

    if not texto_anio:
        raise ValueError("Debes indicar el año.")

    anio = int(texto_anio)

    mapa = {
        "ENERO": 1,
        "FEBRERO": 2,
        "MARZO": 3,
        "ABRIL": 4,
        "MAYO": 5,
        "JUNIO": 6,
        "JULIO": 7,
        "AGOSTO": 8,
        "SEPTIEMBRE": 9,
        "SETIEMBRE": 9,
        "OCTUBRE": 10,
        "NOVIEMBRE": 11,
        "DICIEMBRE": 12,
    }

    if texto_mes.isdigit():
        mes = int(texto_mes)
    else:
        mes = mapa.get(texto_mes)

    if mes is None or mes < 1 or mes > 12:
        raise ValueError(
            "Mes no válido. Escribe DICIEMBRE o 12, por ejemplo."
        )

    return anio, mes


def preguntar_si_regenerar(
    anio: int,
    mes: int
) -> bool:

    respuesta = input(
        f"Ya existe el rol de {mes_nombre(mes)} {anio}. "
        "¿Deseas regenerarlo y reemplazar el guardado? (S/N) [N]: "
    ).strip().upper()

    return respuesta == "S"


def exportar_mes_guardado(datos: dict) -> Path:
    return crear_excel(datos)


# ============================================================
# 17. RESUMEN FINAL
# ============================================================

def mostrar_final(
    datos: dict,
    ruta: Path,
    guardado: bool = False
) -> None:

    print("\n" + "=" * 64)

    if guardado:
        print("ROL RECUPERADO DE LA MEMORIA PERSISTENTE")
    else:
        print("ROL GENERADO Y GUARDADO CORRECTAMENTE")

    print(f"{datos['nombre_mes']} {datos['anio']}")
    print("=" * 64)

    print("Archivo Excel:", ruta.name)
    print("Estado persistente:", ARCHIVO_ESTADO.name)
    print("Ciclo: DIA -> NOCHE -> DESCANSO -> DESCANSO")

    print("\nTitulares Counter guardados:")

    for grupo, personas in datos["titulares_counter"].items():
        print(f"  {grupo}: {', '.join(personas)}")

    print("\nAsignaciones fijas del mes:")

    fijas = datos["asignaciones_fijas"]

    print("  Serpost:", " / ".join(fijas["serpost"]))
    print("  Carga Aérea:", " / ".join(fijas["carga_aerea"]))

    for grupo, personas in fijas["maritima"].items():
        print(f"  Marítima {grupo}:", " / ".join(personas))

    print("  Patrulla/Veterinaria: CIUDAD (lunes a viernes)")

    print("\nResumen de servicios:")

    for persona in GUIAS:
        print(
            f"  {persona:12s}: "
            f"{datos['conteos'].get(persona, 0)}"
        )

    print("=" * 64)
    print("La próxima vez que ejecutes el programa, podrá continuar")
    print("desde el último mes que quedó guardado.")


# ============================================================
# 18. PROGRAMA PRINCIPAL
# ============================================================

def main():

    try:

        estado = cargar_estado()

        anio, mes = preguntar_mes_anio()

        existente = mes_guardado(
            estado, anio, mes
        )

        if existente and not preguntar_si_regenerar(anio, mes):

            ruta = exportar_mes_guardado(existente)

            mostrar_final(
                existente,
                ruta,
                guardado=True
            )

            return

        validar_continuidad_mes(
            estado, anio, mes
        )

        semilla = random.SystemRandom().randint(
            1, 2_000_000_000
        )

        estado_trabajo = json.loads(
            json.dumps(estado, ensure_ascii=False)
        )

        estado_trabajo = reconstruir_estado_antes_de_generar(
            estado_trabajo,
            anio,
            mes
        )

        validar_continuidad_mes(
            estado_trabajo,
            anio,
            mes
        )

        datos = construir_mes(
            anio,
            mes,
            estado_trabajo,
            semilla
        )

        errores = validar_mes(
            datos,
            estado_trabajo
        )

        if errores:

            print("\nNo se guardó el mes porque fallaron validaciones:")

            for error in errores:
                print(" -", error)

            raise RuntimeError(
                "El rol no superó las validaciones."
            )

        guardar_mes_en_estado(
            estado_trabajo,
            datos
        )

        actualizar_base_al_ultimo_dia(
            estado_trabajo,
            datos
        )

        guardar_estado(estado_trabajo)

        ruta = crear_excel(datos)

        mostrar_final(
            datos,
            ruta,
            guardado=False
        )

    except KeyboardInterrupt:

        print("\nProceso cancelado por el usuario.")

    except Exception as exc:

        print("\nERROR:", exc)

        print(
            "No borres ESTADO_ROL_K9.json. "
            "Revisa el mensaje y vuelve a ejecutar el programa."
        )

        raise


if __name__ == "__main__":
    main()
