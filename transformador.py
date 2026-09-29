"""Convierte reportes SNIES por persona en el resumen institucional de 8 columnas.

El análisis no escribe archivos. La generación crea otro libro y nunca sobrescribe
el original ni una salida previa.
"""

from __future__ import annotations

import csv
import os
import re
import tempfile
import unicodedata
import warnings
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterator

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


CABECERAS = (
    "Categoría", "Unidad regional", "Nivel", "Nivel académico",
    "Programa académico", "Cantidad", "Año", "Periodo",
)
ESPECIFICACIONES = (
    ("Inscritos", "inscrito_programa", "PROGRAMA", "MUNICIPIO", "PROGRAMA CONSECUTIVO"),
    ("Admitidos", "admitidos", "PROGRAMA", "MUNICIPIO", "PRO_CONSECUTIVO"),
    ("Primiparos", "estudiantes_primer_curso", "NOMBRE PROGRAMA", "MUNICIPIO", "PROGRAMA CONSECUTIVO"),
    ("Matriculados", "matriculados", "PROGRAMA", "MUNICIPIO", "PROGRAMA CONSECUTIVO"),
    ("Graduados", "graduados", "PROGRAMA", "MUNICIPIO PROGRAMA", "PROGRAMA CONSECUTIVO"),
)
COLUMNAS_DISTINTIVAS = {
    "Inscritos": ("ID TIPO DOCUMENTO",),
    "Admitidos": ("PUNTAJE SISBEN",),
    "Primiparos": ("ID TIPO VINCULACIÓN",),
    "Matriculados": ("CODIGO ESTUDIANTE", "VALOR_DERECHOS_MATRÍCULA"),
    "Graduados": ("FECHA GRADO",),
}
AZUL = "17365D"
AZUL_CLARO = "DCEAF7"
NARANJA = "FFF1D6"


def _texto(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _clave(value: object) -> str:
    value = unicodedata.normalize("NFKD", _texto(value).casefold())
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def _entero(value: object, campo: str) -> int:
    try:
        n = int(float(str(value)))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{campo} no es un entero válido: {_texto(value)!r}") from exc
    if str(value).strip() and float(str(value)) != n:
        raise ValueError(f"{campo} no es un entero válido: {_texto(value)!r}")
    return n


def _clasificacion(programa: str) -> tuple[str, str]:
    nombre = _clave(programa)
    for prefijo, detalle in (
        ("tecnologia", "Tecnología"),
        ("especializacion", "Especialización"),
        ("maestria", "Maestría"),
        ("doctorado", "Doctorado"),
    ):
        if nombre.startswith(prefijo + " ") or nombre == prefijo:
            return ("Pregrado" if prefijo == "tecnologia" else "Posgrado", detalle)
    return "Pregrado", "Profesional universitario"


def _firma(path: Path) -> tuple[str, int, int]:
    stat = path.stat()
    return str(path.resolve()), stat.st_size, stat.st_mtime_ns


def _archivos(carpeta: Path, seleccion: dict[str, Path] | None = None) -> tuple[Path, dict[str, Path]]:
    carpeta = Path(carpeta).resolve()
    base = carpeta / "ESTUDIANTES.xlsx"
    if not base.is_file():
        raise FileNotFoundError("La carpeta debe contener ESTUDIANTES.xlsx.")
    encontrados: dict[str, Path] = {}
    categorias = {especificacion[0] for especificacion in ESPECIFICACIONES}
    if seleccion is not None:
        if set(seleccion) != categorias:
            faltantes = sorted(categorias - set(seleccion))
            adicionales = sorted(set(seleccion) - categorias)
            raise ValueError(f"Seleccione los cinco reportes. Faltantes: {faltantes}; no esperados: {adicionales}.")
        for categoria, *_ in ESPECIFICACIONES:
            path = Path(seleccion[categoria]).resolve()
            if not path.is_file() or path.suffix.lower() not in {".csv", ".xlsx"}:
                raise ValueError(f"El archivo de {categoria} debe existir y ser CSV o XLSX: {path}")
            encontrados[categoria] = path
        if len(set(encontrados.values())) != len(encontrados):
            raise ValueError("Debe elegir un archivo diferente para cada reporte.")
    else:
        reportes = carpeta / "reportes"
        if not reportes.is_dir():
            raise FileNotFoundError("No existe la subcarpeta reportes.")
        for categoria, fragmento, *_ in ESPECIFICACIONES:
            coincidencias = [p for p in reportes.iterdir()
                             if p.is_file() and p.suffix.lower() in {".csv", ".xlsx"}
                             and fragmento in _clave(p.stem).replace(" ", "_")]
            if len(coincidencias) != 1:
                raise ValueError(
                    f"Se esperaba un reporte de {categoria} en reportes; "
                    f"se encontraron {len(coincidencias)}. Revise nombres o duplicados."
                )
            encontrados[categoria] = coincidencias[0]
    return base, encontrados


def _filas_archivo(path: Path) -> Iterator[list[object]]:
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            yield from csv.reader(handle, delimiter=";")
    else:
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Workbook contains no default style")
            book = load_workbook(path, read_only=True, data_only=True)
        try:
            hoja = book.active
            hoja.reset_dimensions()  # Estos reportes declaran A1 aunque contienen miles de filas.
            yield from hoja.iter_rows(values_only=True)
        finally:
            book.close()


def _filas_reporte(path: Path, requeridas: tuple[str, ...]) -> Iterator[dict[str, object]]:
    cabecera: dict[str, int] | None = None
    for numero, fila in enumerate(_filas_archivo(path), 1):
        if cabecera is None:
            intento = {_clave(v): i for i, v in enumerate(fila) if _texto(v)}
            if all(_clave(col) in intento for col in requeridas):
                cabecera = intento
            elif numero >= 10:
                raise ValueError(f"No se reconocieron las columnas de {path.name}.")
            continue
        if not any(_texto(v) for v in fila):
            continue
        yield {col: fila[i] if i < len(fila) else None for col, i in cabecera.items()}
    if cabecera is None:
        raise ValueError(f"No se encontró encabezado en {path.name}.")


def _valor(fila: dict[str, object], nombre: str) -> object:
    return fila.get(_clave(nombre))


def _region(nombre: str, canonicas: dict[str, str]) -> str:
    clave = _clave(nombre)
    if clave == "villa de san diego de ubate":
        clave = "ubate"
    if clave not in canonicas:
        raise ValueError(f"Municipio sin equivalencia en ESTUDIANTES.xlsx: {nombre!r}")
    return canonicas[clave]


def _estilo_resumen(hoja, anchos: dict[str, float]) -> None:
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.alignment = Alignment(wrap_text=True, vertical="center")
    hoja.row_dimensions[1].height = 30
    for col, ancho in anchos.items():
        hoja.column_dimensions[col].width = ancho


@dataclass
class Resultado:
    carpeta: Path
    base: Path
    reportes: dict[str, Path]
    firmas: tuple[tuple[str, int, int], ...]
    periodo: tuple[int, int]
    periodos_salida: tuple[str, ...]
    nombre_sugerido: str
    filas_nuevas: list[tuple[object, ...]]
    canon_programas: dict[str, str]
    regiones: dict[str, str]
    duplicados: list[tuple[tuple[object, ...], list[tuple[int, int]]]]
    atipicos: list[tuple[object, ...]]
    atipicos_historicos: list[tuple[object, ...]]
    posibles_variantes: list[tuple[str, str, float]]
    nuevos_programas: list[tuple[str, int]]
    resumen: str
    conteos_fuente: dict[str, int]
    indicadores: dict[str, int]


def analizar(carpeta: Path, seleccion: dict[str, Path] | None = None) -> Resultado:
    """Analiza las cinco fuentes y el libro histórico sin escribir en disco."""
    base, reportes = _archivos(Path(carpeta), seleccion)
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Workbook contains no default style")
        libro = load_workbook(base, read_only=True, data_only=True)
    try:
        hoja = libro["ESTUDIANTES"]
        filas_historicas = list(hoja.iter_rows(values_only=True))
    finally:
        libro.close()
    if tuple(_clave(v) for v in filas_historicas[0][:8]) != tuple(_clave(v) for v in CABECERAS):
        raise ValueError("Las ocho columnas de ESTUDIANTES.xlsx no coinciden con el formato esperado.")

    nombres = defaultdict(Counter)
    sedes = defaultdict(Counter)
    for fila in filas_historicas[1:]:
        nombres[_clave(fila[4])][_texto(fila[4])] += 1
        sede = _texto(fila[1])
        if _clave(sede) != "villa de san diego de ubate":
            sedes[_clave(sede)][sede] += 1
    canon = {clave: conteo.most_common(1)[0][0] for clave, conteo in nombres.items()}
    regiones = {clave: conteo.most_common(1)[0][0] for clave, conteo in sedes.items()}
    posibles_variantes = []
    claves_programa = sorted(nombres)
    for i, primera in enumerate(claves_programa):
        for segunda in claves_programa[:i]:
            similitud = SequenceMatcher(None, primera, segunda).ratio()
            if similitud >= 0.94:
                posibles_variantes.append((canon[primera], canon[segunda], round(similitud, 3)))

    hist = defaultdict(list)
    formula_caches = 0
    ajustes_clasificacion = 0
    for numero, fila in enumerate(filas_historicas[1:], 2):
        categoria, sede, nivel, detalle, programa, cantidad, anio, periodo = fila[:8]
        programa_can = canon[_clave(programa)]
        sede_can = _region(_texto(sede), regiones)
        clave = (_texto(categoria), sede_can, programa_can, _entero(anio, "Año"), _texto(periodo))
        hist[clave].append((numero, _entero(cantidad, "Cantidad")))
        esperado_nivel, esperado_detalle = _clasificacion(programa_can)
        if _clave(detalle) != _clave(esperado_detalle):
            ajustes_clasificacion += 1
        if nivel is None or detalle is None:
            formula_caches += 1
    if formula_caches:
        raise ValueError("El Excel original tiene clasificaciones sin valor calculado; revise las fórmulas.")
    duplicados = [(clave, datos) for clave, datos in hist.items() if len(datos) > 1]

    conteos_fuente: dict[str, int] = {}
    individuos: dict[str, set[tuple[str, str, int, int]]] = defaultdict(set)
    cruces_programa: dict[str, set[tuple[str, str, int, int]]] = defaultdict(set)
    grupos = Counter()
    etiquetas_nuevas = defaultdict(Counter)
    nombre_por_documento = defaultdict(set)
    campos_nombre_por_documento = defaultdict(set)
    tipo_por_documento = defaultdict(set)
    etiqueta_por_codigo = defaultdict(set)
    instituciones = set()
    periodos = set()
    indicadores = Counter()
    codigos_estudiante = Counter()

    for categoria, _, col_programa, col_municipio, col_codigo in ESPECIFICACIONES:
        path = reportes[categoria]
        requeridas = ("AÑO", "SEMESTRE", "NOMBRE IES", "NUMERO DOCUMENTO", col_codigo,
                      col_programa, col_municipio, *COLUMNAS_DISTINTIVAS[categoria])
        conteo_archivo = 0
        for fila in _filas_reporte(path, requeridas):
            anio = _entero(_valor(fila, "AÑO"), "Año")
            semestre = _entero(_valor(fila, "SEMESTRE"), "Semestre")
            if semestre not in (1, 2):
                raise ValueError(f"Semestre inesperado en {path.name}: {semestre}")
            periodos.add((anio, semestre))
            institucion = _clave(_valor(fila, "NOMBRE IES"))
            instituciones.add(institucion)
            documento = _texto(_valor(fila, "NUMERO DOCUMENTO"))
            codigo = _texto(_valor(fila, col_codigo))
            programa = _texto(_valor(fila, col_programa))
            municipio = _texto(_valor(fila, col_municipio))
            if not all((documento, codigo, programa, municipio)):
                raise ValueError(f"Documento, código, programa o municipio vacío en {path.name}.")
            sede = _region(municipio, regiones)
            clave_programa = _clave(programa)
            etiqueta_por_codigo[codigo].add(clave_programa)
            if clave_programa not in canon:
                etiquetas_nuevas[clave_programa][programa] += 1
            identidad = (documento, codigo, anio, semestre)
            if identidad in individuos[categoria]:
                indicadores["duplicados_fuente"] += 1
                continue
            individuos[categoria].add(identidad)
            cruces_programa[categoria].add((documento, clave_programa, anio, semestre))
            conteo_archivo += 1
            grupos[(categoria, sede, clave_programa, anio, semestre)] += 1
            partes = [_texto(_valor(fila, nombre)) for nombre in
                      ("PRIMER NOMBRE", "SEGUNDO NOMBRE", "PRIMER APELLIDO", "SEGUNDO APELLIDO")]
            if any(partes):
                nombre_por_documento[documento].add(_clave(" ".join(partes)))
                campos_nombre_por_documento[documento].add(tuple(_clave(parte) for parte in partes))
            tipo = _texto(_valor(fila, "ID TIPO DOCUMENTO" if categoria == "Inscritos" else "TIPO DOCUMENTO"))
            if tipo:
                tipo_por_documento[documento].add(_clave(tipo))
            if categoria == "Matriculados":
                codigo_estudiante = _texto(_valor(fila, "CODIGO ESTUDIANTE"))
                if codigo_estudiante:
                    codigos_estudiante[codigo_estudiante] += 1
                matricula = _valor(fila, "VALOR_DERECHOS_MATRÍCULA")
                if _texto(matricula):
                    importe = _entero(matricula, "Valor de matrícula")
                    indicadores["matricula_cero"] += importe == 0
                    indicadores["matricula_mayor_10m"] += importe > 10_000_000
            if categoria == "Admitidos" and not _texto(_valor(fila, "PUNTAJE SISBEN")):
                indicadores["sisben_vacio"] += 1
        conteos_fuente[categoria] = conteo_archivo

    if len(periodos) != 1:
        raise ValueError(f"Los reportes deben corresponder a un solo año y semestre; hallados {sorted(periodos)}.")
    if instituciones != {"universidad de cundinamarca"}:
        raise ValueError(f"Nombre de institución inesperado en reportes: {sorted(instituciones)}")
    codigos_ambiguos = sum(len(v) > 1 for v in etiqueta_por_codigo.values())
    if codigos_ambiguos:
        raise ValueError(f"Hay {codigos_ambiguos} códigos de programa con nombres diferentes; revise los reportes.")

    nuevos = {}
    for clave, variantes in etiquetas_nuevas.items():
        preferida = max(variantes, key=lambda x: (sum(c in "áéíóúÁÉÍÓÚñÑ" for c in x), variantes[x]))
        nuevos[clave] = preferida[:1].upper() + preferida[1:].lower()
    canon.update(nuevos)
    anio, semestre = next(iter(periodos))
    periodo = "IPA" if semestre == 1 else "IIPA"
    periodos_existentes = {_texto(fila[7]) for fila in filas_historicas[1:] if fila[6] == anio}
    periodos_salida = tuple(p for p in ("IPA", "IIPA") if p in periodos_existentes | {periodo})
    etiqueta_salida = "_".join(periodos_salida)
    nombre_sugerido = f"ESTUDIANTES_{anio}_{etiqueta_salida}.xlsx"
    filas_nuevas = []
    for (categoria, sede, clave_programa, anio_fila, semestre_fila), cantidad in sorted(grupos.items()):
        programa = canon[clave_programa]
        nivel, detalle = _clasificacion(programa)
        fila = (categoria, sede, nivel, detalle, programa, cantidad, anio_fila,
                "IPA" if semestre_fila == 1 else "IIPA")
        clave_hist = (categoria, sede, programa, anio_fila, fila[7])
        if clave_hist in hist:
            anteriores = {n for _, n in hist[clave_hist]}
            if anteriores != {cantidad}:
                raise ValueError(f"Ya existe un grupo distinto en ESTUDIANTES.xlsx: {clave_hist}.")
            indicadores["grupos_ya_presentes"] += 1
            continue
        filas_nuevas.append(fila)

    # Comparaciones heurísticas con el mismo semestre del año anterior.
    atipicos = []
    for fila in filas_nuevas:
        categoria, sede, _, _, programa, cantidad, anio_fila, per = fila
        previo = hist.get((categoria, sede, programa, anio_fila - 1, per), [])
        if len(previo) != 1 or previo[0][1] < 20:
            continue
        anterior = previo[0][1]
        razon = cantidad / anterior
        if (razon >= 2.5 and cantidad - anterior >= 30) or (razon <= 0.25 and anterior - cantidad >= 20):
            atipicos.append((categoria, sede, programa, anio_fila - 1, anterior, anio_fila, cantidad, round(razon, 2)))

    indicadores["documentos_nombre_distinto"] = sum(len(v) > 1 for v in nombre_por_documento.values())
    indicadores["documentos_campos_nombre_distintos"] = sum(len(v) > 1 for v in campos_nombre_por_documento.values())
    indicadores["documentos_tipo_distinto"] = sum(len(v) > 1 for v in tipo_por_documento.values())
    indicadores["codigos_estudiante_reutilizados"] = sum(n > 1 for n in codigos_estudiante.values())
    indicadores["clasificaciones_corregidas"] = ajustes_clasificacion
    indicadores["duplicados_historicos"] = len(duplicados)
    indicadores["conflictos_historicos"] = sum(len({n for _, n in datos}) > 1 for _, datos in duplicados)
    indicadores["filas_historicas"] = len(filas_historicas) - 1

    datos_nuevos = sum(f[5] for f in filas_nuevas)
    total_fuente = sum(conteos_fuente.values())
    if datos_nuevos != total_fuente and not indicadores["grupos_ya_presentes"]:
        # Si no había grupos ya cargados, la agregación debe conservar el total.
        raise AssertionError("La agregación no conserva el total de registros fuente.")
    indicadores["primer_curso_sin_admitidos"] = len(cruces_programa["Primiparos"] - cruces_programa["Admitidos"])
    indicadores["primer_curso_sin_matricula"] = len(cruces_programa["Primiparos"] - cruces_programa["Matriculados"])
    indicadores["admitidos_sin_inscripcion"] = len(cruces_programa["Admitidos"] - cruces_programa["Inscritos"])
    historico_por_grupo = defaultdict(dict)
    for (categoria, sede, programa, anio_hist, periodo_hist), valores in hist.items():
        if len(valores) == 1:
            historico_por_grupo[(sede, programa, anio_hist, periodo_hist)][categoria] = valores[0][1]
    atipicos_historicos = []
    for (sede, programa, anio_hist, periodo_hist), cantidades in historico_por_grupo.items():
        if not all(c in cantidades for c in ("Primiparos", "Inscritos", "Admitidos")):
            continue
        primero = cantidades["Primiparos"]
        inscritos = cantidades["Inscritos"]
        admitidos = cantidades["Admitidos"]
        if primero >= 100 and primero >= 2 * max(inscritos, admitidos):
            atipicos_historicos.append((sede, programa, anio_hist, periodo_hist,
                                         primero, inscritos, admitidos))
    lineas = [
        "Archivos seleccionados: " + "; ".join(f"{categoria}: {reportes[categoria].name}"
                                                for categoria, *_ in ESPECIFICACIONES) + ".",
        f"Periodo de los cinco reportes: {anio}-{semestre} ({periodo}).",
        f"Semestres de {anio} en el Excel resultante: {', '.join(periodos_salida)}.",
        f"Excel histórico: {indicadores['filas_historicas']:,} filas. Reportes: {total_fuente:,} registros.",
        "Conteos: " + ", ".join(f"{cat} {conteos_fuente[cat]:,}" for cat, *_ in ESPECIFICACIONES) + ".",
        f"Se agregarán {len(filas_nuevas):,} grupos con {datos_nuevos:,} registros.",
        f"Programas nuevos: {', '.join(f'{nombre} ({n})' for nombre, n in sorted(((canon[k], sum(v.values())) for k, v in etiquetas_nuevas.items())) ) or 'ninguno'}.",
        f"Nombres de programa muy parecidos sin fusionar: {len(posibles_variantes)} par(es); revisar en REVISION.",
        f"Histórico: {indicadores['duplicados_historicos']} claves repetidas, "
        f"{indicadores['conflictos_historicos']} con cantidades contradictorias; se conservarán ambas filas.",
        f"Clasificación: {ajustes_clasificacion} filas históricas de nivel académico se corregirán; "
        "las fórmulas de nivel se convertirán a valores en el archivo nuevo.",
        f"Cruces personales, sin incluir nombres en la salida: "
        f"{indicadores['documentos_nombre_distinto']} documentos con variantes de nombre completo, "
        f"{indicadores['documentos_campos_nombre_distintos']} con diferencias entre campos de nombre y "
        f"{indicadores['documentos_tipo_distinto']} con tipos diferentes.",
        f"Matriculados: {indicadores['codigos_estudiante_reutilizados']} códigos de estudiante repetidos; "
        f"{indicadores['matricula_cero']} matrículas en cero y "
        f"{indicadores['matricula_mayor_10m']} mayores a $10 millones. "
        f"Admitidos: {indicadores['sisben_vacio']} SISBEN vacíos.",
        f"Atípicos de conteo frente al mismo semestre anterior: {len(atipicos)} grupos; se detallarán en REVISION.",
        f"Histórico: {len(atipicos_historicos)} grupos con primíparos al menos al doble de inscritos y admitidos; "
        "la relación es orientativa porque las cohortes pueden diferir.",
        f"Cruces 2026: {indicadores['primer_curso_sin_admitidos']} de primer curso no figuran en admitidos; "
        f"{indicadores['primer_curso_sin_matricula']} no figuran en matriculados; "
        f"{indicadores['admitidos_sin_inscripcion']} admitidos no figuran en inscritos.",
        "El original no se modifica; la salida tendrá hojas ESTUDIANTES, CONTROL y REVISION.",
    ]
    if indicadores["duplicados_fuente"]:
        lineas.append(f"Se omitieron {indicadores['duplicados_fuente']} filas repetidas por documento y programa.")
    firmas = tuple(_firma(path) for path in (base, *reportes.values()))
    return Resultado(Path(carpeta).resolve(), base, reportes, firmas, (anio, semestre),
                     periodos_salida, nombre_sugerido, filas_nuevas,
                     canon, regiones, duplicados, atipicos, atipicos_historicos,
                     posibles_variantes,
                     [(canon[k], sum(v.values())) for k, v in etiquetas_nuevas.items()],
                     "\n".join(lineas), conteos_fuente, dict(indicadores))


def generar(resultado: Resultado, destino: Path) -> Path:
    """Crea y verifica un Excel nuevo; nunca reemplaza una salida previa."""
    destino = Path(destino).resolve()
    if destino.suffix.lower() != ".xlsx":
        raise ValueError("La salida debe ser un archivo .xlsx.")
    if destino.exists() or destino == resultado.base:
        raise FileExistsError(f"Ya existe el archivo de salida: {destino}")
    if not destino.parent.is_dir():
        raise FileNotFoundError(f"No existe la carpeta de salida: {destino.parent}")
    base, reportes = _archivos(resultado.carpeta, resultado.reportes)
    firmas_actuales = tuple(_firma(path) for path in (base, *reportes.values()))
    if firmas_actuales != resultado.firmas:
        raise RuntimeError("Los archivos cambiaron después del análisis. Pulse Analizar de nuevo.")

    libro = load_workbook(resultado.base, data_only=False)
    cache = load_workbook(resultado.base, data_only=True, read_only=True)
    try:
        hoja = libro["ESTUDIANTES"]
        hoja_cache = cache["ESTUDIANTES"]
        niveles_cache = [(fila[2], fila[3]) for fila in hoja_cache.iter_rows(min_row=2, values_only=True)]
        if len(niveles_cache) != hoja.max_row - 1:
            raise ValueError("El Excel original tiene un número inesperado de filas calculadas.")
        for numero in range(2, hoja.max_row + 1):
            programa = resultado.canon_programas[_clave(hoja.cell(numero, 5).value)]
            sede = _region(_texto(hoja.cell(numero, 2).value), resultado.regiones)
            nivel, detalle = _clasificacion(programa)
            hoja.cell(numero, 2).value = sede
            hoja.cell(numero, 3).value = nivel
            hoja.cell(numero, 4).value = detalle
            hoja.cell(numero, 5).value = programa
            # Se preservan cantidades y años del histórico exactamente.
            nivel_cache, detalle_cache = niveles_cache[numero - 2]
            if nivel_cache is None or detalle_cache is None:
                raise ValueError(f"Nivel sin valor calculado en fila {numero} del original.")
        origen_estilo = hoja.max_row
        for fila in resultado.filas_nuevas:
            numero = hoja.max_row + 1
            for columna, valor in enumerate(fila, 1):
                celda = hoja.cell(numero, columna, valor)
                celda._style = copy(hoja.cell(origen_estilo, columna)._style)
                celda.alignment = copy(hoja.cell(origen_estilo, columna).alignment)
        tabla = hoja.tables.get("Tabla1")
        if tabla is not None:
            tabla.ref = f"A1:H{hoja.max_row}"
            if tabla.autoFilter is not None:
                tabla.autoFilter.ref = tabla.ref

        control = libro.create_sheet("CONTROL")
        control.append(("Concepto", "Valor"))
        control.append(("Fuente histórica", resultado.base.name))
        control.append(("Año", resultado.periodo[0]))
        control.append(("Semestre", resultado.periodo[1]))
        control.append(("Semestres del año en la salida", ", ".join(resultado.periodos_salida)))
        control.append(("Filas históricas", resultado.indicadores["filas_historicas"]))
        control.append(("Grupos nuevos", len(resultado.filas_nuevas)))
        control.append(("Registros nuevos", sum(f[5] for f in resultado.filas_nuevas)))
        for categoria, cantidad in resultado.conteos_fuente.items():
            control.append((f"Registros fuente: {categoria}", cantidad))
        for clave, cantidad in resultado.indicadores.items():
            control.append((clave.replace("_", " ").capitalize(), cantidad))
        for nombre, cantidad in resultado.nuevos_programas:
            control.append(("Programa nuevo para revisión", f"{nombre} ({cantidad})"))
        control.append(("Pares de nombres de programa por revisar", len(resultado.posibles_variantes)))
        control.append(("Nota", "Las discrepancias de nombres/documentos solo se reportan como conteos; no se exportan datos personales."))
        _estilo_resumen(control, {"A": 39, "B": 100})

        revision = libro.create_sheet("REVISION")
        revision.append(("Tipo", "Categoría", "Unidad regional", "Programa académico", "Año previo",
                         "Cantidad previa", "Año nuevo", "Cantidad nueva", "Razón / filas", "Nota"))
        for clave, datos in resultado.duplicados:
            categoria, sede, programa, anio, periodo = clave
            tipo = "Duplicado con conflicto" if len({n for _, n in datos}) > 1 else "Duplicado idéntico"
            revision.append((tipo, categoria, sede, programa, anio, ", ".join(str(n) for _, n in datos),
                             None, None, ", ".join(str(r) for r, _ in datos), f"Periodo {periodo}; conservar y revisar fuente"))
        for categoria, sede, programa, anio_ant, anterior, anio_nuevo, nuevo, razon in resultado.atipicos:
            revision.append(("Cambio atípico", categoria, sede, programa, anio_ant, anterior,
                             anio_nuevo, nuevo, razon, "Comparación orientativa con el mismo semestre; verificar oferta y fuente"))
        for sede, programa, anio_hist, periodo_hist, primero, inscritos, admitidos in resultado.atipicos_historicos:
            revision.append(("Relación histórica atípica", "Primiparos", sede, programa, anio_hist,
                             f"Inscritos {inscritos}; Admitidos {admitidos}", anio_hist, primero,
                             periodo_hist, "Primíparos al menos duplican ambas categorías; verificar definiciones y fuente"))
        for nombre, cantidad in resultado.nuevos_programas:
            revision.append(("Programa nuevo", "Inscritos", None, nombre, None, None,
                             resultado.periodo[0], cantidad, None, "No figura en el histórico"))
        for primero, segundo, similitud in resultado.posibles_variantes:
            revision.append(("Posible variante de nombre", None, None, primero, None, None,
                             None, None, similitud, f"Comparar con: {segundo}. No se fusionaron."))
        _estilo_resumen(revision, {"A": 25, "B": 17, "C": 24, "D": 65, "E": 13,
                                   "F": 18, "G": 13, "H": 18, "I": 17, "J": 64})
        for fila in revision.iter_rows(min_row=2):
            if fila[0].value in {"Duplicado con conflicto", "Cambio atípico", "Relación histórica atípica", "Programa nuevo", "Posible variante de nombre"}:
                for celda in fila:
                    celda.fill = PatternFill("solid", fgColor=NARANJA)

        descriptor, temporal = tempfile.mkstemp(prefix=".estudiantes_", suffix=".xlsx", dir=destino.parent)
        os.close(descriptor)
        temporal = Path(temporal)
        try:
            libro.save(temporal)
            prueba = load_workbook(temporal, read_only=True, data_only=True)
            try:
                datos = prueba["ESTUDIANTES"]
                if datos.max_row != resultado.indicadores["filas_historicas"] + len(resultado.filas_nuevas) + 1:
                    raise AssertionError("La salida no tiene el número esperado de filas.")
                if any(fila[2] is None or fila[3] is None
                       for fila in datos.iter_rows(min_row=2, values_only=True)):
                    raise AssertionError("La salida contiene nivel o nivel académico vacío.")
            finally:
                prueba.close()
            os.rename(temporal, destino)  # En Windows falla si el destino ya existe.
        finally:
            temporal.unlink(missing_ok=True)
    finally:
        cache.close()
        libro.close()
    return destino
