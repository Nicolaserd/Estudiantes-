"""Interfaz de escritorio para revisar y transformar reportes estudiantiles."""

from __future__ import annotations

import importlib
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Any


ARCHIVO_BASE = "ESTUDIANTES.xlsx"
CARPETA_REPORTES = "reportes"
CATEGORIAS = (
    ("Inscritos", "Inscritos"),
    ("Admitidos", "Admitidos"),
    ("Primiparos", "Primer curso (Primíparos)"),
    ("Matriculados", "Matriculados"),
    ("Graduados", "Graduados"),
)


def _es_carpeta_de_trabajo(carpeta: Path) -> bool:
    return (carpeta / ARCHIVO_BASE).is_file()


def _carpeta_inicial() -> Path | None:
    """Busca junto al programa y en los dos niveles superiores de cada origen."""
    origenes = [Path.cwd(), Path(__file__).resolve().parent]
    if getattr(sys, "frozen", False):
        origenes.insert(1, Path(sys.executable).resolve().parent)

    visitadas: set[Path] = set()
    for origen in origenes:
        carpeta = origen.resolve()
        for _ in range(3):  # Origen, padre y abuelo.
            if carpeta not in visitadas:
                visitadas.add(carpeta)
                if _es_carpeta_de_trabajo(carpeta):
                    return carpeta
            if carpeta.parent == carpeta:
                break
            carpeta = carpeta.parent
    return None


def _siguiente_destino(carpeta: Path, periodo: tuple[int, int] | None = None) -> Path:
    """Propone un nombre libre y conserva intactos los Excel ya generados."""
    anio, semestre = periodo or (2026, 2)
    etiqueta = "IPA" if semestre == 1 else "IIPA"
    base = carpeta / f"ESTUDIANTES_{anio}_{etiqueta}.xlsx"
    if not base.exists():
        return base

    indice = 2
    while True:
        candidato = carpeta / f"{base.stem}_{indice}{base.suffix}"
        if not candidato.exists():
            return candidato
        indice += 1


class Aplicacion(ttk.Frame):
    def __init__(self, ventana: tk.Tk) -> None:
        super().__init__(ventana, padding=16)
        self.ventana = ventana
        self.carpeta: Path | None = _carpeta_inicial()
        self.reportes: dict[str, Path] = {}
        self.resultado: Any | None = None
        self.ocupada = False
        self.operacion: str | None = None
        self.pendientes: queue.Queue[tuple[str, bool, Any]] = queue.Queue()

        self.ruta_texto = tk.StringVar(
            value=str(self.carpeta) if self.carpeta else "Ninguna carpeta seleccionada"
        )
        self.rutas_reportes = {
            categoria: tk.StringVar(value="Archivo sin seleccionar")
            for categoria, _ in CATEGORIAS
        }
        self.botones_reportes: dict[str, ttk.Button] = {}
        self.salida_texto = tk.StringVar(value="")
        self.estado_texto = tk.StringVar(value="Seleccione los cinco reportes.")

        self._crear_interfaz()
        self._actualizar_salida()
        self._actualizar_botones()
        self.ventana.after(100, self._procesar_pendientes)
        self.ventana.protocol("WM_DELETE_WINDOW", self._cerrar)

    def _crear_interfaz(self) -> None:
        self.ventana.title("Reportes de estudiantes")
        self.ventana.geometry("1000x790")
        self.ventana.minsize(780, 650)
        self.pack(fill="both", expand=True)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        titulo = ttk.Label(
            self,
            text="Transformar reportes de estudiantes",
            font=("Segoe UI", 16, "bold"),
        )
        titulo.grid(row=0, column=0, sticky="w", pady=(0, 14))

        archivos = ttk.LabelFrame(self, text="Archivos de trabajo", padding=12)
        archivos.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        archivos.columnconfigure(1, weight=1)

        ttk.Label(
            archivos,
            text="Carpeta que contiene ESTUDIANTES.xlsx:",
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Entry(
            archivos, textvariable=self.ruta_texto, state="readonly"
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8), pady=(5, 10))
        self.boton_carpeta = ttk.Button(
            archivos, text="Elegir carpeta…", command=self._elegir_carpeta
        )
        self.boton_carpeta.grid(row=1, column=2, sticky="e", pady=(5, 10))

        ttk.Label(
            archivos,
            text="Seleccione manualmente un archivo por cada categoría:",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 5))
        for fila, (categoria, etiqueta) in enumerate(CATEGORIAS, start=3):
            ttk.Label(archivos, text=f"{etiqueta}:").grid(
                row=fila, column=0, sticky="w", padx=(0, 12), pady=3
            )
            ttk.Entry(
                archivos, textvariable=self.rutas_reportes[categoria], state="readonly"
            ).grid(row=fila, column=1, sticky="ew", padx=(0, 8), pady=3)
            boton = ttk.Button(
                archivos,
                text="Elegir archivo…",
                command=lambda nombre=categoria: self._elegir_reporte(nombre),
            )
            boton.grid(row=fila, column=2, sticky="e", pady=3)
            self.botones_reportes[categoria] = boton

        ttk.Label(archivos, text="Nuevo archivo propuesto:").grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(12, 0)
        )
        ttk.Label(archivos, textvariable=self.salida_texto).grid(
            row=9, column=0, columnspan=3, sticky="w", pady=(4, 0)
        )

        analisis = ttk.LabelFrame(self, text="Resultado del análisis", padding=8)
        analisis.grid(row=2, column=0, sticky="nsew", pady=(0, 12))
        analisis.columnconfigure(0, weight=1)
        analisis.rowconfigure(0, weight=1)
        self.resumen = scrolledtext.ScrolledText(
            analisis, wrap="word", font=("Consolas", 10), relief="flat"
        )
        self.resumen.grid(row=0, column=0, sticky="nsew")
        self._mostrar_resumen(
            "Seleccione la carpeta y los cinco reportes; después pulse «Analizar». "
            "El análisis muestra hallazgos antes de crear el nuevo Excel."
        )

        acciones = ttk.Frame(self)
        acciones.grid(row=3, column=0, sticky="ew")
        acciones.columnconfigure(0, weight=1)
        ttk.Label(acciones, textvariable=self.estado_texto).grid(
            row=0, column=0, sticky="w", padx=(0, 12)
        )
        self.boton_analizar = ttk.Button(
            acciones, text="Analizar", command=self._analizar
        )
        self.boton_analizar.grid(row=0, column=1, padx=(0, 8))
        self.boton_generar = ttk.Button(
            acciones, text="Generar Excel", command=self._generar
        )
        self.boton_generar.grid(row=0, column=2)

    def _mostrar_resumen(self, texto: str) -> None:
        self.resumen.configure(state="normal")
        self.resumen.delete("1.0", "end")
        self.resumen.insert("1.0", texto)
        self.resumen.configure(state="disabled")

    def _actualizar_salida(self) -> None:
        self.salida_texto.set(
            str(_siguiente_destino(self.carpeta, self.resultado.periodo))
            if self.carpeta and self.resultado is not None
            else "Se definirá al terminar el análisis."
        )

    def _actualizar_botones(self) -> None:
        self.boton_carpeta.configure(state="disabled" if self.ocupada else "normal")
        for boton in self.botones_reportes.values():
            boton.configure(
                state="normal" if self.carpeta is not None and not self.ocupada else "disabled"
            )
        self.boton_analizar.configure(
            state="normal"
            if self.carpeta is not None
            and all(categoria in self.reportes for categoria, _ in CATEGORIAS)
            and not self.ocupada
            else "disabled"
        )
        self.boton_generar.configure(
            state="normal"
            if self.carpeta and self.resultado is not None and not self.ocupada
            else "disabled"
        )

    def _elegir_carpeta(self) -> None:
        if self.ocupada:
            return
        ruta = filedialog.askdirectory(
            parent=self.ventana,
            title="Seleccione la carpeta de estudiantes",
            initialdir=str(self.carpeta or Path.cwd()),
            mustexist=True,
        )
        if not ruta:
            return

        carpeta = Path(ruta).resolve()
        if not _es_carpeta_de_trabajo(carpeta):
            messagebox.showerror(
                "Carpeta no válida",
                "La carpeta debe contener ESTUDIANTES.xlsx.",
                parent=self.ventana,
            )
            return

        if carpeta == self.carpeta:
            return
        self.carpeta = carpeta
        self.reportes.clear()
        for ruta in self.rutas_reportes.values():
            ruta.set("Archivo sin seleccionar")
        self.resultado = None
        self.ruta_texto.set(str(carpeta))
        self.estado_texto.set("Seleccione los cinco reportes.")
        self._mostrar_resumen("Seleccione cada reporte y después pulse «Analizar».")
        self._actualizar_salida()
        self._actualizar_botones()

    def _elegir_reporte(self, categoria: str) -> None:
        if self.ocupada or self.carpeta is None:
            return
        carpeta_reportes = self.carpeta / CARPETA_REPORTES
        ruta = filedialog.askopenfilename(
            parent=self.ventana,
            title=f"Seleccione el reporte de {categoria}",
            initialdir=str(carpeta_reportes if carpeta_reportes.is_dir() else self.carpeta),
            filetypes=(
                ("Reportes CSV o Excel", "*.csv *.xlsx"),
                ("Archivos CSV", "*.csv"),
                ("Archivos Excel", "*.xlsx"),
            ),
        )
        if not ruta:
            return

        archivo = Path(ruta).resolve()
        if not archivo.is_file() or archivo.suffix.lower() not in {".csv", ".xlsx"}:
            messagebox.showerror(
                "Archivo no válido",
                "Seleccione un archivo existente con extensión .csv o .xlsx.",
                parent=self.ventana,
            )
            return
        for otra_categoria, otro_archivo in self.reportes.items():
            if otra_categoria != categoria and otro_archivo == archivo:
                messagebox.showerror(
                    "Reporte repetido",
                    f"Ese archivo ya está seleccionado para {otra_categoria}.",
                    parent=self.ventana,
                )
                return
        if self.reportes.get(categoria) == archivo:
            return

        self.reportes[categoria] = archivo
        self.rutas_reportes[categoria].set(str(archivo))
        self.resultado = None
        faltantes = [etiqueta for clave, etiqueta in CATEGORIAS if clave not in self.reportes]
        self.estado_texto.set(
            "Faltan: " + ", ".join(faltantes) if faltantes else "Cinco reportes listos para analizar."
        )
        self._mostrar_resumen("La selección cambió. Pulse «Analizar» para revisar los cinco archivos.")
        self._actualizar_salida()
        self._actualizar_botones()

    def _iniciar(self, operacion: str, trabajo: Any) -> None:
        self.ocupada = True
        self.operacion = operacion
        self.estado_texto.set(
            "Analizando archivos…" if operacion == "analizar" else "Generando Excel…"
        )
        self._actualizar_botones()

        def ejecutar() -> None:
            try:
                valor = trabajo()
                self.pendientes.put((operacion, True, valor))
            except Exception as error:
                self.pendientes.put((operacion, False, error))

        threading.Thread(target=ejecutar, daemon=True).start()

    def _analizar(self) -> None:
        if self.ocupada or self.carpeta is None:
            return
        if not _es_carpeta_de_trabajo(self.carpeta):
            messagebox.showerror(
                "Faltan archivos",
                "No encuentro ESTUDIANTES.xlsx. Seleccione de nuevo la carpeta de trabajo.",
                parent=self.ventana,
            )
            return
        faltantes = [etiqueta for clave, etiqueta in CATEGORIAS if clave not in self.reportes]
        if faltantes:
            messagebox.showerror(
                "Faltan reportes",
                "Seleccione un archivo para: " + ", ".join(faltantes) + ".",
                parent=self.ventana,
            )
            return
        for categoria, archivo in self.reportes.items():
            if not archivo.is_file():
                messagebox.showerror(
                    "Falta un reporte",
                    f"No encuentro el archivo de {categoria}:\n{archivo}",
                    parent=self.ventana,
                )
                return

        self.resultado = None
        self._mostrar_resumen("Analizando, espere un momento…")
        self._actualizar_salida()
        carpeta = self.carpeta
        seleccion = dict(self.reportes)

        def trabajo() -> Any:
            modulo = importlib.import_module("transformador")
            return modulo.analizar(carpeta, seleccion)

        self._iniciar("analizar", trabajo)

    def _generar(self) -> None:
        if self.ocupada or self.carpeta is None or self.resultado is None:
            return

        carpeta = self.carpeta
        resultado = self.resultado
        destino = _siguiente_destino(carpeta, resultado.periodo)
        self.salida_texto.set(str(destino))

        def trabajo() -> Path:
            modulo = importlib.import_module("transformador")
            # Revisa otra vez por si apareció un archivo mientras se mostraba
            # la propuesta. El transformador debe evitar sobreescribir también.
            ruta = _siguiente_destino(carpeta, resultado.periodo)
            return Path(modulo.generar(resultado, ruta))

        self._iniciar("generar", trabajo)

    def _procesar_pendientes(self) -> None:
        try:
            while True:
                operacion, correcto, valor = self.pendientes.get_nowait()
                self.ocupada = False
                self.operacion = None

                if not correcto:
                    self.estado_texto.set("No se pudo completar la operación.")
                    if operacion == "analizar":
                        self.resultado = None
                        self._mostrar_resumen("El análisis no se completó.")
                        self._actualizar_salida()
                    messagebox.showerror(
                        "Error",
                        f"{type(valor).__name__}: {valor}",
                        parent=self.ventana,
                    )
                elif operacion == "analizar":
                    self.resultado = valor
                    self._mostrar_resumen(str(valor.resumen))
                    self.estado_texto.set(
                        "Análisis listo. Revise los hallazgos y genere el Excel."
                    )
                    self._actualizar_salida()
                else:
                    self.estado_texto.set(f"Excel creado: {valor}")
                    self._actualizar_salida()
                    messagebox.showinfo(
                        "Excel creado",
                        f"Se creó el archivo:\n{valor}",
                        parent=self.ventana,
                    )

                self._actualizar_botones()
        except queue.Empty:
            pass

        self.ventana.after(100, self._procesar_pendientes)

    def _cerrar(self) -> None:
        if self.ocupada and self.operacion == "generar":
            messagebox.showinfo(
                "Generación en curso",
                "Espere a que termine la creación del Excel antes de cerrar.",
                parent=self.ventana,
            )
            return
        self.ventana.destroy()


def main() -> None:
    ventana = tk.Tk()
    Aplicacion(ventana)
    ventana.mainloop()


if __name__ == "__main__":
    main()
