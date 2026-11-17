import json
import threading
import traceback
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, TOP, VERTICAL, X, Y, filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import pandas as pd

from comparador import MESES_NOMBRE, comparar
from comparador_101 import comparar_f101
from exportar import generar_reporte, generar_reporte_formularios
from mapeo_config import cargar_mapeo, cargar_mapeo_default, guardar_mapeo
from parser_ict import leer_mapeo_declaracion_ict
from parser_libro_mayor import agrupar_por_cuenta_periodo, leer_libro_mayor
from parser_pdf import casilleros_a_dataframe, extraer_formulario, validar_formulario


APP_TITLE = "Auditor SRI - Conciliacion 103/104 vs Libro Mayor"
COLOR_AZUL = "#003087"
COLOR_ROJO = "#C8102E"
COLOR_FONDO = "#F4F7FB"
COLOR_PANEL = "#FFFFFF"
COLOR_BORDE = "#D8DEE9"
VISIBLE_COLUMNS = [
    "PERIODO",
    "PERIODO_PAGO",
    "FORMULARIO",
    "NUMERO_SERIAL",
    "FECHA_RECAUDACION",
    "CAS",
    "CONCEPTO",
    "CUENTA_CONTABLE",
    "TP",
    "BASE_IMPONIBLE",
    "VALOR_RETENIDO",
    "VALOR_SRI",
    "VALOR_LIBRO",
    "DIFERENCIA",
    "ESTADO",
    "COMENTARIO",
]


def formato_moneda(valor: float) -> str:
    return f"$ {float(valor):,.2f}"


def cuentas_mapeo_en_libro(mapeos: list[dict], df_agrupado: pd.DataFrame) -> tuple[list[str], list[str]]:
    if df_agrupado.empty:
        return [], []
    cuentas_libro = set(df_agrupado["CTA_CONTABLE"].astype(str).str.strip())
    if "CTA_CODIGO" in df_agrupado.columns:
        cuentas_libro |= set(df_agrupado["CTA_CODIGO"].astype(str).str.strip())

    encontradas = []
    faltantes = []
    for mapeo in mapeos:
        cuenta = str(mapeo.get("cuenta_contable", "")).strip()
        if not cuenta:
            continue
        if cuenta in cuentas_libro:
            encontradas.append(cuenta)
        else:
            faltantes.append(cuenta)
    return sorted(set(encontradas)), sorted(set(faltantes))


class AuditorDesktopApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1360x820")
        self.minsize(1120, 680)
        self.configure(background=COLOR_FONDO)

        self.pdf_paths: list[str] = []
        self.excel_path: str | None = None
        self.modo: str | None = None
        self.formularios: list[dict] = []
        self.df_libro_mayor = pd.DataFrame()
        self.df_libro_mayor_agrupado = pd.DataFrame()
        self.df_ict = pd.DataFrame()
        self.df_comparacion = pd.DataFrame()
        self.df_filtrado = pd.DataFrame()
        self.errores_pdf: list[str] = []
        self.advertencias_periodo: list[str] = []
        self.mapeos = cargar_mapeo()
        self.comentarios: dict[tuple[str, str, str, str], str] = {}

        self._crear_estilos()
        self._crear_layout()

    def _crear_estilos(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(".", font=("Helvetica", 10), background=COLOR_FONDO)
        style.configure("TFrame", background=COLOR_FONDO)
        style.configure("Panel.TFrame", background=COLOR_PANEL, relief="solid", borderwidth=1)
        style.configure("Header.TFrame", background=COLOR_PANEL)
        style.configure("Title.TLabel", font=("Helvetica", 19, "bold"), foreground=COLOR_AZUL, background=COLOR_PANEL)
        style.configure("Subtitle.TLabel", font=("Helvetica", 10), foreground="#4B5563", background=COLOR_PANEL)
        style.configure("Section.TLabel", font=("Helvetica", 11, "bold"), foreground=COLOR_AZUL, background=COLOR_PANEL)
        style.configure("MetricTitle.TLabel", font=("Helvetica", 9), foreground="#4B5563", background=COLOR_PANEL)
        style.configure("MetricValue.TLabel", font=("Helvetica", 15, "bold"), foreground=COLOR_AZUL, background=COLOR_PANEL)
        style.configure("Status.TLabel", foreground="#374151", background=COLOR_PANEL)
        style.configure("Primary.TButton", font=("Helvetica", 10, "bold"), padding=(10, 7))
        style.configure("Export.TButton", font=("Helvetica", 10, "bold"), padding=(12, 8))
        style.configure("Treeview", rowheight=28, bordercolor=COLOR_BORDE, fieldbackground="white", background="white")
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), background=COLOR_AZUL, foreground="white")
        style.map("Treeview", background=[("selected", "#CFE1FF")], foreground=[("selected", "#111827")])

    def _crear_layout(self) -> None:
        header = ttk.Frame(self, padding=(18, 14, 18, 12), style="Header.TFrame")
        header.pack(side=TOP, fill=X)
        title_box = ttk.Frame(header, style="Header.TFrame")
        title_box.pack(side=LEFT, fill=X, expand=True)
        self.title_label = ttk.Label(title_box, text="Auditor SRI", style="Title.TLabel")
        self.title_label.pack(anchor="w")
        self.subtitle_label = ttk.Label(
            title_box,
            text="Conciliaciones tributarias",
            style="Subtitle.TLabel",
        )
        self.subtitle_label.pack(anchor="w", pady=(2, 0))
        self.home_button = ttk.Button(header, text="Inicio", command=self._mostrar_inicio)
        self.home_button.pack(side=RIGHT, padx=(8, 0))
        self.export_button_top = ttk.Button(
            header,
            text="Exportar reporte Excel",
            style="Export.TButton",
            command=self.exportar_reporte,
            state=tk.DISABLED,
        )
        tk.Frame(self, height=3, background=COLOR_ROJO).pack(fill=X)

        self.home_frame = ttk.Frame(self, padding=28)
        self._crear_inicio(self.home_frame)
        self.main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self._mostrar_inicio()

    def _crear_inicio(self, parent: ttk.Frame) -> None:
        inner = ttk.Frame(parent)
        inner.pack(expand=True)
        ttk.Label(inner, text="Seleccione el tipo de análisis", style="Title.TLabel").pack(pady=(0, 8))
        ttk.Label(
            inner,
            text="Elija el flujo de conciliación que desea ejecutar.",
            style="Subtitle.TLabel",
        ).pack(pady=(0, 24))

        options = ttk.Frame(inner)
        options.pack()
        card_103 = ttk.Frame(options, padding=20, style="Panel.TFrame")
        card_103.pack(side=LEFT, padx=10, ipadx=20, ipady=12)
        ttk.Label(card_103, text="Formularios 103 y 104", style="Section.TLabel").pack()
        ttk.Label(
            card_103,
            text="Comparar declaraciones mensuales\ncontra el Libro Mayor.",
            style="Subtitle.TLabel",
            justify="center",
        ).pack(pady=12)
        ttk.Button(
            card_103,
            text="Abrir análisis 103/104",
            style="Primary.TButton",
            command=lambda: self._seleccionar_modo("103_104"),
        ).pack(fill=X)

        card_101 = ttk.Frame(options, padding=20, style="Panel.TFrame")
        card_101.pack(side=LEFT, padx=10, ipadx=20, ipady=12)
        ttk.Label(card_101, text="Formulario 101", style="Section.TLabel").pack()
        ttk.Label(
            card_101,
            text="Comparar la declaración anual\ncontra el Anexo ICT A1.",
            style="Subtitle.TLabel",
            justify="center",
        ).pack(pady=12)
        ttk.Button(
            card_101,
            text="Abrir análisis 101",
            style="Primary.TButton",
            command=lambda: self._seleccionar_modo("101"),
        ).pack(fill=X)

        card_solo = ttk.Frame(options, padding=20, style="Panel.TFrame")
        card_solo.pack(side=LEFT, padx=10, ipadx=20, ipady=12)
        ttk.Label(card_solo, text="Análisis independiente", style="Section.TLabel").pack()
        ttk.Label(
            card_solo,
            text="Revisar formularios 103/104\nsin cargar Libro Mayor.",
            style="Subtitle.TLabel",
            justify="center",
        ).pack(pady=12)
        ttk.Button(
            card_solo,
            text="Analizar solo formularios",
            style="Primary.TButton",
            command=lambda: self._seleccionar_modo("solo_103_104"),
        ).pack(fill=X)

    def _mostrar_inicio(self) -> None:
        self.main_pane.pack_forget()
        self.export_button_top.pack_forget()
        self.home_button.pack_forget()
        self.home_frame.pack(fill=BOTH, expand=True)
        self.title_label.config(text="Auditor SRI")
        self.subtitle_label.config(text="Conciliaciones tributarias")

    def _seleccionar_modo(self, modo: str) -> None:
        self.modo = modo
        self.pdf_paths = []
        self.excel_path = None
        self.formularios = []
        self.df_comparacion = pd.DataFrame()
        self.df_filtrado = pd.DataFrame()
        self.errores_pdf = []
        self.advertencias_periodo = []
        self.home_frame.pack_forget()
        for child in self.main_pane.winfo_children():
            child.destroy()

        sidebar_host = ttk.Frame(self.main_pane, width=330, style="Panel.TFrame")
        content = ttk.Frame(self.main_pane, padding=10)
        self.main_pane.add(sidebar_host, weight=0)
        self.main_pane.add(content, weight=1)
        sidebar = self._crear_panel_desplazable(sidebar_host)
        self._crear_sidebar(sidebar)
        self._crear_contenido(content)
        self.main_pane.pack(fill=BOTH, expand=True, padx=12, pady=8)
        self.home_button.pack(side=RIGHT, padx=(8, 0))
        self.export_button_top.pack(side=RIGHT)
        self._set_export_state(False)
        if modo == "101":
            self.title_label.config(text="Auditor SRI - Formulario 101")
            self.subtitle_label.config(text="Declaración anual vs Anexo ICT, hoja MAPEO DE LA DECLARACIÓN A1")
        elif modo == "solo_103_104":
            self.title_label.config(text="Auditor SRI - Análisis de formularios")
            self.subtitle_label.config(text="Revisión independiente de formularios 103 y 104")
        else:
            self.title_label.config(text="Auditor SRI - Formularios 103/104")
            self.subtitle_label.config(text="Declaraciones mensuales vs Libro Mayor")

    def _crear_panel_desplazable(self, parent: ttk.Frame) -> ttk.Frame:
        canvas = tk.Canvas(
            parent,
            width=310,
            highlightthickness=0,
            background=COLOR_PANEL,
        )
        scrollbar = ttk.Scrollbar(parent, orient=VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        contenido = ttk.Frame(canvas, padding=12, style="Panel.TFrame")
        ventana = canvas.create_window((0, 0), window=contenido, anchor="nw")

        def actualizar_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def ajustar_ancho(event):
            canvas.itemconfigure(ventana, width=event.width)

        def desplazar(event):
            if getattr(event, "num", None) == 4:
                pasos = -1
            elif getattr(event, "num", None) == 5:
                pasos = 1
            else:
                delta = getattr(event, "delta", 0)
                pasos = -int(delta / 120) if abs(delta) >= 120 else -int(delta)
            if pasos:
                canvas.yview_scroll(pasos, "units")

        def activar_rueda(_event):
            canvas.bind_all("<MouseWheel>", desplazar)
            canvas.bind_all("<Button-4>", desplazar)
            canvas.bind_all("<Button-5>", desplazar)

        def desactivar_rueda(_event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        contenido.bind("<Configure>", actualizar_region)
        canvas.bind("<Configure>", ajustar_ancho)
        canvas.bind("<Enter>", activar_rueda)
        canvas.bind("<Leave>", desactivar_rueda)
        contenido.bind("<Enter>", activar_rueda)
        contenido.bind("<Leave>", desactivar_rueda)
        return contenido

    def _crear_sidebar(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Carga de archivos", style="Section.TLabel").pack(anchor="w", pady=(0, 8))

        texto_pdf = "Seleccionar Formulario 101 PDF" if self.modo == "101" else "Seleccionar PDFs 103/104"
        ttk.Button(parent, text=texto_pdf, command=self.seleccionar_pdfs).pack(fill=X, pady=3)
        self.pdf_label = ttk.Label(parent, text="0 PDFs seleccionados", wraplength=280)
        self.pdf_label.pack(anchor="w", pady=(0, 8))

        if self.modo != "solo_103_104":
            texto_excel = "Seleccionar Anexo ICT Excel" if self.modo == "101" else "Seleccionar Libro Mayor Excel"
            ttk.Button(parent, text=texto_excel, command=self.seleccionar_excel).pack(fill=X, pady=3)
            sin_excel = "Sin Anexo ICT" if self.modo == "101" else "Sin Libro Mayor"
            self.excel_label = ttk.Label(parent, text=sin_excel, wraplength=280)
            self.excel_label.pack(anchor="w", pady=(0, 12))

        ttk.Button(parent, text="Procesar archivos", style="Primary.TButton", command=self.procesar_en_hilo).pack(fill=X, pady=4)
        self.status_label = ttk.Label(parent, text="Listo.", wraplength=280, style="Status.TLabel")
        self.status_label.pack(anchor="w", pady=(8, 14))

        if self.modo == "103_104":
            ttk.Separator(parent).pack(fill=X, pady=10)
            ttk.Label(parent, text="Mapeo", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
            ttk.Button(parent, text="Editar mapeos", command=self.abrir_editor_mapeos).pack(fill=X, pady=3)
            ttk.Button(parent, text="Restaurar mapeo default", command=self.restaurar_mapeo).pack(fill=X, pady=3)
            ttk.Button(parent, text="Importar mapeo JSON", command=self.importar_mapeo).pack(fill=X, pady=3)
            ttk.Button(parent, text="Exportar mapeo JSON", command=self.exportar_mapeo).pack(fill=X, pady=3)

        ttk.Separator(parent).pack(fill=X, pady=10)
        ttk.Label(parent, text="Filtros", style="Section.TLabel").pack(anchor="w", pady=(0, 8))
        if self.modo == "solo_103_104":
            self.filtro_concepto = tk.StringVar(value="")
            self.filtro_estado = tk.StringVar(value="Todos")
        else:
            self.filtro_concepto = self._campo_filtro(parent, "Concepto contiene")
        estados = ["Todos", "OK", "DIFERENCIA", "SOLO_ANEXO"] if self.modo == "101" else [
            "Todos",
            "OK",
            "DIFERENCIA",
            "SOLO_SRI",
            "SOLO_LIBRO",
        ]
        formularios = ["Todos", "F101"] if self.modo == "101" else ["Todos", "F103", "F104"]
        tipos = ["Todos", "IR"] if self.modo == "101" else ["Todos", "IVA", "RF"]
        if self.modo != "solo_103_104":
            self.filtro_estado = self._combo(parent, "Estado", estados)
        self.filtro_formulario = self._combo(parent, "Formulario", formularios)
        self.filtro_tp = self._combo(parent, "TP", tipos)
        ttk.Button(parent, text="Aplicar filtros", command=self.aplicar_filtros).pack(fill=X, pady=(8, 3))
        ttk.Button(parent, text="Limpiar filtros", command=self.limpiar_filtros).pack(fill=X, pady=3)

        ttk.Separator(parent).pack(fill=X, pady=10)
        self.export_button_sidebar = ttk.Button(
            parent,
            text="Exportar reporte Excel",
            style="Export.TButton",
            command=self.exportar_reporte,
            state=tk.DISABLED,
        )
        self.export_button_sidebar.pack(fill=X, pady=3)
        self.casilleros_button = ttk.Button(
            parent,
            text="Ver todos los casilleros extraídos",
            command=self.abrir_casilleros_extraidos,
            state=tk.DISABLED,
        )
        self.casilleros_button.pack(fill=X, pady=3)

    def _campo_filtro(self, parent: ttk.Frame, label: str) -> tk.StringVar:
        var = tk.StringVar()
        ttk.Label(parent, text=label).pack(anchor="w")
        ttk.Entry(parent, textvariable=var).pack(fill=X, pady=(0, 6))
        return var

    def _combo(self, parent: ttk.Frame, label: str, values: list[str]) -> tk.StringVar:
        var = tk.StringVar(value=values[0])
        ttk.Label(parent, text=label).pack(anchor="w")
        combo = ttk.Combobox(parent, textvariable=var, values=values, state="readonly")
        combo.pack(fill=X, pady=(0, 6))
        return var

    def _crear_contenido(self, parent: ttk.Frame) -> None:
        if self.modo == "101":
            self.visible_columns = [
                "PERIODO",
                "FORMULARIO",
                "NUMERO_SERIAL",
                "FECHA_RECAUDACION",
                "CAS",
                "CONCEPTO",
                "VALOR_SRI",
                "VALOR_LIBRO",
                "DIFERENCIA",
                "ESTADO",
                "COMENTARIO",
            ]
        elif self.modo == "solo_103_104":
            self.visible_columns = [
                "ARCHIVO", "PERIODO", "FORMULARIO", "NUMERO_SERIAL", "FECHA_RECAUDACION",
                "BASE_IMPONIBLE", "VALOR_RETENIDO", "COMENTARIO",
            ]
        else:
            self.visible_columns = VISIBLE_COLUMNS
        metrics = ttk.Frame(parent)
        metrics.pack(fill=X, pady=(0, 10))
        titulos = (
            ("Formularios analizados", "Formularios 103", "Formularios 104", "Total retenido")
            if self.modo == "solo_103_104"
            else ("Total comparados", "Casilleros OK", "Con diferencia", "Diferencia total")
        )
        self.metric_total = self._crear_tarjeta_metrica(metrics, titulos[0], "0")
        self.metric_ok = self._crear_tarjeta_metrica(metrics, titulos[1], "0")
        self.metric_diff = self._crear_tarjeta_metrica(metrics, titulos[2], "0")
        self.metric_sum = self._crear_tarjeta_metrica(metrics, titulos[3], "$ 0.00")

        table_frame = ttk.Frame(parent)
        table_frame.pack(fill=BOTH, expand=True)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_frame, columns=self.visible_columns, show="headings", selectmode="browse")
        yscroll = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        widths = {
            "PERIODO": 110,
            "ARCHIVO": 210,
            "PERIODO_PAGO": 120,
            "FORMULARIO": 95,
            "NUMERO_SERIAL": 130,
            "FECHA_RECAUDACION": 125,
            "CAS": 70,
            "CONCEPTO": 260,
            "CUENTA_CONTABLE": 130,
            "TP": 50,
            "BASE_IMPONIBLE": 130,
            "VALOR_RETENIDO": 130,
            "VALOR_SRI": 120,
            "VALOR_LIBRO": 120,
            "DIFERENCIA": 120,
            "ESTADO": 105,
            "COMENTARIO": 220,
        }
        encabezados = {
            "NUMERO_SERIAL": "NÚMERO SERIAL",
            "FECHA_RECAUDACION": "FECHA RECAUDACIÓN",
            "PERIODO_PAGO": "PERIODO PAGO",
            "VALOR_SRI": "VALOR FORMULARIO" if self.modo == "101" else "VALOR SRI",
            "VALOR_LIBRO": "VALOR ANEXO ICT" if self.modo == "101" else "VALOR LIBRO",
            "CUENTA_CONTABLE": "CUENTA CONTABLE",
            "CAS": "CASILLERO" if self.modo in {"101", "solo_103_104"} else "CASILLERO COMPARADO",
            "BASE_IMPONIBLE": "BASE IMPONIBLE",
            "VALOR_RETENIDO": "VALOR RETENIDO",
            "PAGINA": "PÁGINA",
        }
        for col in self.visible_columns:
            self.tree.heading(col, text=encabezados.get(col, col), command=lambda c=col: self.ordenar_tabla(c))
            self.tree.column(col, width=widths.get(col, 120), minwidth=60, anchor="w")

        self.tree.tag_configure("OK", background="#DFF2E1")
        self.tree.tag_configure("DIFERENCIA", background="#F9D6D5")
        self.tree.tag_configure("SOLO", background="#FFF0B8")
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.mostrar_detalle)

        detail = ttk.LabelFrame(parent, text="Detalle de fila seleccionada", padding=8)
        detail.pack(fill=X, pady=(10, 0))
        detail_text_frame = ttk.Frame(detail)
        detail_text_frame.pack(fill=X)
        self.detail_text = tk.Text(detail_text_frame, height=8, wrap="word")
        detail_scroll = ttk.Scrollbar(detail_text_frame, orient=VERTICAL, command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_scroll.set)
        self.detail_text.pack(side=LEFT, fill=X, expand=True)
        detail_scroll.pack(side=RIGHT, fill=Y)

        comment_frame = ttk.Frame(detail)
        comment_frame.pack(fill=X, pady=(8, 0))
        ttk.Label(comment_frame, text="Comentario:").pack(side=LEFT)
        self.comment_var = tk.StringVar()
        ttk.Entry(comment_frame, textvariable=self.comment_var).pack(side=LEFT, fill=X, expand=True, padx=8)
        ttk.Button(comment_frame, text="Guardar comentario", command=self.guardar_comentario).pack(side=LEFT)

    def _crear_tarjeta_metrica(self, parent: ttk.Frame, titulo: str, valor: str) -> ttk.Label:
        card = ttk.Frame(parent, padding=(12, 9), style="Panel.TFrame")
        card.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        ttk.Label(card, text=titulo, style="MetricTitle.TLabel").pack(anchor="w")
        value_label = ttk.Label(card, text=valor, style="MetricValue.TLabel")
        value_label.pack(anchor="w", pady=(3, 0))
        return value_label

    def seleccionar_pdfs(self) -> None:
        if self.modo == "101":
            path = filedialog.askopenfilename(
                title="Seleccionar Formulario 101",
                filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            )
            paths = [path] if path else []
        else:
            paths = filedialog.askopenfilenames(
                title="Seleccionar PDFs 103/104 del SRI",
                filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            )
        if paths:
            self.pdf_paths = list(paths)
            texto = Path(self.pdf_paths[0]).name if self.modo == "101" else f"{len(self.pdf_paths)} PDFs seleccionados"
            self.pdf_label.config(text=texto)

    def seleccionar_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar Anexo ICT" if self.modo == "101" else "Seleccionar Libro Mayor",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
        )
        if path:
            self.excel_path = path
            self.excel_label.config(text=Path(path).name)

    def procesar_en_hilo(self) -> None:
        if not self.pdf_paths:
            messagebox.showerror(APP_TITLE, "Selecciona al menos un PDF del SRI.")
            return
        if self.modo != "solo_103_104" and not self.excel_path:
            fuente = "Anexo ICT" if self.modo == "101" else "Libro Mayor"
            messagebox.showerror(APP_TITLE, f"Selecciona el {fuente} en Excel.")
            return
        self.status_label.config(text="Procesando archivos...")
        threading.Thread(target=self._procesar_archivos, daemon=True).start()

    def _procesar_archivos(self) -> None:
        try:
            if self.modo == "101":
                formulario = extraer_formulario(self.pdf_paths[0])
                if formulario.get("tipo") != "F101":
                    raise ValueError("El PDF seleccionado no corresponde a un formulario 101.")
                df_ict = leer_mapeo_declaracion_ict(self.excel_path)
                df_comparacion = comparar_f101(formulario, df_ict)
                self.formularios = [formulario]
                self.df_ict = df_ict
                self.df_libro_mayor = pd.DataFrame()
                self.df_libro_mayor_agrupado = pd.DataFrame()
                self.df_comparacion = df_comparacion
                self.df_filtrado = df_comparacion
                resumen = (
                    f"Procesado F101 {formulario['periodo']['label']}. "
                    f"Casilleros comparados: {len(df_comparacion):,}. "
                    f"Coincidencias: {(df_comparacion['ESTADO'] == 'OK').sum():,}."
                )
                self.after(0, lambda: self._actualizar_ui_post_proceso(resumen))
                return

            formularios = []
            alertas = []
            errores_pdf = []
            for pdf_path in self.pdf_paths:
                try:
                    formulario = extraer_formulario(pdf_path)
                    formularios.append(formulario)
                    alertas.extend(validar_formulario(formulario))
                except Exception as exc:
                    errores_pdf.append(f"{Path(pdf_path).name}: {exc}")

            if not formularios:
                detalle_errores = "\n".join(errores_pdf)
                raise ValueError(f"No se pudo procesar ningún PDF seleccionado.\n{detalle_errores}")

            if self.modo == "solo_103_104":
                formularios_no_admitidos = [f for f in formularios if f.get("tipo") not in {"F103", "F104"}]
                if formularios_no_admitidos:
                    raise ValueError("El análisis independiente solo admite formularios 103 y 104.")
                df_comparacion = self._crear_analisis_independiente(formularios)
                self.formularios = formularios
                self.df_libro_mayor = pd.DataFrame()
                self.df_libro_mayor_agrupado = pd.DataFrame()
                self.df_comparacion = df_comparacion
                self.df_filtrado = df_comparacion
                self.errores_pdf = errores_pdf
                self.advertencias_periodo = []
                resumen = self._resumen_proceso(alertas, errores_pdf)
                self.after(0, lambda: self._actualizar_ui_post_proceso(resumen, alertas))
                return

            df_libro = leer_libro_mayor(self.excel_path)
            df_agrupado = agrupar_por_cuenta_periodo(df_libro)
            self._asegurar_mapeo_compatible(df_agrupado)
            df_comparacion = comparar(formularios, df_agrupado, self.mapeos)

            periodos_pdf = {
                (int(f["periodo"]["mes"]), int(f["periodo"]["año"]))
                for f in formularios
                if f.get("periodo", {}).get("mes") and f.get("periodo", {}).get("año")
            }
            periodos_libro = {
                (int(mes), int(anio))
                for mes, anio in zip(df_agrupado["MES"], df_agrupado["AÑO"])
            }
            periodos_comunes = periodos_pdf & periodos_libro
            advertencias_periodo = []
            if not periodos_comunes:
                anios_pdf = sorted({anio for _, anio in periodos_pdf})
                anios_libro = sorted({anio for _, anio in periodos_libro})
                advertencias_periodo.append(
                    "No existen periodos coincidentes entre formularios "
                    f"({', '.join(map(str, anios_pdf))}) y Libro Mayor "
                    f"({', '.join(map(str, anios_libro))})."
                )

            self.formularios = formularios
            self.df_libro_mayor = df_libro
            self.df_libro_mayor_agrupado = df_agrupado
            self.df_comparacion = df_comparacion
            self.df_filtrado = df_comparacion
            self.errores_pdf = errores_pdf
            self.advertencias_periodo = advertencias_periodo

            resumen = self._resumen_proceso(alertas, errores_pdf)
            self.after(0, lambda: self._actualizar_ui_post_proceso(resumen, alertas))
        except Exception as exc:
            detalle = traceback.format_exc()
            self.after(0, lambda exc=exc, detalle=detalle: self._mostrar_error(exc, detalle))

    def _asegurar_mapeo_compatible(self, df_agrupado: pd.DataFrame) -> None:
        encontradas, _ = cuentas_mapeo_en_libro(self.mapeos, df_agrupado)
        if encontradas:
            return
        mapeo_disco = cargar_mapeo()
        encontradas_disco, _ = cuentas_mapeo_en_libro(mapeo_disco, df_agrupado)
        if encontradas_disco:
            self.mapeos = mapeo_disco

    def _resumen_proceso(self, alertas: list[str], errores_pdf: list[str]) -> str:
        f103 = sum(1 for f in self.formularios if f.get("tipo") == "F103")
        f104 = sum(1 for f in self.formularios if f.get("tipo") == "F104")
        total_casilleros = sum(len(f.get("casilleros", {})) for f in self.formularios)
        periodos = sorted({f.get("periodo", {}).get("label", "") for f in self.formularios})
        fuente = "" if self.modo == "solo_103_104" else f" Libro Mayor: {len(self.df_libro_mayor):,} movimientos."
        texto = (
            f"Procesado. F103={f103}, F104={f104}. "
            f"Periodos: {', '.join(p for p in periodos if p)}. "
            f"Casilleros extraídos: {total_casilleros:,}." + fuente
        )
        if f103 and not f104 and self.modo == "103_104":
            texto += " Casillero usado para conciliar el pago: 999."
        if alertas:
            texto += f" Alertas SRI: {len(alertas)}."
        if errores_pdf:
            texto += f" PDFs omitidos: {len(errores_pdf)}."
        return texto

    def _actualizar_ui_post_proceso(self, resumen: str, alertas: list[str] | None = None) -> None:
        self.status_label.config(text=resumen)
        self.aplicar_filtros()
        self._set_export_state(True)
        self._set_casilleros_state(True)
        messagebox.showinfo(APP_TITLE, resumen)
        if alertas:
            messagebox.showwarning(APP_TITLE, "Validaciones internas del SRI:\n\n" + "\n".join(alertas))
        if self.errores_pdf:
            messagebox.showwarning(
                APP_TITLE,
                "PDFs omitidos:\n\n" + "\n".join(self.errores_pdf),
            )
        if self.advertencias_periodo:
            messagebox.showwarning(
                APP_TITLE,
                "Advertencia de periodos:\n\n" + "\n".join(self.advertencias_periodo),
            )

    def _mostrar_error(self, exc: Exception, detalle: str) -> None:
        self.status_label.config(text="Error al procesar.")
        self._set_export_state(False)
        self._set_casilleros_state(False)
        messagebox.showerror(APP_TITLE, f"{exc}\n\nDetalle tecnico:\n{detalle[-2000:]}")

    def aplicar_filtros(self) -> None:
        df = self.df_comparacion.copy()
        if df.empty:
            self.df_filtrado = df
            self._render_tabla()
            return

        concepto = self.filtro_concepto.get().strip()
        estado = self.filtro_estado.get()
        formulario = self.filtro_formulario.get()
        tp = self.filtro_tp.get()
        if concepto:
            df = df[df["CONCEPTO"].str.contains(concepto, case=False, na=False)]
        if estado != "Todos" and "ESTADO" in df.columns:
            df = df[df["ESTADO"].eq(estado)]
        if formulario != "Todos":
            df = df[df["FORMULARIO"].eq(formulario)]
        if tp != "Todos":
            df = df[df["TP"].eq(tp)]

        self.df_filtrado = df
        self._render_tabla()

    def limpiar_filtros(self) -> None:
        self.filtro_concepto.set("")
        self.filtro_estado.set("Todos")
        self.filtro_formulario.set("Todos")
        self.filtro_tp.set("Todos")
        self.aplicar_filtros()

    def _render_tabla(self) -> None:
        self.tree.delete(*self.tree.get_children())
        df = self.df_filtrado
        for idx, fila in df.reset_index(drop=False).iterrows():
            estado = str(fila.get("ESTADO", ""))
            tag = "SOLO" if estado.startswith("SOLO_") else estado
            valores = {
                "PERIODO": fila.get("PERIODO", ""),
                "ARCHIVO": fila.get("ARCHIVO", ""),
                "PERIODO_PAGO": fila.get("PERIODO_PAGO", ""),
                "FORMULARIO": fila.get("FORMULARIO", ""),
                "NUMERO_SERIAL": fila.get("NUMERO_SERIAL", ""),
                "FECHA_RECAUDACION": fila.get("FECHA_RECAUDACION", ""),
                "CAS": fila.get("CAS", ""),
                "CONCEPTO": fila.get("CONCEPTO", ""),
                "CUENTA_CONTABLE": fila.get("CUENTA_CONTABLE", ""),
                "TP": fila.get("TP", ""),
                "BASE_IMPONIBLE": self._formatear_moneda_opcional(fila.get("BASE_IMPONIBLE")),
                "VALOR_RETENIDO": self._formatear_moneda_opcional(fila.get("VALOR_RETENIDO")),
                "VALOR_SRI": formato_moneda(fila.get("VALOR_SRI", 0.0)),
                "VALOR_LIBRO": formato_moneda(fila.get("VALOR_LIBRO", 0.0)),
                "DIFERENCIA": formato_moneda(fila.get("DIFERENCIA", 0.0)),
                "ESTADO": estado,
                "COMENTARIO": fila.get("COMENTARIO", ""),
                "PAGINA": fila.get("PAGINA", ""),
            }
            values = [valores.get(col, "") for col in self.visible_columns]
            self.tree.insert("", END, iid=str(idx), values=values, tags=(tag,))
        self._actualizar_metricas()

    def _actualizar_metricas(self) -> None:
        df = self.df_filtrado
        total = len(df)
        if self.modo == "solo_103_104":
            f103 = int(df["FORMULARIO"].eq("F103").sum()) if not df.empty else 0
            f104 = int(df["FORMULARIO"].eq("F104").sum()) if not df.empty else 0
            retenido = float(df["VALOR_RETENIDO"].fillna(0).sum()) if not df.empty else 0.0
            self.metric_total.config(text=f"{total:,}")
            self.metric_ok.config(text=f"{f103:,}")
            self.metric_diff.config(text=f"{f104:,}")
            self.metric_sum.config(text=formato_moneda(retenido))
            return
        ok = int(df["ESTADO"].isin(["OK", "EXTRAIDO"]).sum()) if not df.empty else 0
        diff = int((df["ESTADO"] == "DIFERENCIA").sum()) if not df.empty else 0
        suma = float(df["DIFERENCIA"].sum()) if not df.empty else 0.0
        self.metric_total.config(text=f"{total:,}")
        self.metric_ok.config(text=f"{ok:,}")
        self.metric_diff.config(text=f"{diff:,}")
        self.metric_sum.config(text=formato_moneda(suma))

    @staticmethod
    def _formatear_moneda_opcional(valor) -> str:
        if valor is None or pd.isna(valor):
            return ""
        return formato_moneda(valor)

    @staticmethod
    def _crear_analisis_independiente(formularios: list[dict]) -> pd.DataFrame:
        filas = []
        for formulario in formularios:
            periodo = formulario.get("periodo", {})
            periodo_texto = str(periodo.get("label", ""))
            tipo = str(formulario.get("tipo", ""))
            casilleros = formulario.get("casilleros", {})
            casillero_base = "349" if tipo == "F103" else "409"
            casillero_retenido = "399" if tipo == "F103" else "799"
            filas.append({
                "MES": periodo.get("mes"),
                "AÑO": periodo.get("año"),
                "PERIODO_LABEL": periodo_texto,
                "PERIODO": periodo_texto,
                "ARCHIVO": formulario.get("archivo", ""),
                "FORMULARIO": tipo,
                "NUMERO_SERIAL": formulario.get("numero_serial", ""),
                "FECHA_RECAUDACION": formulario.get("fecha_recaudacion", ""),
                "CASILLERO": casillero_base,
                "CAS": casillero_base,
                "CASILLERO_BASE": casillero_base,
                "BASE_IMPONIBLE": float(casilleros.get(casillero_base, 0.0) or 0.0),
                "CASILLERO_RETENIDO": casillero_retenido,
                "VALOR_RETENIDO": float(casilleros.get(casillero_retenido, 0.0) or 0.0),
                "CONCEPTO": "Resumen del formulario",
                "CUENTA_CONTABLE": "",
                "TP": "IVA" if tipo == "F104" else "RF",
                "COMENTARIO": "",
            })
        return pd.DataFrame(filas)

    def ordenar_tabla(self, columna: str) -> None:
        if self.df_filtrado.empty:
            return
        mapa = {
            "CAS": "CASILLERO",
            "VALOR_SRI": "VALOR_SRI",
            "VALOR_LIBRO": "VALOR_LIBRO",
        }
        col_df = mapa.get(columna, columna)
        if col_df not in self.df_filtrado.columns:
            return
        self.df_filtrado = self.df_filtrado.sort_values(col_df, kind="mergesort")
        self._render_tabla()

    def mostrar_detalle(self, _event=None) -> None:
        seleccion = self.tree.selection()
        if not seleccion or self.df_filtrado.empty:
            return
        pos = int(seleccion[0])
        fila = self.df_filtrado.reset_index(drop=True).iloc[pos]
        self.comment_var.set(str(fila.get("COMENTARIO", "")))

        formulario = next(
            (
                f
                for f in self.formularios
                if f.get("tipo") == fila["FORMULARIO"]
                and f.get("periodo", {}).get("mes") == int(fila["MES"])
                and f.get("periodo", {}).get("año") == int(fila["AÑO"])
            ),
            None,
        )
        evidencia = (formulario or {}).get("evidencias", {}).get(str(fila["CASILLERO"]), {})

        texto = [
            f"Periodo: {fila['PERIODO_LABEL']} | Formulario: {fila['FORMULARIO']} | Casillero: {fila['CASILLERO']}",
            f"Número serial: {fila.get('NUMERO_SERIAL', '')} | Fecha recaudación: {fila.get('FECHA_RECAUDACION', '')}",
            f"Concepto: {fila['CONCEPTO']}",
        ]
        if self.modo == "solo_103_104":
            texto.extend([
                f"Base imponible (casillero {fila['CASILLERO_BASE']}): {formato_moneda(fila['BASE_IMPONIBLE'])}",
                f"Valor retenido (casillero {fila['CASILLERO_RETENIDO']}): {formato_moneda(fila['VALOR_RETENIDO'])}",
            ])
        else:
            texto.extend([
                f"Periodo de pago en Libro Mayor: {fila.get('PERIODO_PAGO', fila['PERIODO_LABEL'])}",
                f"Valor formulario: {formato_moneda(fila['VALOR_SRI'])} | "
                f"Valor {'Anexo ICT' if self.modo == '101' else 'Libro Mayor'}: "
                f"{formato_moneda(fila['VALOR_LIBRO_MAYOR'])} | "
                f"Diferencia: {formato_moneda(fila['DIFERENCIA'])}",
            ])
        texto.append(f"PDF pagina: {evidencia.get('pagina', 'N/D')} | Texto: {evidencia.get('texto_original', 'N/D')}")
        if self.modo == "103_104":
            movimientos = self._movimientos_fila(fila)
            texto.extend(["", f"Movimientos del Libro Mayor encontrados: {len(movimientos):,}"])
            if not movimientos.empty:
                muestra = movimientos.head(12).copy()
                muestra["FECHA"] = muestra["FECHA"].dt.strftime("%d/%m/%Y")
                texto.append(muestra[["FECHA", "CONCEPTO", "DEBE", "HABER", "SALDO"]].to_string(index=False))

        self.detail_text.delete("1.0", END)
        self.detail_text.insert("1.0", "\n".join(texto))

    def _movimientos_fila(self, fila: pd.Series) -> pd.DataFrame:
        if self.df_libro_mayor.empty:
            return pd.DataFrame()
        cuenta = str(fila["CUENTA_CONTABLE"]).strip()
        mascara = self.df_libro_mayor["CTA_CONTABLE"].astype(str).str.strip().eq(cuenta)
        if "CTA_CODIGO" in self.df_libro_mayor.columns:
            mascara = mascara | self.df_libro_mayor["CTA_CODIGO"].astype(str).str.strip().eq(cuenta)
        return self.df_libro_mayor[
            mascara
            & self.df_libro_mayor["MES"].astype(int).eq(int(fila["MES"]))
            & self.df_libro_mayor["AÑO"].astype(int).eq(int(fila["AÑO"]))
        ].copy()

    def guardar_comentario(self) -> None:
        seleccion = self.tree.selection()
        if not seleccion or self.df_filtrado.empty:
            return
        pos = int(seleccion[0])
        fila = self.df_filtrado.reset_index(drop=True).iloc[pos]
        clave = (
            fila["PERIODO_LABEL"],
            fila["FORMULARIO"],
            fila["CASILLERO"],
            fila.get("CUENTA_CONTABLE", ""),
        )
        comentario = self.comment_var.get()
        self.comentarios[clave] = comentario

        mask = (
            self.df_comparacion["PERIODO_LABEL"].eq(fila["PERIODO_LABEL"])
            & self.df_comparacion["FORMULARIO"].eq(fila["FORMULARIO"])
            & self.df_comparacion["CASILLERO"].eq(fila["CASILLERO"])
            & self.df_comparacion["CUENTA_CONTABLE"].eq(fila.get("CUENTA_CONTABLE", ""))
        )
        self.df_comparacion.loc[mask, "COMENTARIO"] = comentario
        self.aplicar_filtros()

    def exportar_reporte(self) -> None:
        if self.df_comparacion.empty:
            messagebox.showwarning(APP_TITLE, "No hay datos procesados para exportar.")
            return
        nombre = "Analisis_Formularios_SRI.xlsx" if self.modo == "solo_103_104" else "Conciliacion_SRI.xlsx"
        if not self.df_comparacion.empty:
            anio = self.df_comparacion["AÑO"].min()
            mes = self.df_comparacion["MES"].min()
            prefijo = "Analisis_Formularios_SRI" if self.modo == "solo_103_104" else "Conciliacion_SRI"
            nombre = f"{prefijo}_{int(anio)}_{int(mes):02d}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Guardar reporte",
            defaultextension=".xlsx",
            initialfile=nombre,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            if self.modo == "101":
                contenido = generar_reporte(
                    self.df_comparacion,
                    self.df_ict,
                    nombre_hoja_fuente="Anexo ICT A1 Cargado",
                    df_casilleros=casilleros_a_dataframe(self.formularios),
                )
            elif self.modo == "solo_103_104":
                contenido = generar_reporte_formularios(
                    self.df_comparacion,
                    casilleros_a_dataframe(self.formularios),
                )
            else:
                contenido = generar_reporte(
                    self.df_comparacion,
                    self.df_libro_mayor,
                    df_casilleros=casilleros_a_dataframe(self.formularios),
                )
            Path(path).write_bytes(contenido)
            messagebox.showinfo(APP_TITLE, f"Reporte guardado en:\n{path}")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"No se pudo generar el reporte Excel:\n\n{exc}")

    def _set_export_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        if hasattr(self, "export_button_top"):
            self.export_button_top.config(state=state)
        if hasattr(self, "export_button_sidebar"):
            self.export_button_sidebar.config(state=state)

    def _set_casilleros_state(self, enabled: bool) -> None:
        if hasattr(self, "casilleros_button"):
            self.casilleros_button.config(state=tk.NORMAL if enabled else tk.DISABLED)

    def abrir_casilleros_extraidos(self) -> None:
        df = casilleros_a_dataframe(self.formularios)
        if df.empty:
            messagebox.showwarning(APP_TITLE, "No hay casilleros extraídos para mostrar.")
            return

        ventana = tk.Toplevel(self)
        ventana.title("Casilleros extraídos de los formularios")
        ventana.geometry("1280x680")

        header = ttk.Frame(ventana, padding=10)
        header.pack(fill=X)
        ttk.Label(
            header,
            text=f"Todos los casilleros detectados: {len(df):,}",
            style="Section.TLabel",
        ).pack(side=LEFT)
        ttk.Label(
            header,
            text="La conciliación de pago usa el casillero 999; esta vista muestra el formulario completo.",
        ).pack(side=RIGHT)

        columnas = ["ARCHIVO", "FORMULARIO", "PERIODO", "CASILLERO", "VALOR_EXTRAIDO", "PAGINA", "TEXTO_ORIGINAL"]
        frame = ttk.Frame(ventana, padding=(10, 0, 10, 10))
        frame.pack(fill=BOTH, expand=True)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columnas, show="headings")
        yscroll = ttk.Scrollbar(frame, orient=VERTICAL, command=tree.yview)
        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        anchos = {
            "ARCHIVO": 210,
            "FORMULARIO": 90,
            "PERIODO": 110,
            "CASILLERO": 90,
            "VALOR_EXTRAIDO": 120,
            "PAGINA": 65,
            "TEXTO_ORIGINAL": 520,
        }
        encabezados = {"VALOR_EXTRAIDO": "VALOR EXTRAÍDO"}
        for columna in columnas:
            tree.heading(columna, text=encabezados.get(columna, columna))
            tree.column(columna, width=anchos[columna], minwidth=60, anchor="w")

        for _, fila in df.iterrows():
            valores = [fila.get(columna, "") for columna in columnas]
            valores[columnas.index("VALOR_EXTRAIDO")] = formato_moneda(fila["VALOR_EXTRAIDO"])
            tree.insert("", END, values=valores)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

    def restaurar_mapeo(self) -> None:
        self.mapeos = cargar_mapeo_default()
        guardar_mapeo(self.mapeos)
        messagebox.showinfo(APP_TITLE, "Mapeo default recargado.")

    def importar_mapeo(self) -> None:
        path = filedialog.askopenfilename(title="Importar mapeo", filetypes=[("JSON", "*.json")])
        if not path:
            return
        datos = json.loads(Path(path).read_text(encoding="utf-8"))
        self.mapeos = datos.get("mapeos", [])
        guardar_mapeo(self.mapeos)
        messagebox.showinfo(APP_TITLE, "Mapeo importado.")

    def exportar_mapeo(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Exportar mapeo",
            defaultextension=".json",
            initialfile="mapeo_sri.json",
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        Path(path).write_text(json.dumps({"mapeos": self.mapeos}, ensure_ascii=False, indent=2), encoding="utf-8")
        messagebox.showinfo(APP_TITLE, f"Mapeo guardado en:\n{path}")

    def abrir_editor_mapeos(self) -> None:
        editor = tk.Toplevel(self)
        editor.title("Editar mapeos")
        editor.geometry("1180x420")

        columns = [
            "id",
            "cuenta_contable",
            "descripcion_cuenta",
            "formulario",
            "casillero",
            "descripcion_casillero",
            "columna_libro",
            "signo",
            "desfase_meses",
        ]
        tree = ttk.Treeview(editor, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=145, minwidth=80)
        tree.pack(fill=BOTH, expand=True, padx=8, pady=8)

        def cargar_tree():
            tree.delete(*tree.get_children())
            for mapeo in self.mapeos:
                tree.insert("", END, values=[mapeo.get(col, "") for col in columns])

        def editar_seleccion():
            sel = tree.selection()
            if not sel:
                return
            valores = tree.item(sel[0], "values")
            self._abrir_form_mapeo(columns, dict(zip(columns, valores)), cargar_tree)

        def nuevo_mapeo():
            siguiente = max([int(m.get("id", 0)) for m in self.mapeos] or [0]) + 1
            self._abrir_form_mapeo(columns, {"id": siguiente, "signo": 1, "columna_libro": "HABER"}, cargar_tree)

        def eliminar():
            sel = tree.selection()
            if not sel:
                return
            id_val = int(tree.item(sel[0], "values")[0])
            self.mapeos = [m for m in self.mapeos if int(m.get("id", 0)) != id_val]
            guardar_mapeo(self.mapeos)
            cargar_tree()

        buttons = ttk.Frame(editor)
        buttons.pack(fill=X, padx=8, pady=(0, 8))
        ttk.Button(buttons, text="Nuevo", command=nuevo_mapeo).pack(side=LEFT, padx=3)
        ttk.Button(buttons, text="Editar", command=editar_seleccion).pack(side=LEFT, padx=3)
        ttk.Button(buttons, text="Eliminar", command=eliminar).pack(side=LEFT, padx=3)
        ttk.Button(buttons, text="Cerrar", command=editor.destroy).pack(side=RIGHT, padx=3)
        cargar_tree()

    def _abrir_form_mapeo(self, columns: list[str], data: dict, on_save) -> None:
        form = tk.Toplevel(self)
        form.title("Mapeo")
        vars_by_col = {}
        for row, col in enumerate(columns):
            ttk.Label(form, text=col).grid(row=row, column=0, sticky="w", padx=8, pady=4)
            var = tk.StringVar(value=str(data.get(col, "")))
            vars_by_col[col] = var
            ttk.Entry(form, textvariable=var, width=70).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        form.columnconfigure(1, weight=1)

        def guardar():
            nuevo = {col: vars_by_col[col].get() for col in columns}
            nuevo["id"] = int(nuevo["id"])
            nuevo["signo"] = int(nuevo.get("signo") or 1)
            ids = [int(m.get("id", 0)) for m in self.mapeos]
            if nuevo["id"] in ids:
                self.mapeos = [nuevo if int(m.get("id", 0)) == nuevo["id"] else m for m in self.mapeos]
            else:
                self.mapeos.append(nuevo)
            guardar_mapeo(self.mapeos)
            on_save()
            form.destroy()

        ttk.Button(form, text="Guardar", command=guardar).grid(row=len(columns), column=1, sticky="e", padx=8, pady=10)


def main() -> None:
    app = AuditorDesktopApp()
    app.mainloop()


if __name__ == "__main__":
    main()
