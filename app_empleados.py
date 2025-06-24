import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os
import tkinter.font as tkFont

# Nombre del archivo Excel
EXCEL_FILE = 'hdc.xlsx'

class App:
    def __init__(self, master):
        self.master = master
        master.title("Control de Empleados")
        master.geometry("870x580") # Tamaño ajustado para una mejor distribución general
        master.resizable(False, False) # NO permitir redimensionar la ventana

        self.df_employees = pd.DataFrame()
        self.load_excel_data()

        # Almacena los empleados escaneados únicos y sus detalles
        # Formato: {employee_id: {'Nombre': str, 'Linea': str, 'Puesto': str, 'Antiguedad_Anos': float, 'Antiguedad_Dias': int, 'F_Servicio': datetime, 'POSITION': str}}
        self.scanned_employees_data = {}

        # Inicialización de totales programados
        self.programmed_total_employees = 0
        self.programmed_total_operadores = 0
        self.programmed_total_soportes = 0
        self.programmed_total_calidad = 0

        self.create_widgets()
        self.update_stats_labels() # Inicializar etiquetas de estadísticas
        self.txt_escaneo.focus_set() # Foco inicial en el textbox de escaneo

        # Mantener el foco en txt_escaneo cuando la ventana principal está activa
        master.bind("<FocusIn>", self._set_focus_on_scan_entry)

    def load_excel_data(self):
        """Carga los datos del archivo Excel 'hdc.xlsx'."""
        if os.path.exists(EXCEL_FILE):
            try:
                self.df_employees = pd.read_excel(EXCEL_FILE, dtype={'Empleado': str}, parse_dates=['F Servicio'])
            except Exception as e:
                messagebox.showerror("Error de Carga", f"No se pudo cargar el archivo Excel: {e}", parent=self.master)
                self.df_employees = pd.DataFrame(columns=['Empleado', 'Nombre', 'Localidad', 'Turno', 'F Servicio', 'Departamento', 'LINEA', 'Puesto', 'Categoria', 'POSITION', 'FUNCTION', 'Proceso', 'Departamento'])
        else:
            messagebox.showwarning("Archivo no encontrado", f"El archivo '{EXCEL_FILE}' no se encontró en el directorio actual. Por favor, asegúrese de que el archivo exista y tenga los encabezados correctos.", parent=self.master)
            self.df_employees = pd.DataFrame(columns=['Empleado', 'Nombre', 'Localidad', 'Turno', 'F Servicio', 'Departamento', 'LINEA', 'Puesto', 'Categoria', 'POSITION', 'FUNCTION', 'Proceso', 'Departamento'])

    def create_widgets(self):
        """Crea y organiza los widgets en la ventana principal."""

        # Mover la definición de vcmd aquí para que esté disponible
        vcmd = (self.master.register(self._validate_numeric_input), '%P')

        # Configurar estilos modernos
        style = ttk.Style()
        style.theme_use('clam') # Un tema neutro y limpio

        # --- CAMBIOS DE COLOR DE FONDO ---

        # 1. Color de fondo de la ventana principal (Root Window)
        # Este es el color base más externo. Si los frames no tienen fondo, se verá este.
        self.master.configure(bg="#F0F2F5") # Un gris claro, similar al del ejemplo anterior. Puedes cambiarlo a #FFFFFF para blanco.

        # 2. Estilos para Titles y Labels
        # Es crucial que estos estilos tengan un 'background' si quieres que los Labels tengan un color específico.
        # Si no lo tienen, usarán el color de fondo de su Frame padre.
        style.configure("Title.TLabel", font=("Arial", 18, "bold"), foreground="#333", anchor="center", background="#F0F2F5") # Asignar un fondo explícito
        style.configure("SubTitle.TLabel", font=("Arial", 13, "bold"), foreground="#555", background="#FFFFFF") # Un blanco para subtítulos
        style.configure("TLabel", font=("Arial", 11), background="#FFFFFF") # Un blanco para etiquetas generales

        # 3. Estilos para LabelFrame (las "cajas" que agrupan secciones)
        # El background aquí controla el fondo DENTRO del área del LabelFrame.
        # El background de TLabelframe.Label controla el fondo de la ETIQUETA del título del LabelFrame.
        style.configure("TLabelframe.Label", font=("Arial", 15, "bold"), foreground="#2C2C2C", background="#F0F2F5") # Fondo para el texto "Escaneo de Empleados", "Línea Activa" etc.
        style.configure("TLabelframe", background="#FFFFFF", relief="groove", borderwidth=1, lightcolor="#ddd", darkcolor="#bbb") # Fondo de la caja del LabelFrame (ej. el recuadro blanco de "Escaneo de Empleados")

        # Estilo para Entry (textbox)
        style.configure("TEntry", font=("Arial", 12), padding=7, relief="flat", borderwidth=1, fieldbackground="#ffffff", foreground="#333")
        style.map("TEntry", bordercolor=[('focus', '#007bff')])

        # Estilo para Combobox
        style.configure("TCombobox", font=("Arial", 11), padding=5, fieldbackground="#ffffff", foreground="#333333")
        style.map("TCombobox", fieldbackground=[('readonly', '#ffffff')], selectbackground=[('readonly', '#007bff')], selectforeground=[('readonly', 'white')])

        # Estilo para los botones principales
        style.configure("TButton", font=("Arial", 13, "bold"), padding=12, borderwidth=0, relief="flat",
                                 background="#007bff", foreground="white")
        style.map("TButton",
                  background=[('active', '#0056b3')], # Azul más oscuro al pasar el ratón
                  foreground=[('active', 'white')])

        # Estilo para Treeview (en la ventana de vista de empleados)
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#e0e0e0", foreground="#333", padding=5)
        style.configure("Treeview", font=("Arial", 10), rowheight=28, background="#ffffff", fieldbackground="#ffffff")
        style.map("Treeview", background=[('selected', '#007bff')])

        # Definir fuentes específicas para mayor claridad
        category_name_font = ("Arial", 12, "bold") # Para "Operadores", "Soportes", "Calidad", "Total" programados
        category_value_font = ("Arial", 50, "bold") # Para los números grandes (ej. "50")
        
        # Fuentes para las estadísticas detalladas
        stats_font_label = ("Arial", 12, "bold")
        stats_font_value = ("Arial", 15, "bold")
        stats_font_difference = ("Arial", 16, "bold")

        # --- Frame principal que contiene todo el diseño ---
        # Si quieres que el fondo del main_frame sea diferente del root, defínelo aquí
        main_frame = ttk.Frame(self.master, padding="5", style="TFrame") # Asegúrate de que use un estilo si quieres cambiar su fondo directamente
        main_frame.pack(fill=tk.BOTH, expand=True)
        # Si main_frame no tiene estilo, su fondo será el de self.master (el root)

        # Puedes definir un estilo para TFrame si quieres que main_frame tenga un color diferente.
        # style.configure("TFrame", background="#F8F8F8") # Por ejemplo, un gris muy claro.

        # Configuración de columnas y filas para el nuevo diseño simétrico
        main_frame.grid_columnconfigure(0, weight=1, uniform="layout_group") 
        main_frame.grid_columnconfigure(1, weight=0, uniform="layout_group") 

        main_frame.grid_rowconfigure(1, weight=0) 
        main_frame.grid_rowconfigure(2, weight=2) 


        # --- HEADER: Título y botones de gestión de ventanas (Arriba de todo) ---
        header_frame = ttk.Frame(main_frame, padding=(0, 0, 0, 15)) # Padding solo en la parte inferior
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew")
        # El fondo de header_frame heredará el fondo de main_frame si no tiene un estilo propio
        # o puedes ponerle un fondo específico:
        # header_frame.configure(background="#E0E0E0") # Ejemplo

        ttk.Label(header_frame, text="Sistema de Control de Personal", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=10)

        # Contenedor para los botones de la derecha del header
        header_buttons_frame = ttk.Frame(header_frame)
        header_buttons_frame.grid(row=0, column=1, columnspan=2, sticky="e", padx=10)
        # El fondo de header_buttons_frame heredará el fondo de header_frame

        self.btn_programar = ttk.Button(header_buttons_frame, text="Programar Personal", command=self.open_programming_window, style="TButton")
        self.btn_programar.pack(side=tk.LEFT, padx=10)

        self.btn_ver_empleados = ttk.Button(header_buttons_frame, text="Ver Registros", command=self.open_employee_view_window, style="TButton")
        self.btn_ver_empleados.pack(side=tk.LEFT, padx=10)


        # --- Contenedor Izquierdo Principal (Escaneo y Línea) ---
        left_panel = ttk.Frame(main_frame, padding="2")
        left_panel.grid(row=1, column=0, padx=5, pady=5, sticky="nsew") 
        # Fondo de left_panel heredará de main_frame. Si quieres un fondo específico:
        # left_panel.configure(background="#F5F5F5") # Ejemplo

        # Sección de Escaneo de Empleados
        scan_frame = ttk.LabelFrame(left_panel, text="Escaneo de Empleados", padding="2") 
        scan_frame.grid(row=0, column=0, sticky="nsew", pady=2) 
        scan_frame.grid_columnconfigure(0, weight=1) 

        # Ajuste de Label y Entry para el escaneo
        # Estos Labels ahora usarán el background de SubTitle.TLabel si está definido, o el de scan_frame.
        ttk.Label(scan_frame, text="Número de Empleado:", style="SubTitle.TLabel").pack(pady=(5, 5), padx=10, anchor="w")
        self.txt_escaneo = ttk.Entry(scan_frame, width=25, font=("Arial", 14), validate="key", validatecommand=vcmd, style="TEntry")
        self.txt_escaneo.pack(pady=(0, 10), padx=10, fill=tk.X, expand=True) 
        self.txt_escaneo.bind("<Return>", self.process_scan)

        # Sección de Selección de Línea
        line_frame = ttk.LabelFrame(left_panel, text="Línea Activa", padding="1") 
        line_frame.grid(row=1, column=0, sticky="nsew", pady=5) 
        line_frame.grid_rowconfigure(0, weight=0) 
        line_frame.grid_rowconfigure(1, weight=1) 
        line_frame.grid_columnconfigure(0, weight=1)

        ttk.Label(line_frame, text="Seleccione la Línea de Trabajo:", style="SubTitle.TLabel").pack(pady=5, anchor="w", padx=10)
        
        # Combobox
        self.lines = ["F37", "F45", "F50", "F60", "F62", "F63", "F66", "F71", "F84", "F86", "T31", "T32", "T33", "T34"]
        self.cb_lines = ttk.Combobox(line_frame, values=self.lines, state="readonly", font=("Arial", 11), style="TCombobox")
        self.cb_lines.pack(pady=10, padx=10, fill=tk.X, expand=True) 
        self.cb_lines.bind("<<ComboboxSelected>>", self.update_stats_labels)

        if self.lines:
            self.cb_lines.set(self.lines[0]) 


        # --- Contenedor Derecho Principal (Estadísticas Programadas y Detalladas) ---
        right_panel = ttk.Frame(main_frame, padding="5")
        right_panel.grid(row=1, column=1, padx=1, pady=1, sticky="nsew") 
        # Fondo de right_panel heredará de main_frame. Si quieres un fondo específico:
        # right_panel.configure(background="#F5F5F5") # Ejemplo

        # Sección de Empleados Programados (Distribución Horizontal)
        programmed_display_frame = ttk.LabelFrame(right_panel, text="Personal Programado", padding="5")
        programmed_display_frame.grid(row=0, column=0, sticky="nsew", pady=10)
        
        # Frame interno para los valores programados (asegura la simetría horizontal)
        inner_prog_frame = ttk.Frame(programmed_display_frame, padding="5")
        inner_prog_frame.pack(fill=tk.BOTH, expand=True)
        # Fondo de inner_prog_frame heredará de programmed_display_frame. Si quieres un fondo específico:
        # inner_prog_frame.configure(background="#FFFFFF") # Ejemplo

        inner_prog_frame.grid_columnconfigure(0, weight=1)
        inner_prog_frame.grid_columnconfigure(1, weight=1)
        inner_prog_frame.grid_columnconfigure(2, weight=1)
        inner_prog_frame.grid_columnconfigure(3, weight=1)

        # Operadores Programados, Soportes, Calidad, Total
        # Estos Labels ahora usarán el background de TLabel si está definido, o el de inner_prog_frame.
        ttk.Label(inner_prog_frame, text="Operadores", font=category_name_font, background="#FFFFFF").grid(row=0, column=0, padx=5, pady=(5,0), sticky="ew") # Fondo específico para esta etiqueta
        self.lbl_programado_operadores = ttk.Label(inner_prog_frame, text="0", font=category_value_font, background="#FFFFFF") # Fondo específico
        self.lbl_programado_operadores.grid(row=1, column=0, padx=5, pady=(0,5), sticky="ew")

        ttk.Label(inner_prog_frame, text="Soportes", font=category_name_font, background="#FFFFFF").grid(row=0, column=1, padx=5, pady=(5,0), sticky="ew")
        self.lbl_programado_soportes = ttk.Label(inner_prog_frame, text="0", font=category_value_font, background="#FFFFFF")
        self.lbl_programado_soportes.grid(row=1, column=1, padx=5, pady=(0,5), sticky="ew")

        ttk.Label(inner_prog_frame, text="Calidad", font=category_name_font, background="#FFFFFF").grid(row=0, column=2, padx=5, pady=(5,0), sticky="ew")
        self.lbl_programado_calidad = ttk.Label(inner_prog_frame, text="0", font=category_value_font, background="#FFFFFF")
        self.lbl_programado_calidad.grid(row=1, column=2, padx=5, pady=(0,5), sticky="ew")

        ttk.Label(inner_prog_frame, text="Total", font=category_name_font, foreground="#007bff", background="#FFFFFF").grid(row=0, column=3, padx=5, pady=(5,0), sticky="ew")
        self.lbl_programado_total = ttk.Label(inner_prog_frame, text="0", font=category_value_font, foreground="#007bff", background="#FFFFFF")
        self.lbl_programado_total.grid(row=1, column=3, padx=5, pady=(0,5), sticky="ew")


        # Sección de Estadísticas Detalladas (Distribución Vertical)
        stats_frame = ttk.LabelFrame(right_panel, text="Registro de Personal", padding="15") 
        stats_frame.grid(row=1, column=0, sticky="nsew", pady=10)
        # Fondo de stats_frame heredará el de TLabelframe style, que ya tiene #FFFFFF

        stats_frame.grid_columnconfigure(0, weight=0) 
        stats_frame.grid_columnconfigure(1, weight=0) 

        label_pady = 2 
        label_padx = 10 

        # Operadores en Piso, Soportes, Calidad, etc.
        # Estos Labels usarán el background de TLabel o el de stats_frame.
        self.lbl_total_empleados_text = ttk.Label(stats_frame, text="Operadores en Piso:", font=stats_font_label, background="#FFFFFF")
        self.lbl_total_empleados_text.grid(row=0, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_total_empleados = ttk.Label(stats_frame, text="0", font=stats_font_value, foreground="#0056b3", background="#FFFFFF")
        self.lbl_total_empleados.grid(row=0, column=1, sticky="e", pady=label_pady, padx=label_padx) 

        self.lbl_total_mfgupo_text = ttk.Label(stats_frame, text="Soportes:", font=stats_font_label, background="#FFFFFF")
        self.lbl_total_mfgupo_text.grid(row=1, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_total_mfgupo = ttk.Label(stats_frame, text="0", font=stats_font_value, foreground="#0056b3", background="#FFFFFF")
        self.lbl_total_mfgupo.grid(row=1, column=1, sticky="e", pady=label_pady, padx=label_padx)

        self.lbl_total_qainsp_text = ttk.Label(stats_frame, text="Calidad:", font=stats_font_label, background="#FFFFFF")
        self.lbl_total_qainsp_text.grid(row=2, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_total_qainsp = ttk.Label(stats_frame, text="0", font=stats_font_value, foreground="#0056b3", background="#FFFFFF")
        self.lbl_total_qainsp.grid(row=2, column=1, sticky="e", pady=label_pady, padx=label_padx)

        # Separador para diferenciar categorías
        ttk.Separator(stats_frame, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=10, padx=label_padx)

        self.lbl_total_experiencia_text = ttk.Label(stats_frame, text="Operadores con Experiencia (>90 días):", font=stats_font_label, background="#FFFFFF")
        self.lbl_total_experiencia_text.grid(row=4, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_total_experiencia = ttk.Label(stats_frame, text="0", font=stats_font_value, background="#FFFFFF")
        self.lbl_total_experiencia.grid(row=4, column=1, sticky="e", pady=label_pady, padx=label_padx)

        self.lbl_total_sin_experiencia_text = ttk.Label(stats_frame, text="Operadores sin Experiencia (<=90 días):", font=stats_font_label, background="#FFFFFF")
        self.lbl_total_sin_experiencia_text.grid(row=5, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_total_sin_experiencia = ttk.Label(stats_frame, text="0", font=stats_font_value, background="#FFFFFF")
        self.lbl_total_sin_experiencia.grid(row=5, column=1, sticky="e", pady=label_pady, padx=label_padx)

        self.lbl_no_linea_seleccionada_text = ttk.Label(stats_frame, text="Operadores de Otras Líneas (Prestados):", font=stats_font_label, background="#FFFFFF")
        self.lbl_no_linea_seleccionada_text.grid(row=6, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_no_linea_seleccionada = ttk.Label(stats_frame, text="0", font=stats_font_value, background="#FFFFFF")
        self.lbl_no_linea_seleccionada.grid(row=6, column=1, sticky="e", pady=label_pady, padx=label_padx)
        
        # Separador para la diferencia
        ttk.Separator(stats_frame, orient="horizontal").grid(row=7, column=0, columnspan=2, sticky="ew", pady=10, padx=label_padx)

        self.lbl_diferencia_text = ttk.Label(stats_frame, text="Diferencia (Registrado - Programado):", font=stats_font_difference, foreground="#dc3545", background="#FFFFFF")
        self.lbl_diferencia_text.grid(row=8, column=0, sticky="w", pady=label_pady, padx=label_padx)
        self.lbl_diferencia = ttk.Label(stats_frame, text="0", font=stats_font_difference, foreground="#dc3545", background="#FFFFFF")
        self.lbl_diferencia.grid(row=8, column=1, sticky="e", pady=label_pady, padx=label_padx)


    def _validate_numeric_input(self, new_value):
        """Valida que la entrada del textbox de escaneo sea numérica y de 8 dígitos."""
        if new_value.isdigit() or new_value == "":
            if len(new_value) <= 8:
                return True
            else:
                return False
        return False

    def _set_focus_on_scan_entry(self, event=None):
        """Establece el foco en el campo de escaneo si la ventana principal es el foco."""
        if self.master.focus_get() is not self.txt_escaneo:
            self.txt_escaneo.focus_set()

    def process_scan(self, event=None):
        """Procesa el número de empleado escaneado."""
        employee_id = self.txt_escaneo.get().strip()
        self.txt_escaneo.delete(0, tk.END)

        if not employee_id:
            messagebox.showwarning("Entrada Vacía", "Por favor, escanee o ingrese un número de empleado.", parent=self.master)
            self.txt_escaneo.focus_set()
            return

        employee_id = str(employee_id)

        employee_info = self.df_employees[self.df_employees['Empleado'] == employee_id]

        if employee_info.empty:
            messagebox.showerror("Empleado No Encontrado", f"El empleado con ID '{employee_id}' no se encontró en la base de datos.", parent=self.master)
        else:
            if employee_id in self.scanned_employees_data:
                messagebox.showinfo("Duplicado", f"El empleado {employee_id} ya ha sido registrado.", parent=self.master)
            else:
                row = employee_info.iloc[0]
                nombre = row.get('Nombre', 'N/A')
                linea = row.get('LINEA', 'N/A')
                puesto = row.get('Puesto', 'N/A')
                position = row.get('POSITION', 'N/A')
                f_servicio = row.get('F Servicio', pd.NaT)

                antiguedad_anos, antiguedad_dias = self.calculate_antiguedad(f_servicio)

                self.scanned_employees_data[employee_id] = {
                    'Nombre': nombre,
                    'Linea': linea,
                    'Puesto': puesto,
                    'POSITION': position,
                    'Antiguedad_Anos': antiguedad_anos,
                    'Antiguedad_Dias': antiguedad_dias,
                    'F_Servicio': f_servicio
                }
                messagebox.showinfo("Registro Exitoso", f"Empleado {employee_id} - {nombre} registrado correctamente.", parent=self.master)

                self.master.event_generate("<<ScanUpdate>>")


        self.update_stats_labels()
        self.txt_escaneo.focus_set()

    def calculate_antiguedad(self, f_servicio):
        """Calcula la antigüedad en años y días desde la fecha de servicio."""
        if pd.isna(f_servicio) or not isinstance(f_servicio, (datetime, pd.Timestamp)):
            return 0.0, 0

        today = datetime.now()
        delta = today - f_servicio

        years = delta.days / 365.25
        days = delta.days

        return round(years, 1), delta.days

    def is_experienced(self, f_servicio):
        """Determina si un empleado tiene experiencia (más de 90 días)."""
        if pd.isna(f_servicio) or not isinstance(f_servicio, (datetime, pd.Timestamp)):
            return False

        today = datetime.now()
        delta = today - f_servicio
        return delta.days > 90

    def update_stats_labels(self, event=None):
        """Actualiza todas las etiquetas de estadísticas en la interfaz."""
        total_unique_scanned = len(self.scanned_employees_data)
        
        self.lbl_total_empleados.config(text=f"{total_unique_scanned}")

        total_mfgupo = sum(1 for emp_data in self.scanned_employees_data.values() if emp_data.get('POSITION', '').lower() == 'mfgupo')
        total_qainsp = sum(1 for emp_data in self.scanned_employees_data.values() if emp_data.get('POSITION', '').lower() == 'qainsp')

        self.lbl_total_mfgupo.config(text=f"{total_mfgupo}")
        self.lbl_total_qainsp.config(text=f"{total_qainsp}")

        total_experienced = 0
        total_inexperienced = 0
        for emp_data in self.scanned_employees_data.values():
            if self.is_experienced(emp_data['F_Servicio']):
                total_experienced += 1
            else:
                total_inexperienced += 1
        self.lbl_total_experiencia.config(text=f"{total_experienced}")
        self.lbl_total_sin_experiencia.config(text=f"{total_inexperienced}")

        # Obtener la línea seleccionada del Combobox
        selected_line = self.cb_lines.get()
        total_not_in_selected_line = 0
        if selected_line:
            total_not_in_selected_line = sum(1 for emp_data in self.scanned_employees_data.values() if emp_data.get('Linea', '').lower() != selected_line.lower())
        self.lbl_no_linea_seleccionada.config(text=f"{total_not_in_selected_line}")

        # Actualizar las etiquetas de programación en la sección horizontal
        self.lbl_programado_operadores.config(text=f"{self.programmed_total_operadores}")
        self.lbl_programado_soportes.config(text=f"{self.programmed_total_soportes}")
        self.lbl_programado_calidad.config(text=f"{self.programmed_total_calidad}")

        self.programmed_total_employees = (
            self.programmed_total_operadores +
            self.programmed_total_soportes +
            self.programmed_total_calidad
        )
        self.lbl_programado_total.config(text=f"{self.programmed_total_employees}")

        difference = total_unique_scanned - self.programmed_total_employees
        self.lbl_diferencia.config(text=f"{difference}")


    def open_programming_window(self):
        """Abre la ventana para programar empleados."""
        ProgrammingWindow(self.master, self)

    def open_employee_view_window(self):
        """Abre la ventana para ver los empleados registrados."""
        EmployeeViewWindow(self.master, self.scanned_employees_data, self)


class ProgrammingWindow(tk.Toplevel):
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.title("Programar Personal")
        self.geometry("380x250") # **Tamaño optimizado para un ajuste perfecto y visibilidad**
        self.resizable(False, False)
        self.app_instance = app_instance
        self.grab_set() # Hace que la ventana de programación sea modal
        self.transient(master) # Relaciona la ventana de programación con la principal
        self.focus_set() # Asegura que el foco esté en esta ventana al abrir

        # 4. Color de fondo de la ventana ProgrammingWindow (Root de esta ventana)
        self.configure(bg="#F0F2F5") # Fondo para toda la ventana de programación

        self.create_widgets()

        # Insertar los valores actuales al abrir la ventana
        self.entry_total_operadores.insert(0, str(self.app_instance.programmed_total_operadores))
        self.entry_total_soportes.insert(0, str(self.app_instance.programmed_total_soportes))
        self.entry_total_calidad.insert(0, str(self.app_instance.programmed_total_calidad))
        self.entry_total_operadores.focus_set() # Pone el foco en el primer campo al abrir

    def create_widgets(self):
        """Crea y organiza los widgets en la ventana de programación."""
        style = ttk.Style()
        # 5. Estilo para los Labels dentro de ProgrammingWindow
        style.configure("Prog.TLabel", font=("Arial", 12, "bold"), foreground="#333", background="#F0F2F5") # Fondo para los labels "Total Operadores:", etc.
        style.configure("Prog.TEntry", font=("Arial", 12), padding=8, relief="flat", borderwidth=1, fieldbackground="#ffffff", foreground="#333")
        style.map("Prog.TEntry", bordercolor=[('focus', '#007bff')])
        style.configure("Prog.TButton", font=("Arial", 13, "bold"), padding=12, borderwidth=0, relief="flat",
                                 background="#28a745", foreground="white") # Botón de guardar en verde
        style.map("Prog.TButton", background=[('active', '#218838')])

        # Frame principal con padding para el contenido
        frame = ttk.Frame(self, padding="1") # Aumentado padding interno
        frame.pack(fill=tk.BOTH, expand=True)
        # El fondo de 'frame' heredará el color de self (la ventana ProgrammingWindow) que ya definimos con bg="#F0F2F5"
        # Si quieres un color diferente para este frame, puedes hacer:
        # frame.configure(background="#E8EBEF") # Un gris un poco más oscuro, por ejemplo

        # Configurar grid para etiquetas y entradas para una alineación perfecta
        frame.grid_columnconfigure(0, weight=1) 
        frame.grid_columnconfigure(1, weight=1) 
        
        # Espaciado vertical entre filas
        row_pady = 2 

        # Operadores
        ttk.Label(frame, text="Total Operadores:", style="Prog.TLabel").grid(row=0, column=0, sticky="w", pady=row_pady, padx=5)
        self.entry_total_operadores = ttk.Entry(frame, width=15, style="Prog.TEntry")
        self.entry_total_operadores.grid(row=0, column=1, sticky="ew", pady=row_pady, padx=5)

        # Soportes
        ttk.Label(frame, text="Total Soportes:", style="Prog.TLabel").grid(row=1, column=0, sticky="w", pady=row_pady, padx=5)
        self.entry_total_soportes = ttk.Entry(frame, width=15, style="Prog.TEntry")
        self.entry_total_soportes.grid(row=1, column=1, sticky="ew", pady=row_pady, padx=5)

        # Calidad
        ttk.Label(frame, text="Total Calidad:", style="Prog.TLabel").grid(row=2, column=0, sticky="w", pady=row_pady, padx=5)
        self.entry_total_calidad = ttk.Entry(frame, width=15, style="Prog.TEntry")
        self.entry_total_calidad.grid(row=2, column=1, sticky="ew", pady=row_pady, padx=5)

        # Botón para registrar la programación
        ttk.Button(frame, text="Registrar Programación", command=self.save_programming, style="Prog.TButton").grid(row=3, column=0, columnspan=2, pady=3) 

    def save_programming(self):
        """Guarda los valores de programación y actualiza la ventana principal."""
        try:
            operadores_prog = int(self.entry_total_operadores.get())
            soportes_prog = int(self.entry_total_soportes.get())
            calidad_prog = int(self.entry_total_calidad.get())

            self.app_instance.programmed_total_operadores = operadores_prog
            self.app_instance.programmed_total_soportes = soportes_prog
            self.app_instance.programmed_total_calidad = calidad_prog

            self.app_instance.programmed_total_employees = (
                operadores_prog +
                soportes_prog +
                calidad_prog
            )

            self.app_instance.update_stats_labels()
            messagebox.showinfo("Programación Guardada", "La programación se ha guardado exitosamente.", parent=self)
            self.destroy()
        except ValueError:
            messagebox.showerror("Error de Entrada", "Por favor, ingrese solo números enteros para la programación.", parent=self)
        finally:
            self.app_instance.txt_escaneo.focus_set()

class EmployeeViewWindow(tk.Toplevel):
    def __init__(self, master, scanned_data, app_instance):
        super().__init__(master)
        self.title("Empleados Registrados")
        self.geometry("1100x650") # Ajustado tamaño
        self.scanned_employees_data = scanned_data
        self.app_instance = app_instance
        self.grab_set()
        self.transient(master)
        self.focus_set()

        # 6. Color de fondo de la ventana EmployeeViewWindow (Root de esta ventana)
        self.configure(bg="#F0F2F5") # Fondo para toda la ventana de vista de empleados

        self.create_widgets()
        self.update_tables()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.bind("<Configure>", self.on_resize)

        # Usar un callback lambda para pasar el evento directamente a update_tables_event
        self.app_instance.master.bind("<<ScanUpdate>>", lambda event: self.update_tables_event())


    def create_widgets(self):
        """Crea y organiza los widgets en la ventana de vista de empleados."""
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        # El fondo de main_frame heredará el de self (la ventana EmployeeViewWindow).

        search_frame = ttk.Frame(main_frame, padding="10") 
        search_frame.pack(fill=tk.X, pady=10)
        # El fondo de search_frame heredará el de main_frame.
        # Puedes añadir: search_frame.configure(background="#FFFFFF") si quieres un color diferente

        # 7. Fondo de la etiqueta "Buscar Empleado"
        ttk.Label(search_frame, text="Buscar Empleado (ID o Nombre):", font=("Arial", 12, "bold"), background="#F0F2F5").pack(side=tk.LEFT, padx=5)
        self.search_entry = ttk.Entry(search_frame, width=50, font=("Arial", 12)) 
        self.search_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.search_entry.bind("<KeyRelease>", self.filter_tables)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=15) 

        self.tabs = {}
        tab_names = ["Todos", "Op. con Experiencia", "Op. sin Experiencia", "Soportes", "Calidad", "Op. Prestados"]

        # 8. Colores de fondo de las pestañas (Notebook)
        style = ttk.Style() # Re-obtener el estilo para esta ventana si es necesario
        style.configure("TNotebook", background="#F0F2F5", borderwidth=0) # Fondo general del área del Notebook
        style.configure("TNotebook.Tab", background="#DCE3EE", foreground="#555", padding=[10, 5]) # Fondo y texto de pestañas no seleccionadas
        style.map("TNotebook.Tab", background=[("selected", "#FFFFFF")], foreground=[("selected", "#333")]) # Fondo y texto de pestaña seleccionada

        for name in tab_names:
            frame = ttk.Frame(self.notebook, padding="10")
            self.notebook.add(frame, text=name)
            # 9. Fondo de los frames dentro de cada pestaña del Notebook
            # Estos frames son los contenedores de las tablas. Su fondo se verá alrededor de la tabla si no la cubre por completo.
            frame.configure(background="#FFFFFF") # Un fondo blanco para el contenido de cada pestaña
            self.tabs[name] = self.create_employee_table(frame)

        self.selected_line_for_tab = self.app_instance.cb_lines.get() if self.app_instance.cb_lines.get() else None


    def create_employee_table(self, parent_frame):
        """Crea una tabla Treeview para mostrar datos de empleados."""
        # 10. Fondo del frame que contiene la tabla (si la tabla no ocupa todo el espacio)
        table_frame = ttk.Frame(parent_frame, background="#FFFFFF") # Fondo blanco para este frame
        table_frame.pack(fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(table_frame, columns=("ID Empleado", "Nombre", "Línea", "Puesto", "Antigüedad", "Experiencia"), show="headings")
        tree.heading("ID Empleado", text="ID Empleado", anchor=tk.W)
        tree.heading("Nombre", text="Nombre", anchor=tk.W)
        tree.heading("Línea", text="Línea", anchor=tk.W)
        tree.heading("Puesto", text="Puesto", anchor=tk.W)
        tree.heading("Antigüedad", text="Antigüedad (Años y Días)", anchor=tk.W)
        tree.heading("Experiencia", text="Experiencia", anchor=tk.W)

        tree.column("ID Empleado", width=100, minwidth=80, stretch=tk.NO)
        tree.column("Nombre", width=280, minwidth=200, stretch=tk.YES) 
        tree.column("Línea", width=80, minwidth=60, stretch=tk.NO)
        tree.column("Puesto", width=120, minwidth=100, stretch=tk.NO)
        tree.column("Antiguedad", width=180, minwidth=150, stretch=tk.NO)
        tree.column("Experiencia", width=100, minwidth=80, stretch=tk.NO)

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=scrollbar_x.set)


        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X) 

        return tree

    def update_tables_event(self):
        """Evento para actualizar tablas (usado por el bind de la ventana principal)."""
        self.update_tables()

    def update_tables(self):
        """Actualiza todas las tablas Treeview con los datos actuales."""
        # ... (código existente para llenar las tablas) ...
        # Aquí continúa tu código para update_tables y on_close, etc.
        # Solo puse un ejemplo del inicio del método update_tables().
        
        # Para que el código sea ejecutable, necesitas el resto de los métodos y la ejecución del mainloop.
        # No los incluí completos aquí para enfocarme en los cambios de color.
        # Asegúrate de copiar solo las líneas modificadas o el código completo actualizado.
        
        # Ejemplo para limpiar y rellenar una tabla (dentro de update_tables):
        for tree in self.tabs.values():
            for item in tree.get_children():
                tree.delete(item)
        
        # Luego, rellenas las tablas según la lógica de tu aplicación
        # Por ejemplo, para la pestaña "Todos":
        all_employees_tree = self.tabs["Todos"]
        for employee_id, data in self.scanned_employees_data.items():
            antiguedad_str = f"{data['Antiguedad_Anos']:.1f} años ({data['Antiguedad_Dias']} días)"
            experiencia_str = "Sí" if self.app_instance.is_experienced(data['F_Servicio']) else "No"
            all_employees_tree.insert("", tk.END, values=(
                employee_id,
                data['Nombre'],
                data['Linea'],
                data['Puesto'],
                antiguedad_str,
                experiencia_str
            ))

        self.filter_tables() # Vuelve a aplicar el filtro de búsqueda después de actualizar los datos


    def filter_tables(self, event=None):
        """Filtra las tablas Treeview basándose en el texto de búsqueda."""
        search_term = self.search_entry.get().lower()

        employee_data_list = []
        for employee_id, data in self.scanned_employees_data.items():
            antiguedad_str = f"{data['Antiguedad_Anos']:.1f} años ({data['Antiguedad_Dias']} días)"
            experiencia_str = "Sí" if self.app_instance.is_experienced(data['F_Servicio']) else "No"
            employee_data_list.append({
                'ID Empleado': employee_id,
                'Nombre': data['Nombre'],
                'Línea': data['Linea'],
                'Puesto': data['Puesto'],
                'Antigüedad_Anos': data['Antiguedad_Anos'],
                'Antiguedad_Dias': data['Antiguedad_Dias'],
                'Experiencia_Bool': self.app_instance.is_experienced(data['F_Servicio']),
                'POSITION': data['POSITION'],
                'Antiguedad_Str': antiguedad_str,
                'Experiencia_Str': experiencia_str
            })

        for tab_name, tree in self.tabs.items():
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)

            filtered_data = []
            for emp in employee_data_list:
                match_search = search_term in emp['ID Empleado'].lower() or search_term in emp['Nombre'].lower()

                if tab_name == "Todos":
                    if match_search:
                        filtered_data.append(emp)
                elif tab_name == "Op. con Experiencia":
                    if match_search and emp['POSITION'].lower() == 'mfgupo' and emp['Experiencia_Bool']:
                        filtered_data.append(emp)
                elif tab_name == "Op. sin Experiencia":
                    if match_search and emp['POSITION'].lower() == 'mfgupo' and not emp['Experiencia_Bool']:
                        filtered_data.append(emp)
                elif tab_name == "Soportes":
                    if match_search and emp['POSITION'].lower() == 'mfgupo' and emp['Puesto'].lower() == 'mfgsup': # Asumiendo 'mfgsup' para soportes
                         filtered_data.append(emp)
                elif tab_name == "Calidad":
                    if match_search and emp['POSITION'].lower() == 'qainsp':
                        filtered_data.append(emp)
                elif tab_name == "Op. Prestados":
                    if match_search and emp['POSITION'].lower() == 'mfgupo' and emp['Línea'].lower() != self.selected_line_for_tab.lower():
                        filtered_data.append(emp)

            for emp in filtered_data:
                tree.insert("", tk.END, values=(emp['ID Empleado'], emp['Nombre'], emp['Línea'], emp['Puesto'], emp['Antiguedad_Str'], emp['Experiencia_Str']))

    def on_close(self):
        """Maneja el cierre de la ventana secundaria."""
        self.app_instance.txt_escaneo.focus_set() # Regresar el foco al campo de escaneo
        self.destroy()

    def on_resize(self, event):
        """Ajusta el ancho de las columnas del Treeview cuando la ventana se redimensiona."""
        # Se activa cuando la ventana cambia de tamaño
        for tab_name, tree in self.tabs.items():
            # Obtener el ancho actual del Treeview
            total_width = tree.winfo_width()
            
            # Anchos fijos (para columnas que no deben estirarse)
            id_width = 100
            line_width = 80
            puesto_width = 120
            antiguedad_width = 180
            experiencia_width = 100
            
            # Ancho disponible para la columna de Nombre (que debe estirarse)
            remaining_width = total_width - (id_width + line_width + puesto_width + antiguedad_width + experiencia_width + 50) # El 50 es un ajuste para el scrollbar y padding
            
            # Asegurarse de que el ancho no sea negativo
            if remaining_width < 100: # Mínimo para Nombre
                remaining_width = 100
            
            tree.column("ID Empleado", width=id_width, minwidth=id_width)
            tree.column("Nombre", width=remaining_width, minwidth=100)
            tree.column("Línea", width=line_width, minwidth=line_width)
            tree.column("Puesto", width=puesto_width, minwidth=puesto_width)
            tree.column("Antigüedad", width=antiguedad_width, minwidth=antiguedad_width)
            tree.column("Experiencia", width=experiencia_width, minwidth=experiencia_width)



if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()