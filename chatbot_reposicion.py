import tkinter as tk
from tkinter import scrolledtext, messagebox
import pandas as pd
import os
from datetime import datetime

# --- DatabaseManager Class ---
class DatabaseManager:
    """
    Gestiona la carga y consulta de datos desde el archivo Excel BDD.xlsx.
    """
    def __init__(self, filename="BDD.xlsx"):
        self.filename = filename
        self.df = self._load_data()

    def _load_data(self):
        """Carga los datos del archivo Excel."""
        if not os.path.exists(self.filename):
            messagebox.showerror("Error de Archivo", f"El archivo '{self.filename}' no se encontró.")
            return pd.DataFrame()
        try:
            # Columnas requeridas para la aplicación.
            # Cargar solo estas columnas reduce el tiempo de lectura de archivos grandes.
            required_cols = [
                'Numero Sencillo', 'Codigos', 'Cod A', 'Cod B', 'Proceso', 
                'Maq', 'Ckt Grp', 'Type', 'Size', 'Color', 'Cut Length', 'General', 'Planta', 'Qty'
            ]
            
            # Leer solo los encabezados para verificar qué columnas existen realmente en el archivo Excel
            excel_columns = pd.read_excel(self.filename, sheet_name=0, nrows=0).columns
            
            # Filtrar las columnas requeridas para incluir solo las que existen en el Excel
            cols_to_use = [col for col in required_cols if col in excel_columns]

            df = pd.read_excel(self.filename, usecols=cols_to_use)
            
            # Asegurarse de que las columnas de código sean de tipo string para búsquedas consistentes
            # y se stripteen espacios en blanco
            # Se aplica solo a las columnas que fueron cargadas y que son relevantes para el stripping
            cols_to_strip_and_str = [
                'Numero Sencillo', 'Codigos', 'Cod A', 'Cod B', 'Proceso', 
                'Maq', 'Ckt Grp', 'Type', 'Size', 'Color', 'Cut Length', 'General', 'Planta'
            ]
            
            for col in cols_to_strip_and_str:
                if col in df.columns: # Verificar si la columna existe en el DataFrame cargado
                    df[col] = df[col].astype(str).str.strip()
            return df
        except Exception as e:
            messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo Excel: {e}")
            return pd.DataFrame()

    def find_direct_code(self, code):
        """Busca un código directo en 'Numero Sencillo' o 'Codigos' y devuelve la fila completa."""
        if self.df.empty:
            return None
        result = self.df[(self.df['Numero Sencillo'] == code) | (self.df['Codigos'] == code)]
        return result if not result.empty else None

    def find_process_related_codes(self, input_code_or_process):
        """
        Busca todos los códigos relacionados con un código de proceso o un código de producto.
        Si se da un código de producto (Columna K), encuentra su proceso asociado (Columna M)
        y luego devuelve todos los códigos de ese proceso con toda su información.
        Devuelve el DataFrame con los items encontrados y el código de proceso identificado.
        """
        if self.df.empty:
            return None, None # Devuelve None para resultados y para el proceso_identificado

        # 1. Intentar encontrar la entrada como un Código de Proceso (Columna M)
        found_by_process = self.df[self.df['Proceso'] == input_code_or_process]
        if not found_by_process.empty:
            return found_by_process, input_code_or_process

        # 2. Intentar encontrar la entrada como un Código de Producto (Columna K)
        found_by_codigo_producto = self.df[self.df['Codigos'] == input_code_or_process]
        if not found_by_codigo_producto.empty:
            # Si se encuentra como código de producto, obtener su código de proceso
            identified_process = found_by_codigo_producto.iloc[0]['Proceso']
            # Y luego buscar todos los elementos de ese proceso
            all_codes_for_this_process = self.df[self.df['Proceso'] == identified_process]
            return all_codes_for_this_process, identified_process
        
        # Si no se encontró ni como proceso ni como código de producto
        return None, None
    
    def find_code_in_process(self, process_df, code_to_find):
        """
        Busca un código específico dentro de un DataFrame que representa un proceso.
        Retorna la fila del DataFrame si la encuentra, de lo contrario None.
        """
        if process_df.empty:
            return None
        
        # Buscar en 'Numero Sencillo' o 'Codigos' dentro del DataFrame del proceso
        result = process_df[(process_df['Numero Sencillo'] == code_to_find) | (process_df['Codigos'] == code_to_find)]
        return result.iloc[0] if not result.empty else None


# --- ChatbotApp Class ---
class ChatbotApp:
    """
    Clase principal para la aplicación de chatbot en Tkinter.
    Maneja la interfaz de usuario y la lógica de la conversación.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Asistente de Reposición Virtual")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        self.db_manager = DatabaseManager()
        self.conversation_state = {}
        self.history = []

        self._create_widgets()
        self._start_conversation()

    def _create_widgets(self):
        """Crea y posiciona los widgets de la interfaz."""
        # Área de conversación
        self.chat_display = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, state='disabled', font=("Arial", 10), bg="#e0e0e0")
        self.chat_display.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Frame para entrada y botón
        input_frame = tk.Frame(self.root)
        input_frame.pack(padx=10, pady=(0, 10), fill=tk.X)

        self.user_input = tk.Entry(input_frame, font=("Arial", 10))
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        self.user_input.bind("<Return>", self._send_message_event) # Permite enviar con ENTER

        self.send_button = tk.Button(input_frame, text="Enviar", command=self._send_message, font=("Arial", 10, "bold"), bg="#4CAF50", fg="white")
        self.send_button.pack(side=tk.RIGHT, padx=(5, 0), ipadx=10, ipady=3)

    def _display_message(self, sender, message, color="black"):
        """Muestra un mensaje en el área de conversación."""
        self.chat_display.config(state='normal')
        self.chat_display.insert(tk.END, f"{sender}: ", "sender_tag")
        self.chat_display.insert(tk.END, f"{message}\n", "message_tag")
        self.chat_display.tag_config("sender_tag", foreground="#007bff" if sender == "Bot" else "#0056b3", font=("Arial", 10, "bold"))
        self.chat_display.tag_config("message_tag", foreground=color, font=("Arial", 10))
        self.chat_display.yview(tk.END)
        self.chat_display.config(state='disabled')

    def _start_conversation(self):
        """Inicia la conversación con el saludo inicial."""
        self._display_message("Bot", "Hola, soy tu Asistente personal Virtual.")
        self.root.after(500, self._ask_initial_reposition) # Pequeño retardo para mejor UX

    def _ask_initial_reposition(self):
        """Pregunta si el usuario desea realizar una reposición."""
        self._display_message("Bot", "¿Desea realizar una reposición? (Sí/No)")
        self.conversation_state = {"step": "ask_reposition"}

    def _send_message_event(self, event=None):
        """Maneja el evento de envío de mensaje (Enter key)."""
        self._send_message()

    def _send_message(self):
        """Procesa el mensaje enviado por el usuario."""
        user_text = self.user_input.get().strip().lower()
        if not user_text:
            return

        self._display_message("Tú", user_text)
        self.history.append(("user", user_text))
        self.user_input.delete(0, tk.END)

        self._process_user_response(user_text)

    def _process_user_response(self, response):
        """Lógica principal para procesar las respuestas del usuario."""
        step = self.conversation_state.get("step")

        if step == "ask_reposition":
            if response in ["si", "sí", "s"]:
                self._display_message("Bot", "¿Desea reponer un **directo** o un **proceso**?")
                self.conversation_state = {"step": "ask_type"}
            elif response in ["no", "n"]:
                self._display_message("Bot", "Entendido. No se realizará ninguna reposición. ¡Hasta luego!")
                self.root.after(2000, self.root.destroy) # Cierra la app después de 2 segundos
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

        elif step == "ask_type":
            if response == "directo":
                self._display_message("Bot", "Por favor, ingrese el **código del directo**:")
                self.conversation_state = {"step": "get_direct_code", "type": "directo"}
            elif response == "proceso":
                self._display_message("Bot", "Por favor, ingrese el **código de proceso** o un **código de producto** relacionado (Columna K):")
                self.conversation_state = {"step": "get_process_code", "type": "proceso"}
            else:
                self._display_message("Bot", "Por favor, responda 'directo' o 'proceso'.")

        elif step == "get_direct_code":
            self.conversation_state["code"] = response.upper()
            found_code_df = self.db_manager.find_direct_code(response.upper())
            if found_code_df is not None and not found_code_df.empty:
                self.conversation_state["found_item"] = found_code_df.iloc[0] # Almacena la fila completa
                
                # Mostrar detalles principales del directo encontrado al usuario
                item = self.conversation_state["found_item"]
                details_msg = (
                    f"He encontrado el código: **{item.get('Numero Sencillo', 'N/A')}** "
                    f"({item.get('Codigos', 'N/A')}).\n"
                    f"Tipo: {item.get('Type', 'N/A')}, Tamaño: {item.get('Size', 'N/A')}, "
                    f"Color: {item.get('Color', 'N/A')}, Largo Corte: {item.get('Cut Length', 'N/A')}.\n"
                    f"¿Es este el artículo que desea reponer? (Sí/No)"
                )
                self._display_message("Bot", details_msg)
                self.conversation_state["step"] = "confirm_direct_item"
            else:
                self._display_message("Bot", f"El código directo '{response.upper()}' no fue encontrado en la base de datos. Por favor, intente de nuevo.")
                # Permanece en el mismo paso para reingresar el código

        elif step == "confirm_direct_item":
            if response in ["si", "sí", "s"]:
                self._display_message("Bot", "¿Cuántas piezas desea reponer?")
                self.conversation_state["step"] = "get_quantity"
            elif response in ["no", "n"]:
                self._display_message("Bot", "Entendido. Por favor, ingrese el **código del directo** correcto:")
                self.conversation_state["step"] = "get_direct_code" # Vuelve al paso de pedir código
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

        elif step == "get_quantity":
            try:
                quantity = int(response)
                if quantity <= 0:
                    raise ValueError
                self.conversation_state["quantity"] = quantity
                self._display_message("Bot", "¿Desea realizar otra reposición? (Sí/No)")
                self.conversation_state["step"] = "ask_another_reposition"
            except ValueError:
                self._display_message("Bot", "Cantidad inválida. Por favor, ingrese un número entero positivo.")

        elif step == "ask_another_reposition":
            if response in ["si", "sí", "s"]:
                self.conversation_state = {} # Reinicia el estado para una nueva reposición
                self._display_message("Bot", "Reiniciando el proceso.")
                self.root.after(500, self._ask_initial_reposition)
            elif response in ["no", "n"]:
                self._display_message("Bot", "¿Desea **imprimir** la información de la reposición? (Sí/No)")
                self.conversation_state["step"] = "ask_print"
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

        elif step == "ask_print":
            if response in ["si", "sí", "s"]:
                self._print_reposition_info()
                self._display_message("Bot", "Reposición completada y enviada a impresión. ¡Gracias por usar el asistente!")
                self.root.after(2000, self.root.destroy)
            elif response in ["no", "n"]:
                self._display_message("Bot", "Reposición completada. No se realizará la impresión. ¡Gracias por usar el asistente!")
                self.root.after(2000, self.root.destroy)
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

        elif step == "get_process_code":
            input_value = response.upper()
            found_items_df, identified_process_code = self.db_manager.find_process_related_codes(input_value)
            
            if found_items_df is not None and not found_items_df.empty:
                self.conversation_state["found_processes"] = found_items_df # DataFrame completo del proceso
                self.conversation_state["process_code_identified"] = identified_process_code # Almacena el proceso real
                
                details_msg = (f"Hemos identificado el proceso: **{identified_process_code}**.\n"
                               f"Este proceso incluye los siguientes códigos generales:\n")
                
                # Modificación para evitar duplicados al listar códigos del proceso
                displayed_codes = set() # Usar un conjunto para almacenar pares (Numero Sencillo, Codigos) ya mostrados
                display_count = 0
                display_limit = 10 # Se aumenta un poco el límite para mostrar más, si existen
                
                for index, row in found_items_df.iterrows():
                    sencillo = row.get('Numero Sencillo', 'N/A')
                    general = row.get('Codigos', 'N/A')
                    code_pair = (sencillo, general) # Tupla para identificar el par único

                    if code_pair not in displayed_codes:
                        details_msg += (
                            f"- Sencillo: {sencillo} "
                            f"(General: {general})\n"
                        )
                        displayed_codes.add(code_pair)
                        display_count += 1
                        if display_count >= display_limit:
                            break # Limitar el número de elementos mostrados inicialmente
                
                if len(found_items_df) > display_count: # Corrección para el mensaje de "y X más"
                    details_msg += f"...y más códigos relacionados.\n"
                
                details_msg += "\n¿Es este el proceso que desea reponer? (Sí/No)"

                self._display_message("Bot", details_msg)
                self.conversation_state["step"] = "confirm_process_items"
            else:
                self._display_message("Bot", f"El código o proceso '{input_value}' no fue encontrado en la base de datos. Por favor, intente de nuevo.")
                # Permanece en el mismo paso para reingresar el código

        elif step == "confirm_process_items":
            if response in ["si", "sí", "s"]:
                self._display_message("Bot", "¿Desea reponer el **grupo completo** de este proceso o un **circuito específico** dentro de él? (Grupo/Especifico)")
                self.conversation_state["step"] = "ask_group_or_specific"
            elif response in ["no", "n"]:
                self._display_message("Bot", "Entendido. Por favor, ingrese el **código de proceso** o un **código de producto** relacionado correcto:")
                self.conversation_state["step"] = "get_process_code" # Vuelve al paso de pedir código de proceso
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

        elif step == "ask_group_or_specific":
            if response == "grupo":
                self._display_message("Bot", "¿Cuál es la **cantidad total** de piezas para el grupo completo?")
                self.conversation_state["step"] = "get_total_group_quantity"
                self.conversation_state["reposition_scope"] = "full_group"
            elif response == "especifico" or response == "específico":
                # Aquí se pide el código específico
                self._display_message("Bot", "Por favor, ingrese el **código del circuito específico** que desea reponer (Numero Sencillo o Código General):")
                self.conversation_state["step"] = "ask_for_specific_process_code"
                self.conversation_state["reposition_scope"] = "single_circuit" # Marcamos el alcance
            else:
                self._display_message("Bot", "Por favor, responda 'Grupo' o 'Especifico'.")
        
        elif step == "ask_for_specific_process_code":
            specific_code = response.upper()
            process_df = self.conversation_state.get("found_processes")
            found_item_in_process = self.db_manager.find_code_in_process(process_df, specific_code)

            if found_item_in_process is not None:
                self.conversation_state["found_item"] = found_item_in_process # Almacena la fila del item específico
                item = found_item_in_process
                details_msg = (
                    f"He encontrado el circuito: **{item.get('Numero Sencillo', 'N/A')}** "
                    f"({item.get('Codigos', 'N/A')}).\n"
                    f"Tipo: {item.get('Type', 'N/A')}, Tamaño: {item.get('Size', 'N/A')}, "
                    f"Color: {item.get('Color', 'N/A')}, Largo Corte: {item.get('Cut Length', 'N/A')}.\n"
                    f"¿Es este el circuito que desea reponer? (Sí/No)"
                )
                self._display_message("Bot", details_msg)
                self.conversation_state["step"] = "confirm_specific_process_item"
            else:
                self._display_message("Bot", f"El código '{specific_code}' no fue encontrado en este proceso. Por favor, revise e intente de nuevo.")
                # Permanece en el mismo paso para reingresar el código

        elif step == "confirm_specific_process_item":
            if response in ["si", "sí", "s"]:
                self._display_message("Bot", "¿Cuántas piezas desea reponer para este circuito?")
                self.conversation_state["step"] = "get_single_circuit_quantity"
            elif response in ["no", "n"]:
                self._display_message("Bot", "Entendido. Por favor, ingrese el **código del circuito específico** correcto:")
                self.conversation_state["step"] = "ask_for_specific_process_code"
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

        elif step == "get_total_group_quantity":
            try:
                quantity = int(response)
                if quantity <= 0:
                    raise ValueError
                self.conversation_state["quantity"] = quantity
                self._display_message("Bot", "¿Desea imprimir la información de la reposición? (Sí/No)")
                self.conversation_state["step"] = "ask_print_process"
            except ValueError:
                self._display_message("Bot", "Cantidad inválida. Por favor, ingrese un número entero positivo.")

        elif step == "get_single_circuit_quantity":
            try:
                quantity = int(response)
                if quantity <= 0:
                    raise ValueError
                self.conversation_state["quantity"] = quantity
                self._display_message("Bot", "¿Desea imprimir la información de la reposición? (Sí/No)")
                self.conversation_state["step"] = "ask_print_process"
            except ValueError:
                self._display_message("Bot", "Cantidad inválida. Por favor, ingrese un número entero positivo.")

        elif step == "ask_print_process":
            if response in ["si", "sí", "s"]:
                self._print_reposition_info()
                self._display_message("Bot", "Reposición de proceso completada y enviada a impresión. ¡Gracias por usar el asistente!")
                self.root.after(2000, self.root.destroy)
            elif response in ["no", "n"]:
                self._display_message("Bot", "Reposición de proceso completada. No se realizará la impresión. ¡Gracias por usar el asistente!")
                self.root.after(2000, self.root.destroy)
            else:
                self._display_message("Bot", "Por favor, responda 'Sí' o 'No'.")

    def _print_reposition_info(self):
        """Simula la impresión de la información de la reposición."""
        reposition_type = self.conversation_state.get("type")
        output_filename = f"reposicion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(f"--- REPORTE DE REPOSICIÓN ---\n")
            f.write(f"Fecha y Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Tipo de Reposición: {reposition_type.upper()}\n")
            f.write("-" * 30 + "\n\n")

            if reposition_type == "directo":
                item = self.conversation_state.get("found_item")
                quantity_to_print = self.conversation_state.get("quantity") # Cantidad proporcionada por el usuario

                f.write(f"Código Directo: {item.get('Numero Sencillo', 'N/A')}\n")
                f.write(f"Cantidad a Reponer: {quantity_to_print} piezas\n\n")

                f.write(f"{'NUMERO DE PARTE':<28}{'CODIGO':<15}{'CIRCUITO A':<15}{'CIRCUITO B':<15}{'PROCESO':<15}{'CANTIDAD':<12}{'GRUPO(SI/NO)':<15}{'PLANTA':<10}\n")
                f.write(f"{'-'*28:<28}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*12:<12}{'-'*15:<15}{'-'*10:<10}\n")

                # Obtener valores para la línea de formato
                numero_parte = item.get('Numero Sencillo', 'N/A')
                codigo_general = item.get('Codigos', 'N/A')
                circuito_a = item.get('Cod A', 'N/A')
                circuito_b = item.get('Cod B', 'N/A')
                proceso_tipo = item.get('Proceso', 'N/A') # Si es "Directo" o el nombre del proceso
                grupo_status = "NO" # Para un directo, no es parte de un grupo de reposición
                planta = item.get('Planta', 'N/A') # Añadir información de planta

                f.write(f"{numero_parte:<28}{codigo_general:<15}{circuito_a:<15}{circuito_b:<15}{proceso_tipo:<15}{quantity_to_print:<12}{grupo_status:<15}{planta:<10}\n\n")
                
                f.write("Detalles Completos del Artículo (BDD):\n")
                for col in item.index:
                    f.write(f"  {col}: {item[col]}\n")
                
                # Ejemplo de impresión para directo (consola)
                print("\n" + "="*80)
                print("--- REPORTE DE IMPRESIÓN (DIRECTO) ---")
                print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"Tipo: DIRECTO")
                print(f"Código Directo: {item.get('Numero Sencillo', 'N/A')}")
                print(f"Cantidad a Reponer: {quantity_to_print} piezas\n")

                print(f"{'NUMERO DE PARTE':<28}{'CODIGO':<15}{'CIRCUITO A':<15}{'CIRCUITO B':<15}{'PROCESO':<15}{'CANTIDAD':<12}{'GRUPO(SI/NO)':<15}{'PLANTA':<10}")
                print(f"{'-'*28:<28}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*12:<12}{'-'*15:<15}{'-'*10:<10}")
                print(f"{numero_parte:<28}{codigo_general:<15}{circuito_a:<15}{circuito_b:<15}{proceso_tipo:<15}{quantity_to_print:<12}{grupo_status:<15}{planta:<10}\n")

                print("Detalles Completos del Artículo (BDD):")
                for col in item.index:
                    print(f"  {col}: {item[col]}")
                print("="*80 + "\n")

            elif reposition_type == "proceso":
                process_code_identified = self.conversation_state.get("process_code_identified")
                scope = self.conversation_state.get("reposition_scope")
                quantity_to_print = self.conversation_state.get("quantity") # Cantidad proporcionada por el usuario
                found_processes_df = self.conversation_state.get("found_processes") # DataFrame completo de items del proceso

                f.write(f"Reposición para Proceso:\n")
                f.write(f"  Código de Proceso Identificado: {process_code_identified}\n")
                f.write(f"  Alcance de Reposición: {'Grupo Completo' if scope == 'full_group' else 'Circuito Específico'}\n")
                f.write(f"  Cantidad Solicitada: {quantity_to_print} piezas\n\n")

                f.write(f"{'NUMERO DE PARTE':<28}{'CODIGO':<15}{'CIRCUITO A':<15}{'CIRCUITO B':<15}{'PROCESO':<15}{'CANTIDAD':<12}{'GRUPO(SI/NO)':<15}{'PLANTA':<10}\n")
                f.write(f"{'-'*28:<28}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*12:<12}{'-'*15:<15}{'-'*10:<10}\n")

                if scope == "full_group":
                    # Almacenar códigos de parte y generales únicos para el reporte detallado
                    reported_items = set()
                    
                    for index, row in found_processes_df.iterrows():
                        numero_parte = row.get('Numero Sencillo', 'N/A')
                        codigo_general = row.get('Codigos', 'N/A')
                        circuito_a = row.get('Cod A', 'N/A')
                        circuito_b = row.get('Cod B', 'N/A')
                        proceso_tipo = row.get('Proceso', 'N/A') 
                        grupo_status = "SI" # Es parte de un grupo de reposición
                        planta = row.get('Planta', 'N/A') # Añadir información de planta

                        # Solo imprimir la línea resumen una vez por combinación única de Numero Sencillo y Codigos
                        if (numero_parte, codigo_general) not in reported_items:
                            f.write(f"{numero_parte:<28}{codigo_general:<15}{circuito_a:<15}{circuito_b:<15}{proceso_tipo:<15}{quantity_to_print:<12}{grupo_status:<15}{planta:<10}\n")
                            reported_items.add((numero_parte, codigo_general))

                    f.write("\nDetalles Completos de los Artículos del Grupo (BDD) - Sin duplicados en la lista:\n")
                    # Crear una lista de tuplas únicas (Numero Sencillo, Codigos) para los detalles completos
                    unique_items_details = []
                    seen_pairs = set()
                    for index, row in found_processes_df.iterrows():
                        sencillo = row.get('Numero Sencillo', 'N/A')
                        general = row.get('Codigos', 'N/A')
                        if (sencillo, general) not in seen_pairs:
                            unique_items_details.append(row)
                            seen_pairs.add((sencillo, general))

                    for i, row in enumerate(unique_items_details):
                        f.write(f"\n--- Item Único {i+1} ({row.get('Numero Sencillo', 'N/A')}) ---\n")
                        for col in row.index:
                            f.write(f"  {col}: {row[col]}\n")

                elif scope == "single_circuit":
                    # Usamos 'found_item' que ahora almacena la fila del circuito específico seleccionado
                    item_info = self.conversation_state.get("found_item")
                    if item_info is not None:
                        numero_parte = item_info.get('Numero Sencillo', 'N/A')
                        codigo_general = item_info.get('Codigos', 'N/A')
                        circuito_a = item_info.get('Cod A', 'N/A')
                        circuito_b = item_info.get('Cod B', 'N/A')
                        proceso_tipo = item_info.get('Proceso', 'N/A') 
                        grupo_status = "NO" # No es una reposición de grupo completo, es un circuito específico
                        planta = item_info.get('Planta', 'N/A') # Añadir información de planta

                        f.write(f"{numero_parte:<28}{codigo_general:<15}{circuito_a:<15}{circuito_b:<15}{proceso_tipo:<15}{quantity_to_print:<12}{grupo_status:<15}{planta:<10}\n")
                        
                        f.write("\nDetalles Completos del Circuito Específico (BDD):\n")
                        for col in item_info.index:
                            f.write(f"  {col}: {item_info[col]}\n")
                    else:
                        f.write("  No se encontraron detalles para el circuito específico seleccionado.\n")

                # Ejemplo de impresión para proceso (consola)
                print("\n" + "="*80)
                print("--- REPORTE DE IMPRESIÓN (PROCESO) ---")
                print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"Tipo: PROCESO")
                print(f"Código de Proceso Identificado: {process_code_identified}")
                print(f"Alcance: {'Grupo Completo' if scope == 'full_group' else 'Circuito Específico'}")
                print(f"Cantidad a Reponer: {quantity_to_print} piezas\n")

                print(f"{'NUMERO DE PARTE':<28}{'CODIGO':<15}{'CIRCUITO A':<15}{'CIRCUITO B':<15}{'PROCESO':<15}{'CANTIDAD':<12}{'GRUPO(SI/NO)':<15}{'PLANTA':<10}")
                print(f"{'-'*28:<28}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*15:<15}{'-'*12:<12}{'-'*15:<15}{'-'*10:<10}")

                if scope == "full_group":
                    # Almacenar códigos de parte y generales únicos para la impresión en consola
                    reported_items_console = set()
                    for index, row in found_processes_df.iterrows():
                        numero_parte = row.get('Numero Sencillo', 'N/A')
                        codigo_general = row.get('Codigos', 'N/A')
                        circuito_a = row.get('Cod A', 'N/A')
                        circuito_b = row.get('Cod B', 'N/A')
                        proceso_tipo = row.get('Proceso', 'N/A')
                        grupo_status = "SI"
                        planta = row.get('Planta', 'N/A')
                        if (numero_parte, codigo_general) not in reported_items_console:
                            print(f"{numero_parte:<28}{codigo_general:<15}{circuito_a:<15}{circuito_b:<15}{proceso_tipo:<15}{quantity_to_print:<12}{grupo_status:<15}{planta:<10}")
                            reported_items_console.add((numero_parte, codigo_general))
                    
                    print("\nDetalles Completos de los Artículos del Grupo (BDD) - Sin duplicados en la lista:")
                    # Reutilizar unique_items_details para la consola
                    for i, row in enumerate(unique_items_details):
                        print(f"\n--- Item Único {i+1} ({row.get('Numero Sencillo', 'N/A')}) ---")
                        for col in row.index:
                            print(f"  {col}: {row[col]}")

                elif scope == "single_circuit":
                    item_info = self.conversation_state.get("found_item")
                    if item_info is not None:
                        numero_parte = item_info.get('Numero Sencillo', 'N/A')
                        codigo_general = item_info.get('Codigos', 'N/A')
                        circuito_a = item_info.get('Cod A', 'N/A')
                        circuito_b = item_info.get('Cod B', 'N/A')
                        proceso_tipo = item_info.get('Proceso', 'N/A')
                        grupo_status = "NO"
                        planta = item_info.get('Planta', 'N/A')
                        print(f"{numero_parte:<28}{codigo_general:<15}{circuito_a:<15}{circuito_b:<15}{proceso_tipo:<15}{quantity_to_print:<12}{grupo_status:<15}{planta:<10}")
                        print("\nDetalles Completos del Circuito Específico (BDD):")
                        for col in item_info.index:
                            print(f"  {col}: {item_info[col]}")
                    else:
                        print("  No se encontraron detalles para el circuito específico seleccionado.")
                print("="*80 + "\n")

        messagebox.showinfo("Impresión Simulada", f"Información de reposición guardada en '{output_filename}'")
        self.history.append(("bot", f"Reporte guardado en {output_filename}"))

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ChatbotApp(root)
    root.mainloop()