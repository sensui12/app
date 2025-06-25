import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import pandas as pd
import os
import sqlite3
from datetime import datetime
import time
import random
import copy

# --- Constantes ---
SQLITE_DB_FILENAME = "reposicion_data.db"
ITEMS_TABLE_NAME = "items_reposicion"
OUTPUT_FILE_PREFIX = "reposicion_"
OUTPUT_EXCEL_SHEET_NAME = "Reporte_Reposicion"

REQUIRED_DB_COLUMNS = [
    'Numero Sencillo', 'Codigos', 'Cod A', 'Cod B', 'Proceso', 'Maq',
    'Ckt Grp', 'Type', 'Size', 'Color', 'Cut Length', 'General', 'Planta', 'Qty'
]
SEARCH_OPTIMIZED_COLUMNS = ['Numero Sencillo', 'Codigos', 'Proceso', 'Maq']
SPECIAL_PROCESSES = ["TW", "BR"]
MACHINE_TO_EXCLUDE = "TW01"

CONV_STEP_ASK_REPOSITION = "ask_reposition"
CONV_STEP_ASK_TYPE = "ask_type"
CONV_STEP_GET_DIRECT_CODE = "get_direct_code"
CONV_STEP_CONFIRM_DIRECT_ITEM = "confirm_direct_item"
CONV_STEP_GET_QUANTITY_DIRECT = "get_quantity_direct"
CONV_STEP_ASK_ANOTHER_REPOSITION = "ask_another_reposition"
CONV_STEP_ASK_PRINT = "ask_print"

CONV_STEP_GET_PROCESS_CODE = "get_process_code"
CONV_STEP_CONFIRM_PROCESS_ITEMS = "confirm_process_items"
CONV_STEP_ASK_GROUP_OR_SPECIFIC = "ask_group_or_specific"
CONV_STEP_ASK_FOR_SPECIFIC_PROCESS_CODE = "ask_for_specific_process_code"
CONV_STEP_CONFIRM_SPECIFIC_PROCESS_ITEM = "confirm_specific_process_item"
CONV_STEP_GET_TOTAL_GROUP_QUANTITY = "get_total_group_quantity"
CONV_STEP_GET_SINGLE_CIRCUIT_QUANTITY = "get_single_circuit_quantity"

USER_YES_RESPONSES = ["si", "sí", "s"]
USER_NO_RESPONSES = ["no", "n"]
USER_DIRECT_RESPONSE = "directo"
USER_PROCESS_RESPONSE = "proceso"
USER_GROUP_RESPONSE = "grupo"
USER_SPECIFIC_RESPONSE = "especifico"
USER_SPECIFIC_RESPONSE_ACCENT = "específico"

# --- DatabaseManager Class ---
class DatabaseManager:
    def __init__(self, db_filename=SQLITE_DB_FILENAME):
        self.db_filename = db_filename
        self.df_original = pd.DataFrame()
        self.df_search = pd.DataFrame()
        self._load_data_from_sql()

    def _set_empty_dfs(self):
        self.df_original = pd.DataFrame()
        self.df_search = pd.DataFrame()

    def _create_sql_connection(self):
        conn = None
        try:
            conn = sqlite3.connect(self.db_filename)
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"Error al conectar a la BD SQLite: {e}")
        return conn

    def _table_exists(self, table_name):
        conn = self._create_sql_connection()
        if not conn: return False
        try:
            cursor = conn.cursor()
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}';")
            return cursor.fetchone() is not None
        except sqlite3.Error as e:
            print(f"Error al verificar si la tabla existe: {e}")
            return False
        finally:
            if conn: conn.close()

    def _load_data_from_sql(self):
        if not self._table_exists(ITEMS_TABLE_NAME):
            print(f"Tabla '{ITEMS_TABLE_NAME}' no encontrada. Use el botón de carga.")
            self._set_empty_dfs()
            return

        conn = self._create_sql_connection()
        if not conn:
            self._set_empty_dfs()
            return

        try:
            start_time = time.time()
            self.df_original = pd.read_sql_query(f"SELECT * FROM {ITEMS_TABLE_NAME}", conn, dtype=str)

            if self.df_original.empty:
                print("La tabla de ítems está vacía en la BD.")
                self._set_empty_dfs()
                return

            self.df_search = self.df_original.copy()
            for col in self.df_search.columns:
                if col in SEARCH_OPTIMIZED_COLUMNS:
                    self.df_search[col] = self.df_search[col].astype(str).str.strip().str.upper()
                elif col in REQUIRED_DB_COLUMNS:
                     self.df_search[col] = self.df_search[col].astype(str).str.strip()

            if 'Maq' in self.df_search.columns and 'Maq' not in SEARCH_OPTIMIZED_COLUMNS:
                 self.df_search['Maq'] = self.df_search['Maq'].astype(str).str.strip().str.upper()

            end_time = time.time()
            print(f"Datos cargados desde SQL y preprocesados en {end_time - start_time:.4f}s.")
        except Exception as e:
            messagebox.showerror("Error al Cargar Datos SQL", f"No se pudo cargar datos desde SQL: {e}")
            self._set_empty_dfs()
        finally:
            if conn: conn.close()

    def load_excel_to_sql(self, excel_filepath):
        try:
            start_time = time.time()
            df_excel = pd.read_excel(excel_filepath, dtype=str)

            missing_cols = [col for col in REQUIRED_DB_COLUMNS if col not in df_excel.columns]
            if any(col not in df_excel.columns for col in ['Numero Sencillo', 'Codigos', 'Proceso']):
                 messagebox.showerror("Error de Excel",
                                     f"El Excel debe contener al menos 'Numero Sencillo', 'Codigos', y 'Proceso'.")
                 return False

            conn = self._create_sql_connection()
            if not conn: return False

            for col in df_excel.columns:
                if col in REQUIRED_DB_COLUMNS:
                    df_excel[col] = df_excel[col].astype(str).str.strip()

            df_excel.to_sql(ITEMS_TABLE_NAME, conn, if_exists='replace', index=False)

            cursor = conn.cursor()
            index_cols = {'idx_numero_sencillo': 'Numero Sencillo',
                          'idx_codigos': 'Codigos',
                          'idx_proceso': 'Proceso',
                          'idx_maq': 'Maq'}
            for idx_name, col_name in index_cols.items():
                if col_name in df_excel.columns:
                    try:
                        cursor.execute(f"CREATE INDEX IF NOT EXISTS {idx_name} ON {ITEMS_TABLE_NAME}({col_name});")
                    except sqlite3.Error as e:
                        print(f"Advertencia: No se pudo crear índice en {col_name}: {e}")
            conn.commit()
            end_time = time.time()
            messagebox.showinfo("Carga Exitosa",
                                f"Datos del Excel cargados a la BD SQL '{self.db_filename}' en {end_time - start_time:.2f}s.\n"
                                "La aplicación ahora usará estos datos.")
            self._load_data_from_sql()
            return True
        except FileNotFoundError:
            messagebox.showerror("Error de Archivo", f"Archivo Excel no encontrado: {excel_filepath}")
            return False
        except Exception as e:
            messagebox.showerror("Error al Procesar Excel", f"Error al cargar Excel a SQL: {e}")
            return False
        finally:
            if conn: conn.close()

    def _get_search_condition(self, df_to_search, column_names, search_term_upper):
        final_condition = pd.Series([False] * len(df_to_search), index=df_to_search.index)
        for col_name in column_names:
            if col_name in df_to_search.columns:
                final_condition |= (df_to_search[col_name] == search_term_upper)
        return final_condition

    def find_direct_code(self, code_to_find):
        if self.df_search is None or self.df_search.empty: return None
        code_upper = code_to_find.upper()
        condition = self._get_search_condition(self.df_search, ['Numero Sencillo', 'Codigos'], code_upper)
        potential_match_indices = self.df_search[condition].index
        if potential_match_indices.empty: return None

        df_potential_matches_search = self.df_search.loc[potential_match_indices]
        df_filtered_search = df_potential_matches_search

        if not df_potential_matches_search.empty and 'Proceso' in df_potential_matches_search.columns:
            first_match_process = df_potential_matches_search.loc[df_potential_matches_search.index[0], 'Proceso']
            if first_match_process in SPECIAL_PROCESSES and 'Maq' in df_potential_matches_search.columns:
                df_filtered_search = df_potential_matches_search[df_potential_matches_search['Maq'] != MACHINE_TO_EXCLUDE]

        if df_filtered_search.empty: return None
        return self.df_original.loc[df_filtered_search.index]

    def find_process_related_codes(self, input_code_or_process):
        if self.df_search is None or self.df_search.empty: return None, None, None
        upper_input = input_code_or_process.upper()
        df_all_items_in_process_search = pd.DataFrame()
        identified_process_code = None

        if 'Proceso' in self.df_search.columns and self.df_search['Proceso'].isin([upper_input]).any():
            identified_process_code = upper_input
            df_all_items_in_process_search = self.df_search[self.df_search['Proceso'] == identified_process_code]
        elif 'Codigos' in self.df_search.columns and 'Proceso' in self.df_search.columns:
            matches_by_codigo = self.df_search[self.df_search['Codigos'] == upper_input]
            if not matches_by_codigo.empty:
                identified_process_code = matches_by_codigo.iloc[0]['Proceso']
                df_all_items_in_process_search = self.df_search[self.df_search['Proceso'] == identified_process_code]

        if df_all_items_in_process_search.empty or identified_process_code is None:
            return None, None, None

        df_filtered_for_tw_br_search = df_all_items_in_process_search
        if identified_process_code in SPECIAL_PROCESSES and 'Maq' in df_filtered_for_tw_br_search.columns:
            df_filtered_for_tw_br_search = df_all_items_in_process_search[df_all_items_in_process_search['Maq'] != MACHINE_TO_EXCLUDE]

        if df_filtered_for_tw_br_search.empty: return None, None, None
        if 'Numero Sencillo' not in df_filtered_for_tw_br_search.columns: return None, None, None

        unique_numeros_sencillos = df_filtered_for_tw_br_search['Numero Sencillo'].unique()
        if not unique_numeros_sencillos.any(): return None, None, None
        numero_sencillo_representante_search = random.choice(unique_numeros_sencillos)

        condition_representante_in_filtered_search = df_filtered_for_tw_br_search['Numero Sencillo'] == numero_sencillo_representante_search
        indices_representante_in_original = df_filtered_for_tw_br_search[condition_representante_in_filtered_search].index

        if indices_representante_in_original.empty: return None, None, None

        if not self.df_original.index.isin(indices_representante_in_original.tolist()).any():
             return None, None, None

        numero_sencillo_representante_original = self.df_original.loc[indices_representante_in_original.tolist()[0]]['Numero Sencillo']
        return identified_process_code, numero_sencillo_representante_original, self.df_original.loc[indices_representante_in_original]

    def find_code_in_process(self, process_original_indices_of_representative, code_to_find):
        if self.df_search is None or self.df_search.empty or \
           process_original_indices_of_representative is None or process_original_indices_of_representative.empty :
            return None

        valid_indices = self.df_search.index.intersection(process_original_indices_of_representative)
        if valid_indices.empty: return None

        current_process_df_search = self.df_search.loc[valid_indices]
        if current_process_df_search.empty: return None

        search_term_upper = code_to_find.upper()

        query_parts = []
        if 'Numero Sencillo' in current_process_df_search.columns:
            query_parts.append(f"`Numero Sencillo` == '{search_term_upper}'")
        if 'Codigos' in current_process_df_search.columns:
            query_parts.append(f"`Codigos` == '{search_term_upper}'")

        if not query_parts:
            # print("Error: Columnas de búsqueda ('Numero Sencillo' o 'Codigos') no encontradas en find_code_in_process.")
            return None

        query_string = " or ".join(query_parts)

        try:
            filtered_df_search = current_process_df_search.query(query_string, engine='python')

            if not filtered_df_search.empty:
                first_match_original_index = filtered_df_search.index[0]
                return self.df_original.loc[first_match_original_index]
            return None
        except Exception as e:
            print(f"Error al ejecutar query en find_code_in_process: {e}")
            # Fallback a la lógica anterior si query falla (aunque no debería si las columnas son strings)
            # Esto es más una medida de seguridad por si la query tiene problemas con ciertos caracteres.
            condition = self._get_search_condition(current_process_df_search,
                                               ['Numero Sencillo', 'Codigos'],
                                               search_term_upper)
            if len(condition) == len(current_process_df_search):
                filtered_df_search_alt = current_process_df_search[condition.values]
                if not filtered_df_search_alt.empty:
                    first_match_original_index = filtered_df_search_alt.index[0]
                    return self.df_original.loc[first_match_original_index]
            return None

# --- ChatbotApp Class ---
class ChatbotApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Asistente de Reposición Virtual")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        self.root.minsize(500, 450)

        self.db_manager = DatabaseManager()

        self.conversation_state = {}
        self.completed_repositions = []
        self.history = []
        self._setup_conversation_handlers()
        self._create_widgets()

        self._update_load_button_visibility()
        if self.db_manager.df_search is not None and not self.db_manager.df_search.empty:
            self._start_conversation()
        else:
             self._display_message("Bot", "Base de datos no cargada. Por favor, use el botón 'Cargar Excel a BD' para cargar datos.")

    def _create_widgets(self):
        self.root.configure(bg="#f0f0f0")
        self.load_button_frame = tk.Frame(self.root, bg="#f0f0f0")
        self.load_button_frame.pack(padx=10, pady=(10,0), fill=tk.X)

        self.load_excel_button = tk.Button(
            self.load_button_frame, text="Cargar Excel y (Re)Crear BD SQL",
            command=self._prompt_load_excel, font=("Segoe UI", 10, "bold"),
            bg="#FF8C00", fg="white", relief=tk.FLAT, activebackground="#FFA500",
            padx=10, pady=5, cursor="hand2"
        )

        self.chat_display = scrolledtext.ScrolledText(
            self.root, wrap=tk.WORD, state='disabled', font=("Segoe UI", 10),
            bg="#ffffff", relief=tk.SOLID, borderwidth=1, padx=5, pady=5)
        self.chat_display.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        input_frame = tk.Frame(self.root, bg="#f0f0f0")
        input_frame.pack(padx=10, pady=(0, 10), fill=tk.X)
        self.user_input = tk.Entry(input_frame, font=("Segoe UI", 11), relief=tk.SOLID, borderwidth=1)
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=7, padx=(0, 10))
        self.user_input.bind("<Return>", self._send_message_event)
        self.send_button = tk.Button(
            input_frame, text="Enviar", command=self._send_message, font=("Segoe UI", 10, "bold"),
            bg="#0078D4", fg="white", relief=tk.FLAT, activebackground="#005a9e",
            activeforeground="white", padx=15, pady=5, cursor="hand2")
        self.send_button.pack(side=tk.RIGHT)
        self.user_input.focus_set()

    def _update_load_button_visibility(self):
        if self.db_manager.df_search is not None and not self.db_manager.df_search.empty:
            if self.load_excel_button.winfo_ismapped():
                self.load_excel_button.pack_forget()
        else:
            if not self.load_excel_button.winfo_ismapped():
                 self.load_excel_button.pack(side=tk.LEFT, padx=(0,5))

    def _prompt_load_excel(self):
        excel_filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel de base de datos",
            filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*"))
        )
        if excel_filepath:
            if self.db_manager.load_excel_to_sql(excel_filepath):
                self._update_load_button_visibility()
                if not self.completed_repositions and (not self.conversation_state or self.conversation_state.get("step") is None):
                    self._start_conversation()
                else:
                    self._display_message("Bot", "Datos recargados. Puede continuar la reposición actual o iniciar una nueva respondiendo a la última pregunta, o diciendo 'Sí' para una nueva si finalizó la anterior.")
        else:
            messagebox.showinfo("Carga Cancelada", "No se seleccionó ningún archivo Excel.")

    def _setup_conversation_handlers(self):
        self.handlers = {
            CONV_STEP_ASK_REPOSITION: self._handle_ask_reposition,
            CONV_STEP_ASK_TYPE: self._handle_ask_type,
            CONV_STEP_GET_DIRECT_CODE: self._handle_get_direct_code,
            CONV_STEP_CONFIRM_DIRECT_ITEM: self._handle_confirm_direct_item,
            CONV_STEP_GET_QUANTITY_DIRECT: self._handle_get_quantity_direct,
            CONV_STEP_ASK_ANOTHER_REPOSITION: self._handle_ask_another_reposition,
            CONV_STEP_ASK_PRINT: self._handle_ask_print_generic,
            CONV_STEP_GET_PROCESS_CODE: self._handle_get_process_code,
            CONV_STEP_CONFIRM_PROCESS_ITEMS: self._handle_confirm_process_items,
            CONV_STEP_ASK_GROUP_OR_SPECIFIC: self._handle_ask_group_or_specific,
            CONV_STEP_ASK_FOR_SPECIFIC_PROCESS_CODE: self._handle_ask_for_specific_process_code,
            CONV_STEP_CONFIRM_SPECIFIC_PROCESS_ITEM: self._handle_confirm_specific_process_item,
            CONV_STEP_GET_TOTAL_GROUP_QUANTITY: self._handle_get_total_group_quantity,
            CONV_STEP_GET_SINGLE_CIRCUIT_QUANTITY: self._handle_get_single_circuit_quantity,
        }
    def _display_message(self, sender, message, color="black"):
        self.chat_display.config(state='normal')
        timestamp = datetime.now().strftime('%H:%M:%S')
        sender_name = "Asistente" if sender == "Bot" else "Usuario"
        self.chat_display.insert(tk.END, f"[{timestamp}] {sender_name}: ",
                                 "sender_bot" if sender == "Bot" else "sender_user")
        parts = message.split("**")
        for i, part in enumerate(parts):
            tag = "bold_message" if i % 2 == 1 else "normal_message"
            self.chat_display.insert(tk.END, part, tag)
        self.chat_display.insert(tk.END, "\n\n")
        self.chat_display.tag_config("sender_bot", foreground="#0078D4", font=("Segoe UI", 10, "bold"))
        self.chat_display.tag_config("sender_user", foreground="#107C10", font=("Segoe UI", 10, "bold"))
        self.chat_display.tag_config("normal_message", foreground=color, font=("Segoe UI", 10))
        self.chat_display.tag_config("bold_message", foreground=color, font=("Segoe UI", 10, "bold"))
        self.chat_display.tag_config("timestamp_tag", foreground="gray", font=("Segoe UI", 8))
        self.chat_display.yview(tk.END)
        self.chat_display.config(state='disabled')

    def _start_conversation(self):
        self.completed_repositions = []
        self.conversation_state = {}
        self._display_message("Bot",
                              "Hola, soy tu Asistente de Reposición Virtual.\n"
                              "Puedes reponer un artículo **directo** o un **proceso** completo.\n"
                              "Escribe **Sí** para comenzar o **No** para salir.")
        self._update_conversation_state(CONV_STEP_ASK_REPOSITION)

    def _send_message_event(self, event=None): self._send_message()

    def _send_message(self):
        if self.db_manager.df_search is None or self.db_manager.df_search.empty:
            self._display_message("Bot", "La base de datos no está cargada. Por favor, use el botón 'Cargar Excel a BD'.")
            return

        user_text = self.user_input.get().strip()
        if not user_text: return
        self._display_message("Tú", user_text)
        self.history.append(("user", user_text))
        self.user_input.delete(0, tk.END)
        self.user_input.focus_set()
        self._process_user_response(user_text.lower())

    def _process_user_response(self, response_lower):
        step = self.conversation_state.get("step")
        handler = self.handlers.get(step)
        if handler: handler(response_lower)
        else:
            self._display_message("Bot", "Error inesperado. Reiniciando.")
            self._start_conversation()

    def _update_conversation_state(self, next_step, **kwargs):
        context_keys_to_preserve = [
            "type", "code_searched",
            "process_code_identified", "numero_sencillo_representante", "df_proceso_representante",
            "reposition_scope", "found_item"
        ]
        preserved_data = {}

        if next_step not in [CONV_STEP_ASK_REPOSITION, CONV_STEP_ASK_TYPE,
                             CONV_STEP_GET_DIRECT_CODE, CONV_STEP_GET_PROCESS_CODE]:
            for key in context_keys_to_preserve:
                if key in self.conversation_state:
                    preserved_data[key] = self.conversation_state[key]

        self.conversation_state.clear()
        self.conversation_state["step"] = next_step

        for key, value in preserved_data.items():
            self.conversation_state[key] = value

        for key, value in kwargs.items():
            self.conversation_state[key] = value

        if next_step == CONV_STEP_GET_DIRECT_CODE:
            keys_to_remove = [k for k in self.conversation_state if k not in ["step", "type"]]
            for k in keys_to_remove: self.conversation_state.pop(k, None)
            for key_proc in ["df_proceso_representante", "process_code_identified", "numero_sencillo_representante", "reposition_scope"]:
                self.conversation_state.pop(key_proc, None)
            self.conversation_state.pop("found_item", None)
            self.conversation_state.pop("code_searched", None)

        elif next_step == CONV_STEP_GET_PROCESS_CODE:
            keys_to_remove = [k for k in self.conversation_state if k not in ["step", "type"]]
            for k in keys_to_remove: self.conversation_state.pop(k, None)
            for key_to_clear in ["found_item", "df_proceso_representante",
                                 "process_code_identified", "numero_sencillo_representante",
                                 "reposition_scope", "code_searched"]:
                self.conversation_state.pop(key_to_clear, None)


    def _handle_ask_reposition(self, r):
        if r in USER_YES_RESPONSES:
            self._display_message("Bot",
                                  "¿Reponer un **directo** o un **proceso**?\n"
                                  "(Escribe 'directo' o 'proceso')")
            self._update_conversation_state(CONV_STEP_ASK_TYPE)
        elif r in USER_NO_RESPONSES:
            if self.completed_repositions:
                 self._display_message("Bot", "¿Desea **imprimir** las reposiciones acumuladas? (**Sí**/**No**)")
                 self._update_conversation_state(CONV_STEP_ASK_PRINT)
            else:
                self._display_message("Bot", "Entendido. ¡Hasta luego!")
                self.root.after(2000, self.root.destroy)
        else: self._display_message("Bot", "Responda '**Sí**' o '**No**'.")

    def _handle_ask_type(self, r):
        if r == USER_DIRECT_RESPONSE:
            self._display_message("Bot", "Ingrese el **código del directo** (Ej: A123):")
            self._update_conversation_state(CONV_STEP_GET_DIRECT_CODE, type="directo")
        elif r == USER_PROCESS_RESPONSE:
            self._display_message("Bot",
                                  "Ingrese el **código de proceso** (Ej: P001) o un **código de producto** relacionado:")
            self._update_conversation_state(CONV_STEP_GET_PROCESS_CODE, type="proceso")
        else: self._display_message("Bot", "Responda '**directo**' o '**proceso**'.")

    def _handle_get_direct_code(self, r):
        df_found_original_all_matches = self.db_manager.find_direct_code(r)
        if df_found_original_all_matches is not None and not df_found_original_all_matches.empty:
            item_original_to_confirm = df_found_original_all_matches.iloc[0].copy()
            msg = (f"Código encontrado: **{item_original_to_confirm.get('Numero Sencillo', 'N/A')}** "
                   f"(General: **{item_original_to_confirm.get('Codigos', 'N/A')}**).\n"
                   f"Proceso: {item_original_to_confirm.get('Proceso', 'N/A')}, Maq: {item_original_to_confirm.get('Maq', 'N/A')}\n"
                   f"Tipo: {item_original_to_confirm.get('Type', 'N/A')}, Tamaño: {item_original_to_confirm.get('Size', 'N/A')}, "
                   f"Color: {item_original_to_confirm.get('Color', 'N/A')}, Largo: {item_original_to_confirm.get('Cut Length', 'N/A')}.\n"
                   f"¿Es este el artículo correcto? (**Sí**/**No**)")
            self._display_message("Bot", msg)
            self._update_conversation_state(CONV_STEP_CONFIRM_DIRECT_ITEM, found_item=item_original_to_confirm, code_searched=r.upper())
        else:
            self._display_message("Bot", f"Código '**{r.upper()}**' no encontrado o filtrado. Verifique e intente de nuevo.")
            self._update_conversation_state(CONV_STEP_GET_DIRECT_CODE)

    def _handle_confirm_direct_item(self, r):
        if r in USER_YES_RESPONSES:
            self._display_message("Bot", "¿Cuántas **piezas** desea reponer?")
            self._update_conversation_state(CONV_STEP_GET_QUANTITY_DIRECT)
        elif r in USER_NO_RESPONSES:
            self._display_message("Bot", "Entendido. Ingrese el **código del directo** correcto:")
            self._update_conversation_state(CONV_STEP_GET_DIRECT_CODE)
        else: self._display_message("Bot", "Responda '**Sí**' o '**No**'.")

    def _handle_get_quantity(self, response, next_step_after_storing, current_step_if_invalid):
        try:
            quantity = int(response)
            if quantity <= 0: raise ValueError("La cantidad debe ser positiva.")

            current_repo_data = {}
            relevant_keys = ["type", "code_searched", "found_item",
                             "process_code_identified", "numero_sencillo_representante",
                             "df_proceso_representante", "reposition_scope"]
            for key in relevant_keys:
                if key in self.conversation_state:
                    value = self.conversation_state[key]
                    if isinstance(value, (pd.DataFrame, pd.Series)):
                        current_repo_data[key] = value.copy(deep=True)
                    elif isinstance(value, (list, dict)):
                        current_repo_data[key] = copy.deepcopy(value)
                    else:
                        current_repo_data[key] = value

            current_repo_data["quantity"] = quantity
            self.completed_repositions.append(current_repo_data)

            item_id_for_msg = current_repo_data.get('code_searched') or \
                              current_repo_data.get('numero_sencillo_representante') or \
                              (current_repo_data.get('found_item', {}).get("Numero Sencillo") if isinstance(current_repo_data.get('found_item'), (pd.Series, dict)) else None) or \
                              "Ítem/Proceso"

            self._display_message("Bot", f"Reposición para '{item_id_for_msg}' (Cantidad: {quantity}) añadida a la lista.")
            return True, next_step_after_storing
        except ValueError:
            self._display_message("Bot", "Cantidad inválida. Ingrese un **número entero positivo**.")
            return False, current_step_if_invalid
        except Exception as e:
            print(f"Error inesperado en _handle_get_quantity: {e}")
            self._display_message("Bot", "Ocurrió un error al procesar la cantidad.")
            return False, current_step_if_invalid


    def _handle_get_quantity_direct(self, r):
        is_valid, next_step = self._handle_get_quantity(r, CONV_STEP_ASK_ANOTHER_REPOSITION, CONV_STEP_GET_QUANTITY_DIRECT)
        if is_valid:
            self._display_message("Bot", "¿Realizar **otra reposición**? (**Sí** para continuar / **No** para finalizar e imprimir si desea)")
        self._update_conversation_state(next_step)

    def _handle_ask_another_reposition(self, r):
        if r in USER_YES_RESPONSES:
            self._display_message("Bot",
                                  "¿Siguiente reposición será **directo** o **proceso**?\n"
                                  "(Escribe 'directo' o 'proceso')")
            self._update_conversation_state(CONV_STEP_ASK_TYPE)
        elif r in USER_NO_RESPONSES:
            if self.completed_repositions:
                self._display_message("Bot", "¿**Imprimir** todas las reposiciones acumuladas? (**Sí**/**No**)")
                self._update_conversation_state(CONV_STEP_ASK_PRINT)
            else:
                self._display_message("Bot", "No hay reposiciones para imprimir. ¡Hasta luego!")
                self.root.after(2000, self.root.destroy)
        else: self._display_message("Bot", "Responda '**Sí**' o '**No**'.")

    def _handle_get_process_code(self, r):
        proc_code, num_sencillo_rep, df_rep_original = self.db_manager.find_process_related_codes(r)
        if df_rep_original is not None and not df_rep_original.empty and num_sencillo_rep:
            item_representante_display = df_rep_original.iloc[0]
            msg = (f"Proceso identificado: **{proc_code}**.\n"
                   f"Número de Parte Representante seleccionado: **{num_sencillo_rep}**.\n"
                   f"(General: {item_representante_display.get('Codigos', 'N/A')}, Maq: {item_representante_display.get('Maq', 'N/A')})\n"
                   f"¿Es este el proceso/representante que desea utilizar? (**Sí**/**No**)")
            self._display_message("Bot", msg)
            self._update_conversation_state(CONV_STEP_CONFIRM_PROCESS_ITEMS,
                                          df_proceso_representante=df_rep_original.copy(deep=True),
                                          process_code_identified=proc_code,
                                          numero_sencillo_representante=num_sencillo_rep,
                                          code_searched=r.upper())
        else:
            self._display_message("Bot", f"Proceso/código '**{r.upper()}**' no encontrado o sin items válidos. Verifique.")
            self._update_conversation_state(CONV_STEP_GET_PROCESS_CODE)

    def _handle_confirm_process_items(self, r):
        if r in USER_YES_RESPONSES:
            self._display_message("Bot",
                                  "¿Reponer el **grupo completo** o un **circuito específico**?\n"
                                  "(Escriba 'grupo' o 'especifico')")
            self._update_conversation_state(CONV_STEP_ASK_GROUP_OR_SPECIFIC)
        elif r in USER_NO_RESPONSES:
            self._display_message("Bot", "Entendido. Ingrese el **código de proceso/producto** correcto:")
            self._update_conversation_state(CONV_STEP_GET_PROCESS_CODE)
        else: self._display_message("Bot", "Responda '**Sí**' o '**No**'.")

    def _handle_ask_group_or_specific(self, r):
        if r == USER_GROUP_RESPONSE:
            self._display_message("Bot", "Cantidad para **cada código general** del grupo:")
            self._update_conversation_state(CONV_STEP_GET_TOTAL_GROUP_QUANTITY, reposition_scope="full_group")
        elif r == USER_SPECIFIC_RESPONSE or r == USER_SPECIFIC_RESPONSE_ACCENT:
            self._display_message("Bot", "Ingrese **código del circuito específico** (Codigo General):")
            self._update_conversation_state(CONV_STEP_ASK_FOR_SPECIFIC_PROCESS_CODE, reposition_scope="single_circuit")
        else: self._display_message("Bot", "Responda '**Grupo**' o '**Especifico**'.")

    def _handle_ask_for_specific_process_code(self, r):
        df_representante_original = self.conversation_state.get("df_proceso_representante")

        if df_representante_original is None or df_representante_original.empty:
            self._display_message("Bot", "Error: No se encontró información del proceso representante. Reiniciando.")
            self._start_conversation()
            return

        item_original_specific_circuit = self.db_manager.find_code_in_process(df_representante_original.index, r)

        if item_original_specific_circuit is not None and not item_original_specific_circuit.empty:
            msg = (f"Circuito/Código General: **{item_original_specific_circuit.get('Codigos', 'N/A')}**\n"
                   f"(Del Numero Sencillo: **{item_original_specific_circuit.get('Numero Sencillo', 'N/A')}**)\n"
                   f"¿Correcto? (**Sí**/**No**)")
            self._display_message("Bot", msg)
            self._update_conversation_state(CONV_STEP_CONFIRM_SPECIFIC_PROCESS_ITEM,
                                          found_item=item_original_specific_circuit.copy(deep=True)
                                          )
        else:
            self._display_message("Bot", f"Código General '**{r.upper()}**' no encontrado para el representante actual. Intente de nuevo.")
            self._update_conversation_state(CONV_STEP_ASK_FOR_SPECIFIC_PROCESS_CODE)


    def _handle_confirm_specific_process_item(self, r):
        if r in USER_YES_RESPONSES:
            self._display_message("Bot", "¿Cuántas **piezas** para este circuito?")
            self._update_conversation_state(CONV_STEP_GET_SINGLE_CIRCUIT_QUANTITY)
        elif r in USER_NO_RESPONSES:
            self._display_message("Bot", "Entendido. Ingrese el **código general específico** correcto:")
            self._update_conversation_state(CONV_STEP_ASK_FOR_SPECIFIC_PROCESS_CODE)
        else: self._display_message("Bot", "Responda '**Sí**' o '**No**'.")

    def _handle_get_total_group_quantity(self, r):
        is_valid, next_step = self._handle_get_quantity(r, CONV_STEP_ASK_ANOTHER_REPOSITION, CONV_STEP_GET_TOTAL_GROUP_QUANTITY)
        if is_valid:
            self._display_message("Bot", "¿Realizar **otra reposición**? (Sí para continuar / No para finalizar e imprimir si desea)")
        self._update_conversation_state(next_step)

    def _handle_get_single_circuit_quantity(self, r):
        is_valid, next_step = self._handle_get_quantity(r, CONV_STEP_ASK_ANOTHER_REPOSITION, CONV_STEP_GET_SINGLE_CIRCUIT_QUANTITY)
        if is_valid:
            self._display_message("Bot", "¿Realizar **otra reposición**? (Sí para continuar / No para finalizar e imprimir si desea)")
        self._update_conversation_state(next_step)

    def _handle_ask_print_generic(self, response):
        if not self.completed_repositions:
            self._display_message("Bot", "No hay reposiciones para imprimir. ¡Hasta luego!")
            self.root.after(2000, self.root.destroy)
            return

        if response in USER_YES_RESPONSES:
            self._print_reposition_info()
            # self._display_message("Bot", "Reposiciones acumuladas enviadas a impresión. ¡Gracias!") # Mensaje ya en _print_reposition_info
            # self.root.after(2500, self.root.destroy) # Cierre ya en _print_reposition_info
        elif response in USER_NO_RESPONSES:
            self._display_message("Bot", "No se imprimirán las reposiciones. ¡Gracias!")
            self.root.after(2500, self.root.destroy)
        else:
            self._display_message("Bot", "Por favor, responda '**Sí**' o '**No**'.")
            self._update_conversation_state(CONV_STEP_ASK_PRINT) # Re-preguntar si la respuesta no es válida


    def _generate_report_data_for_excel(self):
        """Prepara los datos para el reporte Excel como una lista de diccionarios."""
        report_data_list = []
        if not self.completed_repositions:
            return [{"INFO": "No hay reposiciones registradas en esta sesión."}]

        report_data_list.append({"INFO": f"--- INICIO REPORTE GENERAL DE REPOSICIONES ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---"})

        for i, state_data in enumerate(self.completed_repositions):
            report_data_list.append({"INFO": f"--- REPOSICIÓN #{i+1} ---"})
            repo_type = state_data.get("type", "N/A")
            qty = state_data.get("quantity", 0)

            report_data_list.append({"INFO": f"Tipo de Reposición: {repo_type.upper()}"})
            report_data_list.append({"INFO": f"Código/Proceso Solicitado: {state_data.get('code_searched', 'N/A')}"})
            if repo_type == "proceso":
                report_data_list.append({"INFO": f"Proceso Identificado en BD: {state_data.get('process_code_identified', 'N/A')}"})
                report_data_list.append({"INFO": f"Numero Sencillo Representante: {state_data.get('numero_sencillo_representante', 'N/A')}"})

            scope = state_data.get("reposition_scope", "N/A")
            if repo_type == "proceso" and scope == "full_group":
                 report_data_list.append({"INFO": f"Alcance: Grupo Completo"})

            # Añadir línea de cantidad general
            if repo_type != "N/A" and qty > 0 :
                if repo_type == "directo" or (repo_type == "proceso" and scope == "single_circuit"):
                     report_data_list.append({"INFO": f"Cantidad a Reponer (ítem/circuito): {qty} piezas"})
                elif repo_type == "proceso" and scope == "full_group":
                     report_data_list.append({"INFO": f"Cantidad (por cada código general del grupo): {qty} piezas"})

            # Encabezados de la tabla de ítems
            report_data_list.append({
                'NUMERO DE PARTE': 'NUMERO DE PARTE', 'CODIGO GENERAL': 'CODIGO GENERAL',
                'CIRCUITO A': 'CIRCUITO A', 'CIRCUITO B': 'CIRCUITO B',
                'PROCESO EN BD': 'PROCESO EN BD', 'CANTIDAD': 'CANTIDAD',
                'GRUPO(SI/NO)': 'GRUPO(SI/NO)', 'PLANTA': 'PLANTA'
            })

            items_in_current_repo = []
            if repo_type == "directo":
                item_original = state_data.get("found_item")
                if item_original is not None and not item_original.empty:
                    items_in_current_repo.append({"data": item_original, "qty": qty, "is_group": "NO"})
            elif repo_type == "proceso":
                if scope == "full_group":
                    df_representante_original = state_data.get("df_proceso_representante", pd.DataFrame())
                    if not df_representante_original.empty:
                        unique_codigos = df_representante_original['Codigos'].drop_duplicates().tolist()
                        for codigo_gen in unique_codigos:
                            first_row = df_representante_original[df_representante_original['Codigos'] == codigo_gen].iloc[0]
                            items_in_current_repo.append({"data": first_row, "qty": qty, "is_group": "SI"})
                elif scope == "single_circuit":
                    item_original = state_data.get("found_item")
                    if item_original is not None and not item_original.empty:
                        items_in_current_repo.append({"data": item_original, "qty": qty, "is_group": "NO"})

            if not items_in_current_repo:
                report_data_list.append({"INFO": "No hay artículos específicos para detallar para esta reposición."})
            else:
                for item_detail in items_in_current_repo:
                    d = item_detail["data"]
                    report_data_list.append({
                        'NUMERO DE PARTE': d.get('Numero Sencillo', 'N/A'),
                        'CODIGO GENERAL': d.get('Codigos', 'N/A'),
                        'CIRCUITO A': d.get('Cod A', 'N/A'),
                        'CIRCUITO B': d.get('Cod B', 'N/A'),
                        'PROCESO EN BD': d.get('Proceso', 'N/A'),
                        'CANTIDAD': item_detail['qty'],
                        'GRUPO(SI/NO)': item_detail['is_group'],
                        'PLANTA': d.get('Planta', 'N/A')
                    })
            report_data_list.append({"INFO": "--- FIN REPOSICIÓN ---"}) # Separador

        report_data_list.append({"INFO": f"--- FIN REPORTE GENERAL ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---"})
        return report_data_list


    def _print_reposition_info(self):
        report_data = self._generate_report_data_for_excel()

        if not report_data or (len(report_data) == 1 and "No hay reposiciones" in report_data[0].get("INFO","")) :
             self._display_message("Bot", "No hay información de reposición para generar el archivo.")
             return

        # Preparar datos para DataFrame, excluyendo filas de "INFO" para la tabla principal
        table_data = [row for row in report_data if "INFO" not in row]
        df_report = pd.DataFrame(table_data)

        # Crear el archivo Excel
        excel_filename = f"{OUTPUT_FILE_PREFIX}{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        full_excel_path = os.path.abspath(excel_filename)

        try:
            with pd.ExcelWriter(full_excel_path, engine='openpyxl') as writer:
                # Escribir la información general (filas con "INFO")
                # Esto se podría mejorar para formatearlo mejor en Excel, pero por ahora simple.
                info_rows = [list(row.values())[0] for row in report_data if "INFO" in row]
                df_info = pd.DataFrame(info_rows, columns=["Detalles de la Sesión de Reposición"])
                df_info.to_excel(writer, sheet_name=OUTPUT_EXCEL_SHEET_NAME, index=False, startrow=0)

                # Escribir la tabla de datos de reposición debajo de la información
                start_row_for_table = len(df_info) + 2 # Dejar una fila en blanco
                df_report.to_excel(writer, sheet_name=OUTPUT_EXCEL_SHEET_NAME, index=False, startrow=start_row_for_table)

            messagebox.showinfo("Reporte Excel Generado", f"Reporte guardado en:\n'{full_excel_path}'")
            self.history.append(("bot", f"Reporte Excel guardado en {full_excel_path}"))

            # Mantener la impresión en consola para depuración
            console_report_lines = [str(item) for item in report_data] # Convertir dicts a string para join
            print(("\n" + "="*80 + "\n--- REPORTE CONSOLA (datos para Excel) ---\n" + "\n".join(console_report_lines) + "\n" + "="*80 + "\n"))

            self.root.after(2500, self.root.destroy) # Cerrar después de imprimir

        except Exception as e:
            messagebox.showerror("Error al Guardar Excel", f"No se pudo guardar el archivo Excel:\n{e}")
            self.history.append(("bot", f"Error al guardar reporte Excel: {e}"))
            # Imprimir en consola si falla el Excel
            console_report_lines = [str(item) for item in report_data]
            print(("\n" + "="*80 + "\n--- REPORTE CONSOLA (FALLÓ EXCEL) ---\n" + "\n".join(console_report_lines) + "\n" + "="*80 + "\n"))


# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ChatbotApp(root)
    if root.winfo_exists():
        app._update_load_button_visibility()
        root.mainloop()

[end of chatbot_reposicion_mejorado.py]

[end of chatbot_reposicion_mejorado.py]

[end of chatbot_reposicion_mejorado.py]

[end of chatbot_reposicion_mejorado.py]
