import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import os

# Configuración de apariencia
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ChecadorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema Integral de Accesos y Horarios - UTN")
        self.geometry("750x600")
        
        # Variables para almacenar rutas
        self.ruta_csv_checador = None
        self.ruta_excel_horarios = None
        
        # Variable para almacenar el DataFrame de los horarios cargados
        self.df_horarios = pd.DataFrame()

        # --- Sistema de Pestañas ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(padx=20, pady=20, fill="both", expand=True)

        # Invertimos el orden lógico: Primero se cargan los horarios, luego se procesa el CSV
        self.tab_horarios = self.tabview.add("1. Cargar Horarios (Excel)")
        self.tab_procesamiento = self.tabview.add("2. Procesar Checadas (CSV)")

        # ==========================================
        # PESTAÑA 1: GESTIÓN DE HORARIOS (EXCEL)
        # ==========================================
        self.label_titulo_horarios = ctk.CTkLabel(self.tab_horarios, text="Base de Datos de Horarios del Personal", font=("Arial", 18, "bold"))
        self.label_titulo_horarios.pack(pady=(15, 5))

        self.label_desc_horarios = ctk.CTkLabel(self.tab_horarios, text="Sube la plantilla de Excel (Plantilla_Horarios_UTN.xlsx) con los horarios autorizados.", font=("Arial", 12))
        self.label_desc_horarios.pack(pady=(0, 15))

        # Botón para cargar Excel
        self.btn_cargar_horarios = ctk.CTkButton(self.tab_horarios, text="Subir Plantilla de Excel", command=self.cargar_excel_horarios, width=200, height=40, fg_color="#1F4E78", hover_color="#103659")
        self.btn_cargar_horarios.pack(pady=5)

        self.label_excel_seleccionado = ctk.CTkLabel(self.tab_horarios, text="Ningún archivo seleccionado", text_color="gray")
        self.label_excel_seleccionado.pack(pady=(0, 10))

        # Tabla visual para mostrar los horarios cargados
        self.frame_tabla = ctk.CTkFrame(self.tab_horarios)
        self.frame_tabla.pack(padx=10, pady=10, fill="both", expand=True)

        columnas = ("Nombre", "Tipo", "Lun_Ent", "Mar_Ent", "Mie_Ent", "Jue_Ent", "Vie_Ent")
        self.tree = ttk.Treeview(self.frame_tabla, columns=columnas, show="headings", height=8)
        
        # Encabezados de la tabla visual
        self.tree.heading("Nombre", text="Nombre Completo")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.heading("Lun_Ent", text="Lunes (E)")
        self.tree.heading("Mar_Ent", text="Martes (E)")
        self.tree.heading("Mie_Ent", text="Miércoles (E)")
        self.tree.heading("Jue_Ent", text="Jueves (E)")
        self.tree.heading("Vie_Ent", text="Viernes (E)")
        
        self.tree.column("Nombre", width=180)
        self.tree.column("Tipo", width=120)
        for col in columnas[2:]:
            self.tree.column(col, width=80, anchor="center")
        
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # ==========================================
        # PESTAÑA 2: PROCESAMIENTO DE CSV
        # ==========================================
        self.label_titulo_csv = ctk.CTkLabel(self.tab_procesamiento, text="Limpieza y Cruce de Accesos", font=("Arial", 18, "bold"))
        self.label_titulo_csv.pack(pady=(20, 10))

        self.label_desc_csv = ctk.CTkLabel(self.tab_procesamiento, text="Sube el archivo crudo del checador (.csv) para cruzarlo con los horarios autorizados.", font=("Arial", 12))
        self.label_desc_csv.pack(pady=(0, 15))

        self.btn_cargar_csv = ctk.CTkButton(self.tab_procesamiento, text="Seleccionar Archivo CSV", command=self.cargar_csv, width=200, height=40)
        self.btn_cargar_csv.pack(pady=10)

        self.label_csv_seleccionado = ctk.CTkLabel(self.tab_procesamiento, text="Ningún archivo seleccionado", text_color="gray")
        self.label_csv_seleccionado.pack(pady=(0, 10))

        # Botón para limpiar, cruzar y generar Excel
        self.btn_procesar = ctk.CTkButton(self.tab_procesamiento, text="Generar Reporte de Incidencias", command=self.procesar_y_cruzar, width=250, height=45, fg_color="green", hover_color="darkgreen", state="disabled")
        self.btn_procesar.pack(pady=20)


    # --- Funciones de Lógica ---

    def cargar_excel_horarios(self):
        archivo_seleccionado = filedialog.askopenfilename(
            title="Selecciona la plantilla de horarios en Excel",
            filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
        )

        if archivo_seleccionado:
            self.ruta_excel_horarios = archivo_seleccionado
            nombre_archivo = os.path.basename(self.ruta_excel_horarios)
            self.label_excel_seleccionado.configure(text=f"Base cargada: {nombre_archivo}", text_color="green")
            
            try:
                # Usamos pandas para leer el excel (saltamos la primera fila si hay algún título extraño, 
                # pero en nuestra plantilla la fila 1 es el encabezado real)
                self.df_horarios = pd.read_excel(self.ruta_excel_horarios)
                
                # Limpiar tabla visual previa
                for item in self.tree.get_children():
                    self.tree.delete(item)

                # Rellenar la tabla visual (Treeview) con los datos de pandas
                for index, row in self.df_horarios.iterrows():
                    # Formatear celdas NaN (vacías) como un string vacío para que no salga "nan"
                    row = row.fillna("")
                    
                    # Extraer los datos específicos para la previsualización
                    # Asumiendo los encabezados exactos de nuestra plantilla
                    try:
                        valores = (
                            str(row["Nombre Completo"]),
                            str(row["Tipo de Empleado"]),
                            str(row["Lunes Entrada"])[:5] if row["Lunes Entrada"] else "-",
                            str(row["Martes Entrada"])[:5] if row["Martes Entrada"] else "-",
                            str(row["Miércoles Entrada"])[:5] if row["Miércoles Entrada"] else "-",
                            str(row["Jueves Entrada"])[:5] if row["Jueves Entrada"] else "-",
                            str(row["Viernes Entrada"])[:5] if row["Viernes Entrada"] else "-"
                        )
                        self.tree.insert("", "end", values=valores)
                    except KeyError:
                        messagebox.showwarning("Error de Formato", "El Excel no tiene los encabezados correctos de la plantilla.")
                        return

            except Exception as e:
                messagebox.showerror("Error al leer Excel", f"No se pudo cargar el archivo:\n{e}")

    def cargar_csv(self):
        archivo_seleccionado = filedialog.askopenfilename(
            title="Selecciona el archivo del checador",
            filetypes=(("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*"))
        )
        if archivo_seleccionado:
            self.ruta_csv_checador = archivo_seleccionado
            self.label_csv_seleccionado.configure(text=f"Archivo: {os.path.basename(self.ruta_csv_checador)}", text_color="black")
            
            # Solo habilitamos el botón final si AMBOS archivos están cargados
            if not self.df_horarios.empty:
                self.btn_procesar.configure(state="normal")
            else:
                messagebox.showinfo("Aviso", "Asegúrate de cargar primero la plantilla de horarios en la Pestaña 1.")

    def procesar_y_cruzar(self):
        if not self.ruta_csv_checador or self.df_horarios.empty:
            messagebox.showwarning("Faltan datos", "Debes cargar los horarios (Pestaña 1) y el CSV (Pestaña 2).")
            return

        try:
            # --- 1. LEER Y LIMPIAR EL CSV CRUDO ---
            df_csv = pd.read_csv(self.ruta_csv_checador, sep=';', header=None, skiprows=1, encoding='latin1', on_bad_lines='skip')
            
            df_temp = pd.DataFrame()
            df_temp[['Fecha', 'Hora']] = df_csv[0].str.split(' ', expand=True)
            df_temp['Nombre Completo'] = df_csv[3].str.strip()

            # Convertir a formato de reloj real para poder calcular minutos
            df_temp['FechaHora_Real'] = pd.to_datetime(df_temp['Fecha'] + ' ' + df_temp['Hora'], format='%d/%m/%Y %H:%M:%S')

            # Agrupar por persona y día (Primer registro = Entrada, Último = Salida)
            agrupado = df_temp.groupby(['Nombre Completo', 'Fecha'])
            df_limpio = agrupado.agg(
                Entrada_Real=('FechaHora_Real', 'min'),
                Salida_Real=('FechaHora_Real', 'max')
            ).reset_index()

            # Detectar si alguien solo pasó el dedo 1 vez (Olvidó checar salida)
            df_limpio.loc[df_limpio['Entrada_Real'] == df_limpio['Salida_Real'], 'Salida_Real'] = pd.NaT

            # --- 2. TRADUCIR FECHAS A DÍAS (Lunes, Martes...) ---
            dias_espanol = {
                'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
                'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
            }
            # Extraemos el día exacto de la semana
            fechas_dt = pd.to_datetime(df_limpio['Fecha'], format='%d/%m/%Y')
            df_limpio['Dia'] = fechas_dt.dt.day_name().map(dias_espanol)

            # --- 3. EL CRUCE MAESTRO CON LOS HORARIOS DE EXCEL ---
            # Cruzamos los datos basándonos en el 'Nombre Completo'
            df_reporte = pd.merge(df_limpio, self.df_horarios, on="Nombre Completo", how="left")

            # --- 4. LÓGICA MATEMÁTICA DE INCIDENCIAS ---
            resultados = []

            for index, row in df_reporte.iterrows():
                dia = row['Dia']
                ent_real_dt = row['Entrada_Real']
                sal_real_dt = row['Salida_Real']
                
                # Valores por si no hay turno asignado ese día
                ent_ideal_str = "Sin Turno"
                sal_ideal_str = "Sin Turno"
                estatus_ent = "-"
                estatus_sal = "-"
                tipo_empleado = str(row.get('Tipo de Empleado', 'No Registrado en Excel'))

                # Si ocurrió de Lunes a Viernes, buscamos su horario exacto
                if dia in ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']:
                    col_ent = f"{dia} Entrada"
                    col_sal = f"{dia} Salida"
                    
                    # Verificamos si la persona tiene horas asignadas en la plantilla para ese día
                    if col_ent in df_reporte.columns and not pd.isna(row[col_ent]) and str(row[col_ent]).strip() != "":
                        ent_ideal_str = str(row[col_ent]).strip()[:5] # Extraer HH:MM
                        sal_ideal_str = str(row[col_sal]).strip()[:5]
                        
                        # Convertimos para usar lógica > o <
                        hora_ideal_ent = pd.to_datetime(ent_ideal_str, format="%H:%M").time()
                        hora_ideal_sal = pd.to_datetime(sal_ideal_str, format="%H:%M").time()
                        hora_real_ent = ent_real_dt.time()
                        
                        # A. Calificar Entrada
                        if hora_real_ent > hora_ideal_ent:
                            estatus_ent = "RETARDO"
                        else:
                            estatus_ent = "A Tiempo"
                            
                        # B. Calificar Salida
                        if pd.isna(sal_real_dt):
                            estatus_sal = "FALTA SALIDA"
                        else:
                            hora_real_sal = sal_real_dt.time()
                            if hora_real_sal < hora_ideal_sal:
                                estatus_sal = "SALIDA ANTICIPADA"
                            else:
                                estatus_sal = "A Tiempo"
                
                # Preparamos las horas reales para que se vean bonitas en Excel
                str_ent_real = ent_real_dt.strftime('%H:%M:%S')
                str_sal_real = sal_real_dt.strftime('%H:%M:%S') if not pd.isna(sal_real_dt) else "Sin Registro"
                
                # Manejar valores "nan" para el tipo de empleado
                if tipo_empleado == "nan": tipo_empleado = "No Registrado en Excel"

                # Guardamos la fila evaluada
                resultados.append({
                    "Nombre Completo": row['Nombre Completo'],
                    "Tipo": tipo_empleado,
                    "Fecha": row['Fecha'],
                    "Día": dia,
                    "Entrada Autorizada": ent_ideal_str,
                    "Entrada Real": str_ent_real,
                    "Incidencia Entrada": estatus_ent,
                    "Salida Autorizada": sal_ideal_str,
                    "Salida Real": str_sal_real,
                    "Incidencia Salida": estatus_sal
                })

            # Generamos el reporte final y lo ordenamos por Fecha
            df_final = pd.DataFrame(resultados)
            df_final = df_final.sort_values(by=['Fecha', 'Nombre Completo'])

            # --- 5. EXPORTAR REPORTE A EXCEL ---
            ruta_guardado = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="Reporte_Incidencias_UTN.xlsx",
                title="Guardar Reporte Final",
                filetypes=(("Archivos Excel", "*.xlsx"),)
            )

            if ruta_guardado:
                df_final.to_excel(ruta_guardado, index=False, engine='openpyxl')
                messagebox.showinfo("Éxito", f"Reporte generado exitosamente en:\n{ruta_guardado}")
                
                # Reiniciar UI para un nuevo archivo
                self.ruta_csv_checador = None
                self.label_csv_seleccionado.configure(text="Ningún archivo seleccionado", text_color="gray")
                self.btn_procesar.configure(state="disabled")

        except Exception as e:
            messagebox.showerror("Error de Análisis", f"Hubo un problema al cruzar los datos:\n\n{str(e)}")

if __name__ == "__main__":
    app = ChecadorApp()
    app.mainloop()