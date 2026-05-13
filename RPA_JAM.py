import os
import shutil
import pyautogui
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import warnings
import time 

# ==========================================
# CLASE CONSOLA DE PROGRESO (GUI)
# ==========================================
# ==========================================
# CLASE CONSOLA DE PROGRESO (GUI)
# ==========================================
class VentanaConsola:
    # CONSTRUCTOR: Inicializa la ventana de consola con barra de progreso y area de logs
    def __init__(self, titulo="Ejecucion RPA en curso"):
        self.root = tk.Tk()  # Crear ventana principal Tk
        self.root.title(titulo)  # Establecer titulo de ventana
        self.root.geometry("750x500")  # Tamaño inicial (ancho x alto)
        self.root.configure(bg="#f4f4f4")  # Color fondo gris claro
        
        # CENTRAR LA VENTANA EN LA PANTALLA
        window_width, window_height = 750, 500  # Dimensiones
        screen_width = self.root.winfo_screenwidth()  # Ancho pantalla total
        screen_height = self.root.winfo_screenheight()  # Alto pantalla total
        position_top = int(screen_height/2 - window_height/2)  # Centro vertical
        position_right = int(screen_width/2 - window_width/2)  # Centro horizontal
        self.root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
        
        # ETIQUETA DE ESTADO: Muestra mensaje del proceso actual
        self.lbl_estado = tk.Label(self.root, text="Iniciando proceso...", font=("Segoe UI", 10, "bold"), bg="#f4f4f4", fg="#333333")
        self.lbl_estado.pack(pady=(15, 5))  # Margen 15 arriba, 5 abajo
        
        # BARRA DE PROGRESO: Indicador visual del avance 0-100%
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=700, mode="determinate")
        self.progress.pack(pady=5)  # Margen de 5 pixeles
        
        # FRAME PARA AREA DE TEXTO Y SCROLL
        frame_text = tk.Frame(self.root)  # Marco contenedor
        frame_text.pack(pady=10, padx=20, fill="both", expand=True)  # Ocupa espacio disponible
        
        # BARRA DE DESPLAZAMIENTO VERTICAL
        self.scrollbar = tk.Scrollbar(frame_text)  # Barra vertical
        self.scrollbar.pack(side="right", fill="y")  # Colocar a derecha
        
        # VENTANA DE LOGS: Area de texto para mostrar proceso en tiempo real
        # Fondo oscuro #1e1e1e, texto verde #00ff00 para efecto terminal
        self.text_log = tk.Text(frame_text, height=20, width=80, font=("Consolas", 9), yscrollcommand=self.scrollbar.set, bg="#1e1e1e", fg="#00ff00")
        self.text_log.pack(side="left", fill="both", expand=True)  # Ocupa el espacio principal
        self.scrollbar.config(command=self.text_log.yview)  # Vincular barra a ventana de texto
        
        self.root.update()  # Actualizar pantalla

    # METODO LOG: Agregar mensaje al registro de logs
    def log(self, mensaje):  # mensaje = texto a registrar
        self.text_log.insert(tk.END, mensaje + "\n")  # Insertar al final con salto de linea
        self.text_log.see(tk.END)  # Hacer scroll automatico al final
        self.root.update()  # Refrescar pantalla

    # METODO PROGRESO: Actualizar barra de progreso y etiqueta de estado
    def set_progreso(self, actual, total, mensaje_estado=""):  # actual = items procesados, total = items totales
        if total > 0:  # Validar que no sea division por cero
            porcentaje = (actual / total) * 100  # Calcular porcentaje
            self.progress["value"] = porcentaje  # Actualizar barra visual
        if mensaje_estado:  # Si se proporciono mensaje
            self.lbl_estado.config(text=mensaje_estado)  # Actualizar etiqueta de estado
        self.root.update()  # Refrescar pantalla
        
    # METODO CERRAR: Destruir la ventana
    def cerrar(self):  # Cerrar ventana de consola
        self.root.destroy()  # Destruir widget

# ==========================================
# FUNCIONES AUXILIARES DE BUSQUEDA
# ==========================================

# FUNCION: Buscar archivos por nombre en carpeta y subcarpetas
def buscar_archivos(carpeta_base, nombres_buscar):  # carpeta_base = donde empezar, nombres_buscar = lista de nombres
    encontrados = []  # Lista para almacenar resultados
    for raiz, _, archivos in os.walk(carpeta_base):  # Recorrer arbol de directorios
        for archivo in archivos:  # Para cada archivo
            if archivo in nombres_buscar:  # Si nombre esta en lista de busqueda
                encontrados.append(os.path.normpath(os.path.join(raiz, archivo)))  # Agregar ruta completa
    return encontrados  # Retornar lista de rutas encontradas

# FUNCION: Encontrar archivo Excel especifico de tablas EBSS
def encontrar_excel_ebss(carpeta_base, codigo, empresa):  # Busca el Excel de datos EBSS
    nombre_excel = f"{codigo} TABLAS EBSS {empresa}.xlsx"  # Nombre esperado del archivo
    # ESTRATEGIA 1: Busqueda hacia abajo desde carpeta_base
    for raiz, _, archivos in os.walk(carpeta_base):  # Recorrer desde carpeta_base
        if nombre_excel in archivos:  # Si encontramos el archivo
            return os.path.normpath(os.path.join(raiz, nombre_excel))  # Retornar ruta completa
    
    # ESTRATEGIA 2: Subir hacia arriba hasta encontrar carpeta proyecto
    ruta = carpeta_base  # Punto de inicio
    nombre_proyecto = f"{codigo} {empresa}"  # Nombre esperado de carpeta padre
    # Subir en jerarquia hasta encontrar carpeta del proyecto o llegar a raiz
    while os.path.basename(ruta) != nombre_proyecto and os.path.dirname(ruta) != ruta:
        ruta = os.path.dirname(ruta)  # Ir un nivel arriba
    
    # ESTRATEGIA 3: Buscar desde carpeta padre encontrada
    for raiz, _, archivos in os.walk(ruta):  # Recorrer desde carpeta padre
        if nombre_excel in archivos:  # Si encontramos
            return os.path.normpath(os.path.join(raiz, nombre_excel))  # Retornar ruta
            
    return None  # No encontrado

# ==========================================
# FASE 1: CLONAR Y RENOMBRAR
# ==========================================
def fase1_crear_entorno(ruta_plantilla, consola):
    if not ruta_plantilla or not os.path.exists(ruta_plantilla):
        consola.log("[ERROR] Ruta de plantilla no valida.")
        return False

    codigo_proyecto = pyautogui.prompt("Introduce el CODIGO del proyecto:", "Datos del Proyecto")
    if not codigo_proyecto: return None 
        
    nombre_empresa = pyautogui.prompt("Introduce el NOMBRE DE LA EMPRESA:", "Datos del Proyecto")
    if not nombre_empresa: return None

    directorio_padre = os.path.normpath(os.path.dirname(ruta_plantilla))
    nombre_nueva_carpeta = f"{codigo_proyecto} {nombre_empresa}"
    carpeta_destino = os.path.join(directorio_padre, nombre_nueva_carpeta)

    consola.set_progreso(10, 100, f"Creando carpeta: {nombre_nueva_carpeta}...")
    consola.log(f"==========================================")
    consola.log(f"FASE 1: CLONACION Y RENOMBRADO")
    consola.log(f"==========================================")
    consola.log(f"Ruta origen: {ruta_plantilla}")
    consola.log(f"Ruta destino: {carpeta_destino}\n")

    try:
        shutil.copytree(ruta_plantilla, carpeta_destino)
        consola.set_progreso(50, 100, "Carpeta base clonada. Renombrando directorios y archivos...")

        total_archivos = sum([len(files) for r, d, files in os.walk(carpeta_destino)])
        procesados = 0

        for raiz, carpetas, archivos in os.walk(carpeta_destino, topdown=False):
            for nombre_archivo in archivos:
                procesados += 1
                if "XXXX" in nombre_archivo or "[[EMPRESA]]" in nombre_archivo:
                    nuevo_nombre = nombre_archivo.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    os.rename(os.path.join(raiz, nombre_archivo), os.path.join(raiz, nuevo_nombre))
                    consola.log(f"   [RENOMBRADO] Archivo: {nuevo_nombre}")
                if procesados % 5 == 0:
                    consola.set_progreso(50 + (procesados/total_archivos)*40, 100)
            
            for nombre_carpeta in carpetas:
                if "XXXX" in nombre_carpeta or "[[EMPRESA]]" in nombre_carpeta:
                    nuevo_nombre = nombre_carpeta.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    os.rename(os.path.join(raiz, nombre_carpeta), os.path.join(raiz, nuevo_nombre))
                    consola.log(f"   [RENOMBRADO] Carpeta: {nuevo_nombre}")

        consola.set_progreso(100, 100, "Fase 1 completada con exito.")
        consola.log("\n[OK] Clonacion y configuracion del entorno finalizada.")
        return True 

    except FileExistsError:
        consola.log("\n[ERROR] Ya existe una carpeta con ese codigo y nombre de empresa.")
        return False
    except Exception as e:
        consola.log(f"\n[ERROR CRITICO] {str(e)}")
        return False

# ==========================================
# FASE 2: REEMPLAZAR TEXTOS
# ==========================================
def fase2_reemplazar_textos(carpeta_proyecto, ruta_excel, codigo, empresa, consola):
    if not ruta_excel or not os.path.exists(ruta_excel):
        consola.log("[ERROR] Archivo Excel de configuracion no encontrado.")
        return

    consola.log(f"==========================================")
    consola.log(f"FASE 2: SUSTITUCION PARAMETRICA DE TEXTOS")
    consola.log(f"==========================================")
    consola.log(f"Carpeta objetivo: {carpeta_proyecto}")
    consola.set_progreso(5, 100, "Iniciando motor COM de Word...")

    archivos_objetivo = {
        'PORTADA': [f"01 {codigo} PORTADA {empresa}.docx"],
        'MEMORIA ETIQUETAS': [f"03 {codigo} MEMORIA OBRAS {empresa}.docx", f"03 {codigo} MEMORIA ACT {empresa}.docx"],
        'ANEXO DISP MIN': [f"05 {codigo} ANEXO DISP MIN {empresa}.doc"],
        'EBSS': [f"05 {codigo} EBSS {empresa}.doc", f"07 {codigo} EBSS {empresa}.doc"],
        'PLIEGO CONDICIONES': [f"09 {codigo} PLIEGO CONDICIONES {empresa}.doc", f"11 {codigo} PLIEGO CONDICIONES {empresa}.doc"]
    }

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False 
        word.DisplayAlerts = False 

        total_archivos = sum(len(lista) for lista in archivos_objetivo.values())
        archivos_procesados = 0

        for pestana_excel, lista_nombres in archivos_objetivo.items():
            try:
                df_textos = pd.read_excel(ruta_excel, sheet_name=pestana_excel)
                consola.log(f"\n[INFO] Leyendo pestana Excel: [{pestana_excel}] con {len(df_textos)} etiquetas declaradas.")
            except Exception:
                continue 

            rutas_encontradas = buscar_archivos(carpeta_proyecto, lista_nombres)
            
            if not rutas_encontradas:
                consola.log(f"[INFO] Documentos para [{pestana_excel}] no encontrados en esta ruta.")

            for ruta_absoluta in rutas_encontradas:
                nombre_doc = os.path.basename(ruta_absoluta)
                consola.log(f"\n   [ABRIENDO DOCUMENTO] {nombre_doc}")
                doc = word.Documents.Open(ruta_absoluta)
                
                total_etiquetas = len(df_textos)
                etiquetas_procesadas = 0

                for index, row in df_textos.iterrows():
                    etiqueta = str(row.iloc[0]).strip()
                    valor = str(row.iloc[1]).strip()
                    etiquetas_procesadas += 1
                    
                    if valor.lower() == 'nan' or valor == '':
                        continue

                    consola.log(f"      > Buscando: {etiqueta} | Nuevo valor: {valor[:25]}...")
                    
                    progreso_actual = (archivos_procesados / total_archivos * 100) + ((etiquetas_procesadas / total_etiquetas) * (100 / total_archivos))
                    consola.set_progreso(progreso_actual, 100, f"Editando: {nombre_doc} ({etiquetas_procesadas}/{total_etiquetas})")

                    for story in doc.StoryRanges:
                        story.Find.Execute(FindText=etiqueta, ReplaceWith=valor, Replace=2)
                        while story.NextStoryRange:
                            story = story.NextStoryRange
                            story.Find.Execute(FindText=etiqueta, ReplaceWith=valor, Replace=2)

                doc.Save(); doc.Close()
                consola.log(f"   [GUARDADO OK] Cambios aplicados en: {nombre_doc}")
                archivos_procesados += 1

        word.Quit()
        consola.set_progreso(100, 100, "Sustitucion de textos completada.")
    except Exception as e:
        consola.log(f"\n[ERROR CRITICO] Fase 2: {str(e)}")
        try: word.Quit() 
        except: pass

# ==========================================
# FASE 3: INSERCION DE IMAGENES
# ==========================================
def fase3_insertar_imagenes(carpeta_proyecto, ruta_excel, codigo, empresa, consola):
    if not ruta_excel or not os.path.exists(ruta_excel):
        consola.log("[ERROR] Archivo Excel de configuracion no encontrado.")
        return

    consola.log(f"\n==========================================")
    consola.log(f"FASE 3: INSERCION DE IMAGENES TECNICAS")
    consola.log(f"==========================================")
    
    nombre_memoria = f"03 {codigo} MEMORIA ACT {empresa}.docx"
    rutas_encontradas = buscar_archivos(carpeta_proyecto, [nombre_memoria])

    if not rutas_encontradas:
        consola.log(f"[AVISO] No se encontro el documento objetivo: {nombre_memoria} en la ruta actual.")
        return

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df_tablas = pd.read_excel(ruta_excel, sheet_name='MEMORIA TABLAS')
            consola.log(f"[INFO] Leyendo base de datos de imagenes. Total declaradas: {len(df_tablas)}")
    except Exception:
        consola.log("[ERROR] Pestana [MEMORIA TABLAS] no encontrada en el Excel.")
        return

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False 
        
        total_imagenes = len(df_tablas)
        imagenes_procesadas = 0

        for ruta_absoluta in rutas_encontradas:
            consola.log(f"\n   [ABRIENDO DOCUMENTO] {os.path.basename(ruta_absoluta)}")
            doc = word.Documents.Open(ruta_absoluta)
            ancho_util = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

            for index, row in df_tablas.iterrows():
                etiqueta = str(row['Etiqueta']).strip()
                ruta_imagen = os.path.normpath(str(row['Direccion']).strip())
                imagenes_procesadas += 1

                consola.set_progreso(imagenes_procesadas, total_imagenes, f"Buscando marcador de imagen: {etiqueta}")

                if ruta_imagen.lower().endswith(('.png', '.jpg', '.jpeg')):
                    if os.path.exists(ruta_imagen):
                        consola.log(f"      > Insertando: {etiqueta} desde {ruta_imagen}")
                        word.Selection.Find.ClearFormatting()
                        if word.Selection.Find.Execute(FindText=etiqueta):
                            word.Selection.Delete()
                            imagen = word.Selection.InlineShapes.AddPicture(FileName=ruta_imagen)
                            imagen.LockAspectRatio = -1 
                            imagen.Width = ancho_util
                            imagen.Range.ParagraphFormat.Alignment = 1
                            consola.log(f"        [OK] Imagen acoplada correctamente.")
                        else:
                            consola.log(f"        [AVISO] Etiqueta {etiqueta} no encontrada en el texto.")
                    else:
                        consola.log(f"      > [ERROR] Archivo de imagen no encontrado en la ruta: {ruta_imagen}")

            doc.Save(); doc.Close()
            consola.log(f"   [GUARDADO OK] Memoria actualizada.")
            
        word.Quit()
        consola.set_progreso(100, 100, "Insercion de imagenes finalizada.")
    except Exception as e:
        consola.log(f"\n[ERROR CRITICO] Fase 3: {e}")
        try: word.Quit()
        except: pass

# ==========================================
# FASE 4: CUADROS EBSS 
# ==========================================
def exportar_rango_a_imagen(ws, rango_str, ruta_temp):
    try:
        rango = ws.Range(rango_str)
        if rango.Width == 0 or rango.Height == 0: return False
        rango.CopyPicture(Appearance=1, Format=2)
        time.sleep(1) 
        chart_obj = ws.ChartObjects().Add(50, 50, rango.Width, rango.Height)
        chart_obj.Activate() 
        try: chart_obj.Chart.ChartArea.Format.Line.Visible = 0
        except: pass
        chart_obj.Chart.Paste()
        time.sleep(0.5) 
        chart_obj.Chart.Export(ruta_temp)
        chart_obj.Delete()
        return True
    except Exception:
        try: chart_obj.Delete() 
        except: pass
        return False

def fase4_insertar_ebss(carpeta_proyecto, ruta_excel, codigo, empresa, consola):
    if not ruta_excel or not os.path.exists(ruta_excel):
        consola.log("[ERROR] Archivo Excel de configuracion no encontrado.")
        return

    consola.log(f"\n==========================================")
    consola.log(f"FASE 4: EXTRACCION Y MAQUETACION DE EBSS")
    consola.log(f"==========================================")
    consola.set_progreso(5, 100, "Analizando oficios habilitados en el Excel...")

    RANGO_CUADRO_1 = "A1:I37"
    RANGO_CUADRO_2 = "K39:P73"

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df_oficios = pd.read_excel(ruta_excel, sheet_name='PERSONAL_EBSS')
    except Exception:
        consola.log("[ERROR] Pestana [PERSONAL_EBSS] no encontrada.")
        return

    oficios_seleccionados = []
    for index, row in df_oficios.iterrows():
        oficio = str(row.iloc[0]).strip()
        marcado = str(row.iloc[1]).strip().upper()
        if marcado == 'SI':
            oficios_seleccionados.append(oficio)

    # Si no hay oficios seleccionados, no hay nada que procesar
    if not oficios_seleccionados:
        # Registrar aviso de cancelacion
        consola.log("[AVISO] La tabla indica que no hay oficios marcados con 'SI'. Operacion cancelada.")
        # Salir de la funcion sin continuar
        return

    # INFORMAR al usuario cuantos oficios se van a procesar
    # y mostrar los primeros 3 como muestra (si hay mas de 3, truncar)
    consola.log(f"[INFO] Detectados {len(oficios_seleccionados)} oficios a procesar: {', '.join(oficios_seleccionados[:3])}...")

    # PASO 2: Encontrar documentos Word destino donde insertar los cuadros EBSS
    # Buscar dos archivos especificos: uno de nivel 05 y otro de nivel 07
    nombres_destino = [f"05 {codigo} EBSS {empresa}.doc", f"07 {codigo} EBSS {empresa}.doc"]
    # Busqueda recursiva en la carpeta del proyecto
    docs_ebss_destino = buscar_archivos(carpeta_proyecto, nombres_destino)

    # Validar que encontramos documentos destino
    if not docs_ebss_destino:
        # Error: no hay documentos donde insertar los datos
        consola.log("[AVISO] No se encontraron documentos EBSS en la ruta objetivo para incrustar las tablas.")
        # Cancelar fase
        return

    # PASO 3: Encontrar el Excel fuente con las tablas EBSS
    # Este archivo contiene las hojas con los datos de cada oficio
    ruta_origen_ebss = encontrar_excel_ebss(carpeta_proyecto, codigo, empresa)
    # Validar que existe el archivo
    if not ruta_origen_ebss:
        # Error critico: no hay datos que copiar
        consola.log("[ERROR] No se encontro el archivo de calculo Excel de Tablas EBSS.")
        # Cancelar fase
        return

    # PASO 4: Crear ruta para archivo temporal de imagen
    # Las imagenes se generan temporalmente, se usan, y se eliminan
    ruta_temp = os.path.normpath(os.path.join(carpeta_proyecto, "temp_cuadro.png"))

    try:  # BLOQUE TRY: Capturar errores durante la automatizacion
        # PASO 5A: INICIAR APLICACION EXCEL
        # Usar COM (Component Object Model) para acceder a Excel
        excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        excel_app.Visible = False  # No mostrar ventana de Excel
        excel_app.DisplayAlerts = False  # No mostrar dialogos emergentes
        
        # PASO 5B: INICIAR APLICACION WORD
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False  # No mostrar ventana de Word
        word.DisplayAlerts = False  # No mostrar dialogos
        
        # PASO 5C: ABRIR ARCHIVO EXCEL CON DATOS EBSS
        # Este archivo contiene las hojas para cada oficio
        consola.log(f"\n   [ABRIENDO EXCEL DE CALCULO] {os.path.basename(ruta_origen_ebss)}")
        # Abrir como workbook (libro de trabajo)
        wb_tablas = excel_app.Workbooks.Open(ruta_origen_ebss)

        # PASO 6: RECORRER CADA DOCUMENTO WORD DESTINO
        # Procesaremos cada archivo Word secuencialmente
        for doc_idx, ruta_doc in enumerate(docs_ebss_destino):
            # Informar cual documento se esta procesando
            consola.log(f"\n   [ABRIENDO DOCUMENTO WORD] {os.path.basename(ruta_doc)}")
            # Abrir documento Word
            doc = word.Documents.Open(ruta_doc)
            
            # Limpiar criterios de busqueda previos
            word.Selection.Find.ClearFormatting()
            # BUSCAR etiqueta matriz donde insertar los cuadros
            if word.Selection.Find.Execute(FindText="[[TABLAS_OFICIOS]]"):
                # ENCONTRADA la etiqueta
                consola.log("      [+] Etiqueta matriz [[TABLAS_OFICIOS]] localizada. Preparando insercion...")
                # Eliminar la etiqueta (sera reemplazada por los cuadros)
                word.Selection.Delete()

                # Bandera para no insertar salto de pagina antes del primer oficio
                primer_paso = True
                # Total de oficios para barra de progreso
                total_oficios = len(oficios_seleccionados)
                
                # PASO 7: PROCESAR CADA OFICIO SELECCIONADO
                for idx, oficio in enumerate(oficios_seleccionados):
                    # Actualizar barra de progreso
                    consola.set_progreso(idx + 1, total_oficios, f"Procesando oficios EBSS: {oficio}...")
                    # Informar que se esta extrayendo este oficio
                    consola.log(f"      > Extrayendo datos de oficio: {oficio}")
                    try:
                        # OBTENER hoja Excel correspondiente a este oficio
                        # El nombre de la hoja debe coincidir exactamente con el nombre del oficio
                        ws = wb_tablas.Sheets(oficio)
                        # Activar la hoja (ponerla como activa en Excel)
                        ws.Activate()
                        
                        # Si NO es el primer oficio, insertar salto de pagina
                        if not primer_paso: 
                            # Type=7 = Page Break (salto de pagina)
                            word.Selection.InsertBreak(Type=7)
                            
                        # INSERTAR PRIMER CUADRO
                        # Exportar rango CUADRO_1 a imagen PNG temporal
                        if exportar_rango_a_imagen(ws, RANGO_CUADRO_1, ruta_temp) and os.path.exists(ruta_temp):
                            # Insertar imagen como figura incrustada (no vinculada)
                            shape1 = word.Selection.InlineShapes.AddPicture(
                                FileName=ruta_temp,  # Ruta de imagen temporal
                                LinkToFile=False,  # No vincular a archivo (copiar)
                                SaveWithDocument=True  # Guardar imagen dentro del .doc
                            )
                            # Desbloquear aspecto para redimensionamiento libre
                            shape1.LockAspectRatio = False
                            # Establecer dimensiones exactas (18 x 20 puntos * factor conversion)
                            # 28.35 = factor de conversion puntos a pixeles Word
                            shape1.Width = 18 * 28.35  # Ancho fijo
                            shape1.Height = 20 * 28.35  # Alto fijo
                            # Alineacion: 1 = centrado
                            shape1.Range.ParagraphFormat.Alignment = 1
                            # Mover cursor al final de la imagen
                            word.Selection.Collapse(Direction=0)
                            # Insertar parrafo nuevo
                            word.Selection.TypeParagraph()
                            # Confirmar en log
                            consola.log(f"        [OK] Cuadro 1 incrustado a escala estricta 18x20.")
                            # Eliminar archivo temporal
                            os.remove(ruta_temp)

                        # Esperar procesamiento
                        time.sleep(1)
                        # Insertar salto de pagina antes del segundo cuadro
                        word.Selection.InsertBreak(Type=7)
                        
                        # INSERTAR SEGUNDO CUADRO
                        # Exportar rango CUADRO_2 a imagen PNG temporal
                        if exportar_rango_a_imagen(ws, RANGO_CUADRO_2, ruta_temp) and os.path.exists(ruta_temp):
                            # Insertar imagen (mismo procedimiento que cuadro 1)
                            shape2 = word.Selection.InlineShapes.AddPicture(
                                FileName=ruta_temp,  # Ruta de imagen temporal
                                LinkToFile=False,  # No vincular
                                SaveWithDocument=True  # Guardar dentro del documento
                            )
                            # Desbloquear aspecto
                            shape2.LockAspectRatio = False
                            # Dimensiones exactas iguales a cuadro 1
                            shape2.Width = 18 * 28.35  # Ancho
                            shape2.Height = 20 * 28.35  # Alto
                            # Alineacion centrada
                            shape2.Range.ParagraphFormat.Alignment = 1
                            # Mover cursor
                            word.Selection.Collapse(Direction=0)
                            # Nuevo parrafo
                            word.Selection.TypeParagraph()
                            # Confirmar
                            consola.log(f"        [OK] Cuadro 2 incrustado a escala estricta 18x20.")
                            # Eliminar temporal
                            os.remove(ruta_temp)
                        
                        # Marcar que completamos el primer oficio
                        primer_paso = False
                    except Exception as e:
                        # Error en este oficio: registrar pero continuar con los demas
                        consola.log(f"        [ERROR] Excepcion capturada en oficio '{oficio}': {e}")
            else:
                # Etiqueta no encontrada en este documento
                consola.log(f"      [AVISO] No se encontro la etiqueta raiz [[TABLAS_OFICIOS]]")

            # Guardar documento con cuadros insertados
            doc.Save()
            # Cerrar documento
            doc.Close()
            # Confirmar que documento fue maquetado
            consola.log(f"   [GUARDADO OK] EBSS maquetado.")

        # PASO 8: LIMPIAR RECURSOS
        # Cerrar Excel sin guardar (no queremos guardar cambios en Excel)
        wb_tablas.Close(False)
        # Cerrar aplicacion Excel (libera memoria)
        excel_app.Quit()
        # Cerrar aplicacion Word
        word.Quit()
        # Actualizar progreso a 100%
        consola.set_progreso(100, 100, "Cuadros EBSS procesados correctamente.")

    except Exception as e:
        consola.log(f"\n[ERROR CRITICO] Fase 4: {e}")
        try: excel_app.Quit(); word.Quit()
        except: pass

# ==========================================
# FASE 5: EXPORTACION MASIVA A PDF
# ==========================================
def fase5_exportar_pdf(consola):
    pyautogui.alert("Seleccione la carpeta donde se encuentran los documentos Word que desea convertir.", "Conversor a PDF")
    
    root = tk.Tk(); root.withdraw()
    carpeta_origen = filedialog.askdirectory(title="Selecciona la carpeta de los Documentos Word")
    
    if not carpeta_origen:
        consola.log("[AVISO] Seleccion de carpeta cancelada por el usuario.")
        return

    carpeta_origen = os.path.normpath(carpeta_origen)
    carpeta_padre = os.path.dirname(carpeta_origen)
    carpeta_pdf = os.path.join(carpeta_padre, "03 PDF")

    consola.log(f"==========================================")
    consola.log(f"FASE 5: CONVERSION MASIVA A FORMATO PDF")
    consola.log(f"==========================================")
    consola.log(f"Origen de lectura: {carpeta_origen}")

    if not os.path.exists(carpeta_pdf):
        os.makedirs(carpeta_pdf)
        consola.log(f"[INFO] Creada nueva carpeta de entrega en: {carpeta_pdf}")
    else:
        consola.log(f"[INFO] Exportando en carpeta existente: {carpeta_pdf}")

    archivos_word = [f for f in os.listdir(carpeta_origen) if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
    
    if not archivos_word:
        consola.log("[ERROR] No se detectaron archivos Microsoft Word validos en el directorio indicado.")
        return

    total_archivos = len(archivos_word)
    consola.log(f"[INFO] Puestos en cola de impresion: {total_archivos} documentos.\n")

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        for idx, archivo in enumerate(archivos_word):
            ruta_doc = os.path.join(carpeta_origen, archivo)
            nombre_pdf = os.path.splitext(archivo)[0] + ".pdf"
            ruta_pdf = os.path.join(carpeta_pdf, nombre_pdf)

            consola.set_progreso(idx + 1, total_archivos, f"Convirtiendo a formato vectorial: {archivo}...")
            consola.log(f"   > Cargando motor de impresion: {archivo}")
            try:
                doc = word.Documents.Open(ruta_doc)
                doc.ExportAsFixedFormat(OutputFileName=ruta_pdf, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0, CreateBookmarks=1)
                doc.Close(False)
                consola.log(f"     [OK] PDF vectorizado y salvado con exito.")
            except Exception as e:
                consola.log(f"     [ERROR] Fallo en la conversion: {e}")

        word.Quit()
        consola.set_progreso(100, 100, "Motor de impresion detenido. Tarea finalizada.")
        consola.log("\n[FINALIZADO] Paquete de entregables generado.")

    except Exception as e:
        consola.log(f"\n[ERROR CRITICO] Fase 5: {e}")
        try: word.Quit()
        except: pass

# ==========================================
# INTERFAZ GRAFICA DEL MENU (TKINTER)
# ==========================================
# PROPOSITO: Crear la ventana principal del RPA donde el usuario:
# 1. Carga la plantilla de proyecto
# 2. Carga el Excel de configuracion
# 3. Selecciona cual fase desea ejecutar
# Esta interfaz es el punto de entrada a todo el sistema RPA
# ==========================================
def lanzar_interfaz_principal():
    # CONFIGURAR RUTAS POR DEFECTO
    # Diccionario con las rutas iniciales (seran actualizadas si el usuario carga otros archivos)
    rutas_cfg = {
        "plantilla": r"C:\Users\joleh\Dropbox\INGENIERIA\PLANTILLA",  # Ruta carpeta plantilla
        "excel": r"C:\Users\joleh\Desktop\MII\SEGUNDO\2Cuatrimestre\IIA\RPA_Jolehisy_Acevedo_Medina\Configuración_memoria.xlsx"  # Ruta Excel maestro
    }

    # CREAR VENTANA PRINCIPAL
    root = tk.Tk()  # Crear instancia de ventana Tk
    root.title("RPA Proyectos - Panel de Control")  # Titulo de ventana
    root.geometry("450x550")  # Tamaño inicial
    root.configure(bg="#f4f4f4")  # Color fondo gris claro
    
    # CENTRAR LA VENTANA EN LA PANTALLA (mismo procedimiento que en VentanaConsola)
    window_width, window_height = 450, 550  # Dimensiones de la ventana
    screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()  # Obtener tamaño pantalla
    # Calcular posicion para centrar (arriba-abajo)
    position_top = int(screen_height/2 - window_height/2)
    # Calcular posicion para centrar (izquierda-derecha)
    position_right = int(screen_width/2 - window_width/2)
    # Aplicar geometria centra
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    # VARIABLES DE CONTROL DE INTERFAZ
    # StringVar para almacenar la opcion elegida por el usuario
    eleccion_usuario = tk.StringVar()
    
    # Variable para mostrar estado de la plantilla (con checkmark o X)
    # chr(10004) = ✓ (checkmark), chr(10006) = ✗ (equis)
    txt_estado_plantilla = tk.StringVar(
        value="[ " + chr(10004) + " ] Plantilla Ok" if os.path.exists(rutas_cfg["plantilla"]) else "[ " + chr(10006) + " ] Sin cargar"
    )
    # Variable para mostrar estado del Excel (con checkmark o X)
    txt_estado_excel = tk.StringVar(
        value="[ " + chr(10004) + " ] Excel Ok" if os.path.exists(rutas_cfg["excel"]) else "[ " + chr(10006) + " ] Sin cargar"
    )

    # FUNCIONES DE CONTROL DE INTERFAZ
    # Funcion: Cuando usuario selecciona una opcion, guardarla y cerrar ventana
    def seleccionar(opcion):
        # Guardar la opcion elegida en la variable
        eleccion_usuario.set(opcion)
        # Cerrar ventana principal (destroy = destruir widget)
        root.destroy()

    # Funcion: Cargar carpeta plantilla personalizada
    def cargar_plantilla():
        # Abrir dialogo para seleccionar carpeta
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta Plantilla")
        # Si usuario selecciono una carpeta (no cancelo)
        if carpeta:
            # Guardar ruta normalizada en configuracion
            rutas_cfg["plantilla"] = os.path.normpath(carpeta)
            # Actualizar etiqueta de estado
            txt_estado_plantilla.set("[ " + chr(10004) + " ] Cargada")
            # Cambiar color a verde (#009900) para indicar exito
            lbl_plantilla.config(fg="#009900")

    # Funcion: Cargar archivo Excel personalizado
    def cargar_excel():
        # Abrir dialogo para seleccionar archivo Excel
        # filetypes filtra para mostrar solo archivos .xlsx y .xls
        archivo = filedialog.askopenfilename(
            title="Selecciona Excel de Configuracion", 
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        # Si usuario selecciono un archivo (no cancelo)
        if archivo:
            # Guardar ruta normalizada en configuracion
            rutas_cfg["excel"] = os.path.normpath(archivo)
            # Actualizar etiqueta de estado
            txt_estado_excel.set("[ " + chr(10004) + " ] Cargado")
            # Cambiar color a verde para indicar exito
            lbl_excel.config(fg="#009900")

    # ELEMENTO 1: TITULO PRINCIPAL
    tk.Label(
        root, 
        text="Asistente RPA de Proyectos",  # Texto del titulo
        font=("Segoe UI", 16, "bold"),  # Fuente, tamaño, estilo
        bg="#f4f4f4",  # Color fondo
        fg="#333333"  # Color texto gris oscuro
    ).pack(pady=(15, 5))  # Empaquetar con margen arriba 15, abajo 5

    # ELEMENTO 2: MARCO DE CONFIGURACION (PLANTILLA Y EXCEL)
    frame_config = tk.LabelFrame(
        root, 
        text=" Configuracion Previa ",  # Titulo del marco
        font=("Segoe UI", 10, "bold"),  # Fuente titulo
        bg="#f4f4f4",  # Color fondo
        fg="#333333",  # Color titulo
        padx=10,  # Margen interno izq-der
        pady=10  # Margen interno arriba-abajo
    )
    frame_config.pack(fill="x", padx=30, pady=10)  # Llenar horizontalmente, margen externo

    # BOTON Y ETIQUETA PARA PLANTILLA
    frame_p = tk.Frame(frame_config, bg="#f4f4f4")  # Marco para organizar boton y etiqueta
    frame_p.pack(fill="x", pady=2)  # Empaquetar horizontalmente
    # Boton para cargar plantilla
    tk.Button(
        frame_p, 
        text="Cargar Plantilla",  # Texto del boton
        width=18,  # Ancho
        command=cargar_plantilla,  # Funcion a ejecutar al clickear
        font=("Segoe UI", 9)  # Fuente
    ).pack(side="left")  # Colocar a la izquierda
    # Etiqueta de estado (checkmark o X)
    lbl_plantilla = tk.Label(
        frame_p, 
        textvariable=txt_estado_plantilla,  # Variable que muestra el estado
        font=("Consolas", 10, "bold"),  # Fuente monoespaciada
        bg="#f4f4f4",  # Color fondo
        # Color texto: verde si existe, rojo si no
        fg="#009900" if "Ok" in txt_estado_plantilla.get() else "#CC0000"
    )
    lbl_plantilla.pack(side="left", padx=10)  # Colocar a la derecha, con margen

    # BOTON Y ETIQUETA PARA EXCEL
    frame_e = tk.Frame(frame_config, bg="#f4f4f4")  # Marco para organizar
    frame_e.pack(fill="x", pady=2)  # Empaquetar horizontalmente
    # Boton para cargar Excel
    tk.Button(
        frame_e, 
        text="Cargar Excel",  # Texto del boton
        width=18,  # Ancho
        command=cargar_excel,  # Funcion a ejecutar al clickear
        font=("Segoe UI", 9)  # Fuente
    ).pack(side="left")  # Colocar a la izquierda
    # Etiqueta de estado
    lbl_excel = tk.Label(
        frame_e, 
        textvariable=txt_estado_excel,  # Variable que muestra el estado
        font=("Consolas", 10, "bold"),  # Fuente
        bg="#f4f4f4",  # Color fondo
        # Color texto: verde si existe, rojo si no
        fg="#009900" if "Ok" in txt_estado_excel.get() else "#CC0000"
    )
    lbl_excel.pack(side="left", padx=10)  # Colocar a la derecha

    # ELEMENTO 3: ETIQUETA INSTRUCCION
    tk.Label(
        root, 
        text="Selecciona el modulo que deseas ejecutar:",  # Texto de instruccion
        font=("Segoe UI", 10),  # Fuente
        bg="#f4f4f4",  # Color fondo
        fg="#666666"  # Color texto gris
    ).pack(pady=(5, 10))  # Margen superior 5, inferior 10

    # ELEMENTO 4: BOTONES DE OPCIONES
    # Lista de tuplas (texto_boton, color_fondo, color_texto)
    botones = [
        ("1. Crear Estructura de Proyecto", "#4CAF50", "white"),  # Verde para crear proyecto
        ("2. Actualizar Textos (Word)", "#2196F3", "white"),  # Azul para editar textos
        ("3. Insertar Imagenes (Memoria)", "#2196F3", "white"),  # Azul para imagenes
        ("4. Generar Cuadros EBSS", "#2196F3", "white"),  # Azul para EBSS
        ("5. Convertir Documentos a PDF", "#FF9800", "white"),  # Naranja para PDF
        ("6. Ejecucion Completa (Fases 2, 3 y 4)", "#9C27B0", "white")  # Morado para multifase
    ]

    # Crear cada boton con su configuracion
    for texto, bg, fg in botones:
        tk.Button(
            root, 
            text=texto,  # Etiqueta del boton
            bg=bg,  # Color fondo
            fg=fg,  # Color texto
            font=("Segoe UI", 11, "bold"),  # Fuente grande y negrita
            activebackground="#e0e0e0",  # Color al presionar
            relief="flat",  # Sin bordes 3D (plano)
            cursor="hand2",  # Cursor mano al pasar
            # Cuando se presiona, pasar el texto a seleccionar()
            command=lambda t=texto: seleccionar(t), 
            width=35,  # Ancho del boton
            pady=4  # Margen vertical interno
        ).pack(pady=3)  # Empaquetar con margen entre botones
        
    # INICIAR VENTANA (BLOQUEA HASTA QUE SE CIERRE)
    root.mainloop()
    
    # RETORNAR VALORES AL CIERRE
    # Retornar: opcion elegida, ruta plantilla, ruta excel
    return eleccion_usuario.get(), rutas_cfg["plantilla"], rutas_cfg["excel"]

# ==========================================
# MOTOR PRINCIPAL / PUNTO DE ENTRADA
# ==========================================
# PROPOSITO: Seccion que se ejecuta cuando se corre el script directamente
# Coordina todo el flujo del RPA desde la interfaz hasta la ejecucion
# ==========================================

if __name__ == "__main__":  # Solo ejecutar si es el archivo principal
    # PASO 1: LANZAR INTERFAZ MENU Y OBTENER ELECCION
    # Retorna: opcion elegida, ruta plantilla, ruta excel
    opcion_elegida, ruta_plantilla, excel_maestro = lanzar_interfaz_principal()

    # PASO 2: VALIDAR QUE NO CANCELO (selecciono algo)
    if opcion_elegida:
        # CREAR CONSOLA PARA MOSTRAR PROGRESO
        # titulo = muestra que modulo se esta ejecutando
        consola_activa = VentanaConsola(titulo=f"Ejecutando: {opcion_elegida}")

        # ========== OPCION 1: CREAR ESTRUCTURA ==========
        if opcion_elegida == '1. Crear Estructura de Proyecto':
            # Ejecutar fase 1 (clonado de plantilla y renombrado)
            fase1_crear_entorno(ruta_plantilla, consola_activa)
            # Notificar al usuario que termino
            pyautogui.alert("Modulo finalizado.", "RPA Completado")
            
        # ========== OPCION 5: CONVERSION PDF ==========
        elif opcion_elegida == '5. Convertir Documentos a PDF':
            # Ejecutar fase 5 (conversion de Word a PDF)
            fase5_exportar_pdf(consola_activa)
            # Notificar al usuario
            pyautogui.alert("Modulo finalizado.", "RPA Completado")
            
        # ========== OPCIONES 2, 3, 4, 6: REQUIEREN CARPETA ==========
        # Estas opciones trabajan sobre un proyecto existente
        elif opcion_elegida in ['2. Actualizar Textos (Word)', '3. Insertar Imagenes (Memoria)', '4. Generar Cuadros EBSS', '6. Ejecucion Completa (Fases 2, 3 y 4)']:
            # CREAR SELECTOR DE CARPETA
            root = tk.Tk()  # Ventana temporal
            root.withdraw()  # Ocultar ventana (solo mostrar dialogo)
            # DIALOGO: Seleccionar carpeta del proyecto
            carpeta = filedialog.askdirectory(title="Selecciona la carpeta a procesar")
            
            # VALIDAR: Usuario selecciono una carpeta
            if carpeta:
                # NORMALIZAR RUTA (convertir / a \\)
                carpeta = os.path.normpath(carpeta)
                
                # ===== SISTEMA RADAR: DETECTAR CODIGO Y EMPRESA AUTOMATICAMENTE =====
                # Este sistema intenta descubrir el codigo y empresa subiendo por carpetas
                # Util si el usuario selecciono una subcarpeta en lugar de la carpeta raiz
                
                # Inicializar variables de deteccion
                ruta_analisis = carpeta  # Punto de partida para busqueda
                cod_sug = ""  # Codigo sugerido (vacio al inicio)
                emp_sug = ""  # Empresa sugerida (vacia al inicio)
                
                # BUCLE: Subir en la jerarquia de carpetas hasta encontrar patron
                # Busca una carpeta con nombre "CODIGO EMPRESA"
                # CODIGO debe ser numerico para identificarse como codigo
                while ruta_analisis != os.path.dirname(ruta_analisis):  # Hasta llegar a raiz
                    # Obtener nombre de la carpeta actual
                    nombre_c = os.path.basename(ruta_analisis)
                    # Dividir nombre por espacio (espera formato "CODIGO EMPRESA")
                    partes = nombre_c.split(' ', 1)  # Maximo 2 partes
                    # VERIFICAR: Tiene 2 partes Y la primera es numerica
                    if len(partes) > 1 and partes[0].isdigit():
                        # ENCONTRADO! Guardar codigo y empresa detectados
                        cod_sug = partes[0]  # Primer componente = codigo
                        emp_sug = partes[1]  # Segundo componente = empresa
                        break  # Salir del bucle (encontramos)
                    # NO ENCONTRADO: Subir un nivel en la jerarquia
                    ruta_analisis = os.path.dirname(ruta_analisis)  # Ir a carpeta padre

                # ===== SOLICITAR AL USUARIO CONFIRMACION DE CODIGO =====
                # Mostrar dialogo con valor sugerido (si se encontro) o vacio
                codigo = pyautogui.prompt(
                    "Confirma el CODIGO del proyecto:",  # Pregunta
                    "Datos de Ejecucion",  # Titulo
                    default=cod_sug  # Valor por defecto (sugerencia)
                )
                # VALIDAR: Usuario ingreso codigo (no cancelo)
                if not codigo: 
                    # Usuario cancelo o dejo vacio
                    consola_activa.log("[AVISO] Operacion cancelada. Falta el codigo.")
                else:
                    # ===== SOLICITAR AL USUARIO CONFIRMACION DE EMPRESA =====
                    empresa = pyautogui.prompt(
                        "Confirma el nombre de la EMPRESA:",  # Pregunta
                        "Datos de Ejecucion",  # Titulo
                        default=emp_sug  # Valor por defecto (sugerencia)
                    )
                    # VALIDAR: Usuario ingreso empresa (no cancelo)
                    if not empresa:
                        # Usuario cancelo o dejo vacio
                        consola_activa.log("[AVISO] Operacion cancelada. Falta la empresa.")
                    else:
                        # ===== EJECUTAR LA FASE SELECCIONADA =====
                        # Ya tenemos: carpeta, codigo, empresa, excel_maestro, consola_activa
                        
                        # OPCION 2: Solo reemplazar textos
                        if opcion_elegida == '2. Actualizar Textos (Word)':
                            fase2_reemplazar_textos(carpeta, excel_maestro, codigo, empresa, consola_activa)
                        # OPCION 3: Solo insertar imagenes
                        elif opcion_elegida == '3. Insertar Imagenes (Memoria)':
                            fase3_insertar_imagenes(carpeta, excel_maestro, codigo, empresa, consola_activa)
                        # OPCION 4: Solo generar EBSS
                        elif opcion_elegida == '4. Generar Cuadros EBSS':
                            fase4_insertar_ebss(carpeta, excel_maestro, codigo, empresa, consola_activa)
                        # OPCION 6: Ejecutar todas las fases (2, 3 y 4)
                        elif opcion_elegida == '6. Ejecucion Completa (Fases 2, 3 y 4)':
                            # EJECUTAR EN SECUENCIA
                            # Primero textos, luego imagenes, luego EBSS
                            fase2_reemplazar_textos(carpeta, excel_maestro, codigo, empresa, consola_activa)
                            fase3_insertar_imagenes(carpeta, excel_maestro, codigo, empresa, consola_activa)
                            fase4_insertar_ebss(carpeta, excel_maestro, codigo, empresa, consola_activa)
                        
                        # MENSAJE FINAL DE EXITO
                        consola_activa.log("\n==========================================")  # Separador
                        consola_activa.log("PROCESO TOTALMENTE COMPLETADO")  # Mensaje principal
                        consola_activa.log("==========================================")  # Separador
                        # Alerta de exito
                        pyautogui.alert("Modulo finalizado con exito.", "RPA Completado")
            else:
                # Usuario cancelo selector de carpeta
                consola_activa.log("[AVISO] Operacion cancelada. No se selecciono ninguna carpeta.")

        # ===== MANTENER CONSOLA ABIERTA =====
        # Actualizar mensaje de estado final
        consola_activa.lbl_estado.config(text="Proceso terminado. Puedes cerrar esta ventana con seguridad.")
        # Mostrar ventana de consola bloqueante (espera hasta que usuario la cierre)
        consola_activa.root.mainloop()