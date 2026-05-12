import os
import shutil
import pyautogui
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import warnings
import time 

# Modulo de automatizacion para gestionar proyectos RPA.
# Contiene funciones para clonar una plantilla, reemplazar textos en Word,
# insertar imagenes, generar cuadros EBSS y convertir documentos a PDF.

# ==========================================
# FUNCIONES DE BUSQUEDA DINAMICA
# ==========================================
def buscar_archivos(carpeta_base, nombres_buscar):
    # Explora recursivamente la carpeta base y devuelve las rutas de los archivos
    # cuyo nombre coincide con alguno de los nombres objetivo.
    encontrados = []
    for raiz, _, archivos in os.walk(carpeta_base):
        for archivo in archivos:
            if archivo in nombres_buscar:
                encontrados.append(os.path.normpath(os.path.join(raiz, archivo)))
    return encontrados

def encontrar_excel_ebss(carpeta_base, codigo, empresa):
    # Busca el archivo de tablas EBSS dentro de la carpeta base y, si no lo
    # encuentra, realiza una busqueda ascendiente hasta la carpeta del proyecto.
    nombre_excel = f"{codigo} TABLAS EBSS {empresa}.xlsx"
    for raiz, _, archivos in os.walk(carpeta_base):
        if nombre_excel in archivos:
            return os.path.normpath(os.path.join(raiz, nombre_excel))
    
    ruta = carpeta_base
    nombre_proyecto = f"{codigo} {empresa}"
    # Busca hacia arriba en la jerarquia de carpetas hasta encontrar la raiz del proyecto.
    while os.path.basename(ruta) != nombre_proyecto and os.path.dirname(ruta) != ruta:
        ruta = os.path.dirname(ruta)
    
    for raiz, _, archivos in os.walk(ruta):
        if nombre_excel in archivos:
            return os.path.normpath(os.path.join(raiz, nombre_excel))
            
    return None

# ==========================================
# FASE 1: CLONAR Y RENOMBRAR
# ==========================================
def fase1_crear_entorno(ruta_plantilla):
    if not ruta_plantilla or not os.path.exists(ruta_plantilla):
        pyautogui.alert("Error: La ruta de la plantilla no es valida o no existe. Carguela en el menu principal.", "Error de Configuracion")
        return False

    # Solicita los datos de proyecto y empresa al usuario.
    codigo_proyecto = pyautogui.prompt("Introduce el CODIGO del proyecto:", "Datos del Proyecto")
    if not codigo_proyecto: return None 
        
    nombre_empresa = pyautogui.prompt("Introduce el NOMBRE DE LA EMPRESA:", "Datos del Proyecto")
    if not nombre_empresa: return None

    # Define la ruta de la plantilla y la carpeta destino de la copia.
    directorio_padre = os.path.normpath(os.path.dirname(ruta_plantilla))
    nombre_nueva_carpeta = f"{codigo_proyecto} {nombre_empresa}"
    carpeta_destino = os.path.join(directorio_padre, nombre_nueva_carpeta)

    print("\n" + "="*50)
    print(f"INICIANDO FASE 1: CLONADO DEL PROYECTO")
    print("="*50)
    print(f"Creando carpeta: {nombre_nueva_carpeta}...")

    try:
        # Copia la plantilla completa en la nueva ubicacion.
        shutil.copytree(ruta_plantilla, carpeta_destino)

        # Recorre la estructura copiada de abajo hacia arriba para renombrar los
        # archivos y carpetas que contienen los marcadores de plantilla.
        for raiz, carpetas, archivos in os.walk(carpeta_destino, topdown=False):
            for nombre_archivo in archivos:
                if "XXXX" in nombre_archivo or "[[EMPRESA]]" in nombre_archivo:
                    nuevo_nombre = nombre_archivo.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    os.rename(os.path.join(raiz, nombre_archivo), os.path.join(raiz, nuevo_nombre))
            
            for nombre_carpeta in carpetas:
                if "XXXX" in nombre_carpeta or "[[EMPRESA]]" in nombre_carpeta:
                    nuevo_nombre = nombre_carpeta.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    os.rename(os.path.join(raiz, nombre_carpeta), os.path.join(raiz, nuevo_nombre))

        print("Fase 1 completada con exito.")
        pyautogui.alert("La estructura de carpetas se ha creado con exito.", "Completado")
        return True 

    except FileExistsError:
        pyautogui.alert("Error: ya existe una carpeta con ese codigo y nombre de empresa.", "Error")
        return False
    except Exception as e:
        pyautogui.alert(f"Ha ocurrido un error inesperado:\n{str(e)}", "Error Critico")
        return False

# ==========================================
# FASE 2: REEMPLAZAR TEXTOS
# ==========================================
def fase2_reemplazar_textos(carpeta_proyecto, ruta_excel, codigo, empresa):
    # Verifica que el archivo Excel de configuracion exista antes de proceder.
    if not ruta_excel or not os.path.exists(ruta_excel):
        print("ERROR: Archivo Excel de configuracion no encontrado. Carguelo en el menu.")
        return

    # Muestra un mensaje de inicio de la fase para informar al usuario.
    print("\n" + "="*50)
    print(f"INICIANDO FASE 2: REEMPLAZO DE ETIQUETAS DE TEXTO")
    print("="*50)

    # Define un diccionario que mapea las pestanas del Excel con los nombres de archivos Word a procesar.
    # Cada clave es una pestana del Excel, y el valor es una lista de nombres de archivos.
    archivos_objetivo = {
        'PORTADA': [f"01 {codigo} PORTADA {empresa}.docx"],
        'MEMORIA ETIQUETAS': [f"03 {codigo} MEMORIA OBRAS {empresa}.docx", f"03 {codigo} MEMORIA ACT {empresa}.docx"],
        'ANEXO DISP MIN': [f"05 {codigo} ANEXO DISP MIN {empresa}.doc"],
        'EBSS': [f"05 {codigo} EBSS {empresa}.doc", f"07 {codigo} EBSS {empresa}.doc"],
        'PLIEGO CONDICIONES': [f"09 {codigo} PLIEGO CONDICIONES {empresa}.doc", f"11 {codigo} PLIEGO CONDICIONES {empresa}.doc"]
    }

    try:
        # Inicializa la aplicacion Word en segundo plano para evitar mostrar la interfaz al usuario.
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False 
        word.DisplayAlerts = False  # Desactiva alertas para evitar interrupciones.

        # Itera sobre cada pestana del Excel y sus archivos correspondientes.
        for pestana_excel, lista_nombres in archivos_objetivo.items():
            try:
                # Intenta cargar la hoja de Excel especificada por la pestana.
                df_textos = pd.read_excel(ruta_excel, sheet_name=pestana_excel)
            except Exception:
                # Si la pestana no existe, salta a la siguiente.
                continue 

            # Busca los archivos Word en la carpeta del proyecto que coincidan con los nombres definidos.
            rutas_encontradas = buscar_archivos(carpeta_proyecto, lista_nombres)

            # Procesa cada documento Word encontrado.
            for ruta_absoluta in rutas_encontradas:
                print(f"   Modificando textos: {os.path.basename(ruta_absoluta)}")
                # Abre el documento Word para editarlo.
                doc = word.Documents.Open(ruta_absoluta)

                # Itera sobre cada fila del DataFrame de la pestana Excel.
                for index, row in df_textos.iterrows():
                    # Extrae la etiqueta (columna 0) y el valor (columna 1) de la fila.
                    etiqueta = str(row.iloc[0]).strip()
                    valor = str(row.iloc[1]).strip()
                    # Omite filas donde el valor sea nulo o vacio.
                    if valor.lower() == 'nan' or valor == '':
                        continue

                    # Reemplaza la etiqueta en todas las historias del documento (cuerpo, encabezados, pies, etc.).
                    for story in doc.StoryRanges:
                        story.Find.Execute(FindText=etiqueta, ReplaceWith=valor, Replace=2)  # Replace=2 significa reemplazar todas las ocurrencias.
                        while story.NextStoryRange:
                            story = story.NextStoryRange
                            story.Find.Execute(FindText=etiqueta, ReplaceWith=valor, Replace=2)

                # Guarda y cierra el documento despues de procesarlo.
                doc.Save(); doc.Close()
                print(f"   COMPLETADO: {os.path.basename(ruta_absoluta)}")

        # Cierra la aplicacion Word al finalizar.
        word.Quit()
    except Exception as e:
        print(f"Ocurrio un error en la Fase 2:\n{str(e)}")
        try: word.Quit()  # Asegura que Word se cierre incluso si hay error.
        except: pass

# ==========================================
# FASE 3: INSERCION DE IMAGENES
# ==========================================
def fase3_insertar_imagenes(carpeta_proyecto, ruta_excel, codigo, empresa):
    # Verifica que el archivo Excel de configuracion exista.
    if not ruta_excel or not os.path.exists(ruta_excel):
        print("ERROR: Archivo Excel de configuracion no encontrado. Carguelo en el menu.")
        return

    # Muestra mensaje de inicio de la fase.
    print("\n" + "="*50)
    print(f"INICIANDO FASE 3: INSERCION DE IMAGENES")
    print("="*50)

    # Define el nombre del documento de memoria de actividad a procesar.
    nombre_memoria = f"03 {codigo} MEMORIA ACT {empresa}.docx"
    # Busca el archivo en la carpeta del proyecto.
    rutas_encontradas = buscar_archivos(carpeta_proyecto, [nombre_memoria])

    # Si no se encuentra el documento, informa y sale.
    if not rutas_encontradas:
        print("   No se encontro la memoria de actividad en la carpeta seleccionada.")
        return

    try:
        # Suprime advertencias al leer el Excel para evitar mensajes innecesarios.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Carga la hoja 'MEMORIA TABLAS' que contiene las etiquetas y rutas de imagenes.
            df_tablas = pd.read_excel(ruta_excel, sheet_name='MEMORIA TABLAS')
    except Exception:
        # Si no puede cargar la hoja, sale silenciosamente.
        return

    try:
        # Inicializa Word en segundo plano.
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False 

        # Procesa cada documento de memoria encontrado (generalmente uno).
        for ruta_absoluta in rutas_encontradas:
            print(f"   Procesando Word: {os.path.basename(ruta_absoluta)}")
            # Abre el documento Word.
            doc = word.Documents.Open(ruta_absoluta)
            # Calcula el ancho util de la pagina (ancho total menos margenes).
            ancho_util = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

            # Itera sobre cada fila de la tabla de imagenes.
            for index, row in df_tablas.iterrows():
                # Extrae la etiqueta y la ruta de la imagen de la fila.
                etiqueta = str(row['Etiqueta']).strip()
                ruta_imagen = os.path.normpath(str(row['Direccion']).strip())

                # Verifica que la ruta sea una imagen valida y que el archivo exista.
                if ruta_imagen.lower().endswith(('.png', '.jpg', '.jpeg')) and os.path.exists(ruta_imagen):
                    # Limpia el formato de busqueda para evitar interferencias.
                    word.Selection.Find.ClearFormatting()
                    # Busca la etiqueta en el documento.
                    if word.Selection.Find.Execute(FindText=etiqueta):
                        # Elimina la etiqueta encontrada.
                        word.Selection.Delete()
                        # Inserta la imagen en la posicion de la etiqueta.
                        imagen = word.Selection.InlineShapes.AddPicture(FileName=ruta_imagen)
                        # Bloquea la relacion de aspecto para mantener proporciones.
                        imagen.LockAspectRatio = -1 
                        # Ajusta el ancho de la imagen al ancho util de la pagina.
                        imagen.Width = ancho_util
                        # Centra la imagen horizontalmente.
                        imagen.Range.ParagraphFormat.Alignment = 1
                        print(f"      Imagen insertada: {etiqueta}")

            # Guarda y cierra el documento.
            doc.Save(); doc.Close()
            
        # Cierra Word.
        word.Quit()
    except Exception:
        # En caso de error, intenta cerrar Word.
        try: word.Quit()
        except: pass

# ==========================================
# FASE 4: CUADROS EBSS 
# ==========================================
def exportar_rango_a_imagen(ws, rango_str, ruta_temp):
    # Funcion auxiliar para exportar un rango de celdas de Excel como imagen PNG.
    try:
        # Selecciona el rango especificado en la hoja de trabajo.
        rango = ws.Range(rango_str)
        # Verifica que el rango tenga dimensiones validas (ancho y alto mayores a cero).
        if rango.Width == 0 or rango.Height == 0:
            return False

        # Copia el rango como imagen (Appearance=1 para pantalla, Format=2 para bitmap).
        rango.CopyPicture(Appearance=1, Format=2)
        time.sleep(1)  # Espera para que la copia se complete.

        # Crea un objeto de grafico temporal en la hoja para pegar la imagen.
        chart_obj = ws.ChartObjects().Add(50, 50, rango.Width, rango.Height)
        chart_obj.Activate()
        # Intenta hacer invisible el borde del grafico (puede fallar en algunas versiones).
        try: chart_obj.Chart.ChartArea.Format.Line.Visible = 0
        except: pass
        # Pega la imagen copiada en el grafico.
        chart_obj.Chart.Paste()
        time.sleep(0.5)  # Espera para que el pegado se complete.
        # Exporta el grafico como imagen PNG.
        chart_obj.Chart.Export(ruta_temp)
        # Elimina el objeto grafico temporal.
        chart_obj.Delete()
        return True
    except Exception:
        # En caso de error, intenta eliminar el grafico si existe.
        try: chart_obj.Delete()
        except: pass
        return False

def fase4_insertar_ebss(carpeta_proyecto, ruta_excel, codigo, empresa):
    # Verifica que el Excel de configuracion exista.
    if not ruta_excel or not os.path.exists(ruta_excel):
        print("ERROR: Archivo Excel de configuracion no encontrado. Carguelo en el menu.")
        return

    # Muestra mensaje de inicio.
    print("\n" + "="*50)
    print(f"INICIANDO FASE 4: GENERACION DE CUADROS EBSS")
    print("="*50)

    # Define los rangos de celdas para los cuadros 1 y 2 en Excel.
    RANGO_CUADRO_1 = "A1:I37"
    RANGO_CUADRO_2 = "K39:P73"

    try:
        # Suprime advertencias y carga la hoja 'PERSONAL_EBSS' que lista los oficios.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df_oficios = pd.read_excel(ruta_excel, sheet_name='PERSONAL_EBSS')
    except Exception:
        print("Pestana [PERSONAL_EBSS] no encontrada.")
        return

    # Filtra los oficios marcados con 'SI' en la columna 1.
    oficios_seleccionados = []
    for index, row in df_oficios.iterrows():
        oficio = str(row.iloc[0]).strip()
        marcado = str(row.iloc[1]).strip().upper()
        if marcado == 'SI':
            oficios_seleccionados.append(oficio)

    # Si no hay oficios seleccionados, informa y sale.
    if not oficios_seleccionados:
        print("   No hay oficios marcados con 'SI'.")
        return

    # Define los nombres de los documentos EBSS destino.
    nombres_destino = [f"05 {codigo} EBSS {empresa}.doc", f"07 {codigo} EBSS {empresa}.doc"]
    # Busca estos documentos en la carpeta del proyecto.
    docs_ebss_destino = buscar_archivos(carpeta_proyecto, nombres_destino)

    # Si no se encuentran, informa y sale.
    if not docs_ebss_destino:
        print("   No se encontraron documentos EBSS en la carpeta seleccionada.")
        return

    # Busca el Excel de tablas EBSS en la estructura del proyecto.
    ruta_origen_ebss = encontrar_excel_ebss(carpeta_proyecto, codigo, empresa)
    if not ruta_origen_ebss:
        print("Error: no se encontro el Excel de Tablas EBSS.")
        return

    # Define la ruta temporal para la imagen del cuadro.
    ruta_temp = os.path.normpath(os.path.join(carpeta_proyecto, "temp_cuadro.png"))

    try:
        # Inicializa Excel y Word en segundo plano.
        excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        
        print(f"   Abriendo Excel: {os.path.basename(ruta_origen_ebss)}")
        # Abre el libro de Excel con las tablas EBSS.
        wb_tablas = excel_app.Workbooks.Open(ruta_origen_ebss)

        # Procesa cada documento EBSS destino.
        for ruta_doc in docs_ebss_destino:
            print(f"   Procesando Word: {os.path.basename(ruta_doc)}")
            # Abre el documento Word.
            doc = word.Documents.Open(ruta_doc)
            
            # Busca la etiqueta [[TABLAS_OFICIOS]] para reemplazarla con los cuadros.
            word.Selection.Find.ClearFormatting()
            if word.Selection.Find.Execute(FindText="[[TABLAS_OFICIOS]]"):
                # Elimina la etiqueta.
                word.Selection.Delete()

                # Bandera para saber si es el primer oficio (no insertar salto de pagina).
                primer_paso = True
                # Para cada oficio seleccionado.
                for oficio in oficios_seleccionados:
                    try:
                        # Activa la hoja correspondiente al oficio en Excel.
                        ws = wb_tablas.Sheets(oficio)
                        ws.Activate()

                        # Si no es el primer oficio, inserta un salto de pagina.
                        if not primer_paso:
                            word.Selection.InsertBreak(Type=7)  # Type=7 es salto de pagina.

                        # Exporta el rango del cuadro 1 como imagen y la inserta si tiene exito.
                        if exportar_rango_a_imagen(ws, RANGO_CUADRO_1, ruta_temp) and os.path.exists(ruta_temp):
                            # Inserta la imagen en Word.
                            shape1 = word.Selection.InlineShapes.AddPicture(FileName=ruta_temp, LinkToFile=False, SaveWithDocument=True)
                            shape1.LockAspectRatio = False
                            # Establece dimensiones fijas (18x20 cm en puntos, 28.35 puntos por cm).
                            shape1.Width = 18 * 28.35
                            shape1.Height = 20 * 28.35
                            # Centra la imagen.
                            shape1.Range.ParagraphFormat.Alignment = 1
                            # Colapsa la seleccion al final y agrega un parrafo.
                            word.Selection.Collapse(Direction=0)
                            word.Selection.TypeParagraph()
                            print(f"      {oficio} - Cuadro 1 importado correctamente")
                            # Elimina el archivo temporal.
                            os.remove(ruta_temp)

                        # Espera y inserta un salto de seccion.
                        time.sleep(1)
                        word.Selection.InsertBreak(Type=7)

                        # Exporta e inserta el cuadro 2 de manera similar.
                        if exportar_rango_a_imagen(ws, RANGO_CUADRO_2, ruta_temp) and os.path.exists(ruta_temp):
                            shape2 = word.Selection.InlineShapes.AddPicture(FileName=ruta_temp, LinkToFile=False, SaveWithDocument=True)
                            shape2.LockAspectRatio = False
                            shape2.Width = 18 * 28.35
                            shape2.Height = 20 * 28.35
                            shape2.Range.ParagraphFormat.Alignment = 1
                            word.Selection.Collapse(Direction=0)
                            word.Selection.TypeParagraph()
                            print(f"      {oficio} - Cuadro 2 importado correctamente")
                            os.remove(ruta_temp)

                        # Marca que ya no es el primer paso.
                        primer_paso = False
                    except Exception as e:
                        print(f"      Error en oficio '{oficio}': {e}")
            else:
                print(f"   No se encontro la etiqueta [[TABLAS_OFICIOS]]")

            # Guarda y cierra el documento Word.
            doc.Save(); doc.Close()

        # Cierra el libro de Excel y las aplicaciones.
        wb_tablas.Close(False)
        excel_app.Quit()
        word.Quit()

    except Exception as e:
        print(f"Error critico en Fase 4: {e}")
        # Asegura que las aplicaciones se cierren en caso de error.
        try:
            excel_app.Quit()
            word.Quit()
        except:
            pass

# ==========================================
# FASE 5: EXPORTACION MASIVA A PDF
# ==========================================
def fase5_exportar_pdf():
    pyautogui.alert("Seleccione la carpeta donde se encuentran los documentos Word que desea convertir.\n(Ej: la carpeta '01 Documentos')", "Conversor a PDF")
    
    root = tk.Tk(); root.withdraw()
    carpeta_origen = filedialog.askdirectory(title="Selecciona la carpeta de los Documentos Word")
    
    if not carpeta_origen:
        print("Operacion cancelada. No se selecciono ninguna carpeta.")
        return

    carpeta_origen = os.path.normpath(carpeta_origen)
    carpeta_padre = os.path.dirname(carpeta_origen)
    carpeta_pdf = os.path.join(carpeta_padre, "03 PDF")

    if not os.path.exists(carpeta_pdf):
        os.makedirs(carpeta_pdf)
        print(f"Creada nueva carpeta de destino: {carpeta_pdf}")

    archivos_word = [f for f in os.listdir(carpeta_origen) if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
    
    if not archivos_word:
        pyautogui.alert("No se encontraron documentos Word en la carpeta seleccionada.", "Sin archivos")
        return

    print("\n" + "="*50)
    print(f"INICIANDO FASE 5: CONVERSION A PDF")
    print("="*50)
    print(f"   Origen:  {carpeta_origen}")
    print(f"   Destino: {carpeta_pdf}\n")

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        for archivo in archivos_word:
            ruta_doc = os.path.join(carpeta_origen, archivo)
            nombre_pdf = os.path.splitext(archivo)[0] + ".pdf"
            ruta_pdf = os.path.join(carpeta_pdf, nombre_pdf)

            print(f"   Convirtiendo: {archivo} ...")
            try:
                doc = word.Documents.Open(ruta_doc)
                doc.ExportAsFixedFormat(OutputFileName=ruta_pdf, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0, CreateBookmarks=1)
                doc.Close(False)
                print(f"      Guardado: {nombre_pdf}")
            except Exception as e:
                print(f"      Error al convertir {archivo}: {e}")

        word.Quit()
        
        print("\n" + "="*50)
        print("CONVERSION COMPLETADA")
        print("="*50 + "\n")
        pyautogui.alert(f"Todos los documentos han sido convertidos.\n\nSe han guardado en:\n{carpeta_pdf}", "PDFs Generados")

    except Exception as e:
        print(f"Error critico en Fase 5: {e}")
        try: word.Quit()
        except: pass

# ==========================================
# INTERFAZ GRAFICA DEL MENU (TKINTER)
# ==========================================
def lanzar_interfaz_principal():
    # Diccionario con las rutas por defecto para plantilla y Excel de configuracion.
    rutas_cfg = {
        "plantilla": r"C:\Users\joleh\Dropbox\INGENIERIA\PLANTILLA",
        "excel": r"C:\Users\joleh\Desktop\MII\SEGUNDO\2Cuatrimestre\IIA\RPA_Jolehisy_Acevedo_Medina\Configuración_memoria.xlsx"
    }

    # Crea la ventana principal de Tkinter.
    root = tk.Tk()
    root.title("RPA Proyectos - Panel de Control")
    root.geometry("450x550")
    root.configure(bg="#f4f4f4")
    
    # Calcula la posicion para centrar la ventana en la pantalla.
    window_width, window_height = 450, 550
    screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
    position_top, position_right = int(screen_height/2 - window_height/2), int(screen_width/2 - window_width/2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    # Variable para almacenar la eleccion del usuario.
    eleccion_usuario = tk.StringVar()
    
    # Variables para mostrar el estado de carga de archivos con iconos ASCII.
    txt_estado_plantilla = tk.StringVar(value="[ " + chr(10004) + " ] Carga la Plantilla" if os.path.exists(rutas_cfg["plantilla"]) else "[ " + chr(10006) + " ] Sin cargar")
    txt_estado_excel = tk.StringVar(value="[ " + chr(10004) + " ] Carga el Excel" if os.path.exists(rutas_cfg["excel"]) else "[ " + chr(10006) + " ] Sin cargar")

    def seleccionar(opcion):
        # Funcion para capturar la opcion seleccionada y cerrar la ventana.
        eleccion_usuario.set(opcion)
        root.destroy()

    def cargar_plantilla():
        # Funcion para cargar una nueva ruta de plantilla via filedialog.
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta Plantilla")
        if carpeta:
            rutas_cfg["plantilla"] = os.path.normpath(carpeta)
            txt_estado_plantilla.set("[ " + chr(10004) + " ] Cargada")
            lbl_plantilla.config(fg="#009900")  # Cambia el color del label a verde.

    def cargar_excel():
        # Funcion para cargar un nuevo archivo Excel de configuracion.
        archivo = filedialog.askopenfilename(title="Selecciona Excel de Configuracion", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if archivo:
            rutas_cfg["excel"] = os.path.normpath(archivo)
            txt_estado_excel.set("[ " + chr(10004) + " ] Cargado")
            lbl_excel.config(fg="#009900")  # Cambia el color del label a verde.

    # Etiqueta de cabecera de la interfaz.
    tk.Label(root, text="Asistente RPA de Proyectos", font=("Segoe UI", 16, "bold"), bg="#f4f4f4", fg="#333333").pack(pady=(15, 5))

    # Frame para la seccion de configuracion previa.
    frame_config = tk.LabelFrame(root, text=" Configuracion Previa ", font=("Segoe UI", 10, "bold"), bg="#f4f4f4", fg="#333333", padx=10, pady=10)
    frame_config.pack(fill="x", padx=30, pady=10)

    # Fila para la plantilla: boton y label de estado.
    frame_p = tk.Frame(frame_config, bg="#f4f4f4")
    frame_p.pack(fill="x", pady=2)
    tk.Button(frame_p, text="Cargar Plantilla", width=18, command=cargar_plantilla, font=("Segoe UI", 9)).pack(side="left")
    lbl_plantilla = tk.Label(frame_p, textvariable=txt_estado_plantilla, font=("Consolas", 10, "bold"), bg="#f4f4f4", fg="#009900" if "Ok" in txt_estado_plantilla.get() else "#CC0000")
    lbl_plantilla.pack(side="left", padx=10)

    # Fila para el Excel: boton y label de estado.
    frame_e = tk.Frame(frame_config, bg="#f4f4f4")
    frame_e.pack(fill="x", pady=2)
    tk.Button(frame_e, text="Cargar Excel", width=18, command=cargar_excel, font=("Segoe UI", 9)).pack(side="left")
    lbl_excel = tk.Label(frame_e, textvariable=txt_estado_excel, font=("Consolas", 10, "bold"), bg="#f4f4f4", fg="#009900" if "Ok" in txt_estado_excel.get() else "#CC0000")
    lbl_excel.pack(side="left", padx=10)

    # Etiqueta instructiva para seleccionar modulo.
    tk.Label(root, text="Selecciona el modulo que deseas ejecutar:", font=("Segoe UI", 10), bg="#f4f4f4", fg="#666666").pack(pady=(5, 10))

    # Lista de botones con texto, color de fondo y texto.
    botones = [
        ("1. Crear Estructura de Proyecto", "#4CAF50", "white"),
        ("2. Actualizar Textos (Word)", "#2196F3", "white"),
        ("3. Insertar Imagenes (Memoria)", "#2196F3", "white"),
        ("4. Generar Cuadros EBSS", "#2196F3", "white"),
        ("5. Convertir Documentos a PDF", "#FF9800", "white"),
        ("6. Ejecucion Completa (Fases 2, 3 y 4)", "#9C27B0", "white")
    ]

    # Crea y empaqueta cada boton.
    for texto, bg, fg in botones:
        tk.Button(root, text=texto, bg=bg, fg=fg, font=("Segoe UI", 11, "bold"), 
                  activebackground="#e0e0e0", relief="flat", cursor="hand2",
                  command=lambda t=texto: seleccionar(t), width=35, pady=4).pack(pady=3)
        
    # Inicia el bucle principal de Tkinter y retorna la eleccion y rutas.
    root.mainloop()
    return eleccion_usuario.get(), rutas_cfg["plantilla"], rutas_cfg["excel"]

# ==========================================
# MOTOR PRINCIPAL
# ==========================================
if __name__ == "__main__":
    # Limpia la pantalla de la consola al iniciar.
    os.system('cls' if os.name == 'nt' else 'clear')

    # Lanza la interfaz grafica para que el usuario seleccione la opcion y configure rutas.
    opcion_elegida, ruta_plantilla, excel_maestro = lanzar_interfaz_principal()

    # Ejecuta la fase 1 si se selecciono crear estructura de proyecto.
    if opcion_elegida == '1. Crear Estructura de Proyecto':
        fase1_crear_entorno(ruta_plantilla)
        
    # Ejecuta la fase 5 si se selecciono conversion a PDF.
    elif opcion_elegida == '5. Convertir Documentos a PDF':
        fase5_exportar_pdf()
        
    # Para las opciones que requieren una carpeta de proyecto (fases 2,3,4,6).
    elif opcion_elegida in ['2. Actualizar Textos (Word)', '3. Insertar Imagenes (Memoria)', '4. Generar Cuadros EBSS', '6. Ejecucion Completa (Fases 2, 3 y 4)']:
        # Crea una ventana Tkinter oculta para usar filedialog.
        root = tk.Tk(); root.withdraw()
        # Solicita al usuario seleccionar la carpeta del proyecto a procesar.
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta (Raiz o Subcarpeta) a procesar")
        if carpeta:
            # Normaliza la ruta de la carpeta seleccionada.
            carpeta = os.path.normpath(carpeta)
            # Extrae el codigo y nombre de empresa del nombre de la carpeta (formato: CODIGO EMPRESA).
            nombre_carpeta = os.path.basename(carpeta)
            partes = nombre_carpeta.split(' ', 1)
            if len(partes) == 2:
                codigo = partes[0]
                empresa = partes[1]
            else:
                # Si el formato no es correcto, muestra error y sale.
                pyautogui.alert("El nombre de la carpeta no tiene el formato esperado (CODIGO EMPRESA).", "Error")
                exit()
                
            # Ejecuta la fase correspondiente segun la opcion elegida.
            if opcion_elegida == '2. Actualizar Textos (Word)':
                fase2_reemplazar_textos(carpeta, excel_maestro, codigo, empresa)
            elif opcion_elegida == '3. Insertar Imagenes (Memoria)':
                fase3_insertar_imagenes(carpeta, excel_maestro, codigo, empresa)
            elif opcion_elegida == '4. Generar Cuadros EBSS':
                fase4_insertar_ebss(carpeta, excel_maestro, codigo, empresa)
            elif opcion_elegida == '6. Ejecucion Completa (Fases 2, 3 y 4)':
                # Ejecuta las fases 2, 3 y 4 en secuencia.
                fase2_reemplazar_textos(carpeta, excel_maestro, codigo, empresa)
                fase3_insertar_imagenes(carpeta, excel_maestro, codigo, empresa)
                fase4_insertar_ebss(carpeta, excel_maestro, codigo, empresa)