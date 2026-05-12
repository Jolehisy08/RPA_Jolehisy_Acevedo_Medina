import os
import shutil
import pyautogui
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import warnings
import time 

# ==========================================
# FUNCIONES "RADAR" (BÚSQUEDA DINÁMICA)
# ==========================================
def buscar_archivos(carpeta_base, nombres_buscar):
    encontrados = []
    for raiz, _, archivos in os.walk(carpeta_base):
        for archivo in archivos:
            if archivo in nombres_buscar:
                encontrados.append(os.path.normpath(os.path.join(raiz, archivo)))
    return encontrados

def encontrar_excel_ebss(carpeta_base, codigo, empresa):
    nombre_excel = f"{codigo} TABLAS EBSS {empresa}.xlsx"
    for raiz, _, archivos in os.walk(carpeta_base):
        if nombre_excel in archivos:
            return os.path.normpath(os.path.join(raiz, nombre_excel))
    
    ruta = carpeta_base
    nombre_proyecto = f"{codigo} {empresa}"
    while os.path.basename(ruta) != nombre_proyecto and os.path.dirname(ruta) != ruta:
        ruta = os.path.dirname(ruta)
    
    for raiz, _, archivos in os.walk(ruta):
        if nombre_excel in archivos:
            return os.path.normpath(os.path.join(raiz, nombre_excel))
            
    return None

# ==========================================
# FASE 1: CLONAR Y RENOMBRAR
# ==========================================
def fase1_crear_entorno():
    codigo_proyecto = pyautogui.prompt("Introduce el CÓDIGO del proyecto:", "Datos del Proyecto")
    if not codigo_proyecto: return None 
        
    nombre_empresa = pyautogui.prompt("Introduce el NOMBRE DE LA EMPRESA:", "Datos del Proyecto")
    if not nombre_empresa: return None

    ruta_plantilla = r"C:\Users\joleh\Dropbox\INGENIERIA\PLANTILLA"
    
    directorio_padre = os.path.normpath(os.path.dirname(ruta_plantilla))
    nombre_nueva_carpeta = f"{codigo_proyecto} {nombre_empresa}"
    carpeta_destino = os.path.join(directorio_padre, nombre_nueva_carpeta)

    print("\n" + "="*50)
    print(f"🚀 INICIANDO FASE 1: CLONANDO PROYECTO")
    print("="*50)
    print(f">> Creando carpeta: {nombre_nueva_carpeta}...")

    try:
        shutil.copytree(ruta_plantilla, carpeta_destino)
        for raiz, carpetas, archivos in os.walk(carpeta_destino, topdown=False):
            for nombre_archivo in archivos:
                if "XXXX" in nombre_archivo or "[[EMPRESA]]" in nombre_archivo:
                    nuevo_nombre = nombre_archivo.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    os.rename(os.path.join(raiz, nombre_archivo), os.path.join(raiz, nuevo_nombre))
            
            for nombre_carpeta in carpetas:
                if "XXXX" in nombre_carpeta or "[[EMPRESA]]" in nombre_carpeta:
                    nuevo_nombre = nombre_carpeta.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    os.rename(os.path.join(raiz, nombre_carpeta), os.path.join(raiz, nuevo_nombre))

        print("✔ Fase 1 completada con éxito.")
        pyautogui.alert("¡Estructura de carpetas creada con éxito!", "Completado")
        return True 

    except FileExistsError:
        pyautogui.alert("¡Error! Ya existe una carpeta con ese código y nombre de empresa.", "Error")
        return False
    except Exception as e:
        pyautogui.alert(f"Ha ocurrido un error inesperado:\n{str(e)}", "Error Crítico")
        return False

# ==========================================
# FASE 2: REEMPLAZAR TEXTOS
# ==========================================
def fase2_reemplazar_textos(carpeta_proyecto, ruta_excel, codigo, empresa):
    print("\n" + "="*50)
    print(f"🚀 INICIANDO FASE 2: REEMPLAZO DE ETIQUETAS DE TEXTO")
    print("="*50)

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

        for pestana_excel, lista_nombres in archivos_objetivo.items():
            try:
                df_textos = pd.read_excel(ruta_excel, sheet_name=pestana_excel)
            except Exception as e:
                continue 

            rutas_encontradas = buscar_archivos(carpeta_proyecto, lista_nombres)

            for ruta_absoluta in rutas_encontradas:
                print(f"   ⏳ Modificando textos: {os.path.basename(ruta_absoluta)}")
                doc = word.Documents.Open(ruta_absoluta)

                for index, row in df_textos.iterrows():
                    etiqueta = str(row.iloc[0]).strip() 
                    valor = str(row.iloc[1]).strip()    
                    if valor.lower() == 'nan' or valor == '': continue

                    for story in doc.StoryRanges:
                        story.Find.Execute(FindText=etiqueta, ReplaceWith=valor, Replace=2)
                        while story.NextStoryRange:
                            story = story.NextStoryRange
                            story.Find.Execute(FindText=etiqueta, ReplaceWith=valor, Replace=2)

                doc.Save(); doc.Close()
                print(f"   ✔ COMPLETADO: {os.path.basename(ruta_absoluta)}")

        word.Quit()
    except Exception as e:
        print(f"❌ Ocurrió un error en la Fase 2:\n{str(e)}")
        try: word.Quit() 
        except: pass

# ==========================================
# FASE 3: INSERCIÓN DE IMÁGENES
# ==========================================
def fase3_insertar_imagenes(carpeta_proyecto, ruta_excel, codigo, empresa):
    print("\n" + "="*50)
    print(f"🚀 INICIANDO FASE 3: IMÁGENES DB SI")
    print("="*50)

    nombre_memoria = f"03 {codigo} MEMORIA ACT {empresa}.docx"
    rutas_encontradas = buscar_archivos(carpeta_proyecto, [nombre_memoria])

    if not rutas_encontradas:
        print("   ℹ️ No se encontró la memoria de actividad en la carpeta seleccionada.")
        return

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df_tablas = pd.read_excel(ruta_excel, sheet_name='MEMORIA TABLAS')
    except: return

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False 

        for ruta_absoluta in rutas_encontradas:
            print(f"   📄 Procesando Word: {os.path.basename(ruta_absoluta)}")
            doc = word.Documents.Open(ruta_absoluta)
            ancho_util = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin

            for index, row in df_tablas.iterrows():
                etiqueta = str(row['Etiqueta']).strip()
                ruta_imagen = str(row['Dirección']).strip()
                ruta_imagen = os.path.normpath(ruta_imagen)

                if ruta_imagen.lower().endswith(('.png', '.jpg', '.jpeg')) and os.path.exists(ruta_imagen):
                    word.Selection.Find.ClearFormatting()
                    if word.Selection.Find.Execute(FindText=etiqueta):
                        word.Selection.Delete()
                        imagen = word.Selection.InlineShapes.AddPicture(FileName=ruta_imagen)
                        imagen.LockAspectRatio = -1 
                        imagen.Width = ancho_util
                        imagen.Range.ParagraphFormat.Alignment = 1
                        print(f"      ✔ Imagen insertada: {etiqueta}")

            doc.Save(); doc.Close()
            
        word.Quit()
    except Exception as e:
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
    except Exception as e:
        try: chart_obj.Delete() 
        except: pass
        return False

def fase4_insertar_ebss(carpeta_proyecto, ruta_excel, codigo, empresa):
    print("\n" + "="*50)
    print(f"🚀 INICIANDO FASE 4: CUADROS DE OFICIOS EBSS")
    print("="*50)

    RANGO_CUADRO_1 = "A1:I37"
    RANGO_CUADRO_2 = "K39:P73"

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df_oficios = pd.read_excel(ruta_excel, sheet_name='PERSONAL_EBSS')
    except Exception as e:
        print("⚠️ Pestaña [PERSONAL_EBSS] no encontrada.")
        return

    oficios_seleccionados = []
    for index, row in df_oficios.iterrows():
        oficio = str(row.iloc[0]).strip()
        marcado = str(row.iloc[1]).strip().upper()
        if marcado == 'SI':
            oficios_seleccionados.append(oficio)

    if not oficios_seleccionados:
        print("   ℹ️ No hay oficios marcados con 'SI'.")
        return

    nombres_destino = [f"05 {codigo} EBSS {empresa}.doc", f"07 {codigo} EBSS {empresa}.doc"]
    docs_ebss_destino = buscar_archivos(carpeta_proyecto, nombres_destino)

    if not docs_ebss_destino:
        print("   ℹ️ No se encontraron documentos EBSS en la carpeta seleccionada.")
        return

    ruta_origen_ebss = encontrar_excel_ebss(carpeta_proyecto, codigo, empresa)
    if not ruta_origen_ebss:
        print(f"❌ Error: No se encontró el Excel de Tablas EBSS.")
        return

    ruta_temp = os.path.normpath(os.path.join(carpeta_proyecto, "temp_cuadro.png"))

    try:
        excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        
        print(f"   ⏳ Abriendo Excel: {os.path.basename(ruta_origen_ebss)}")
        wb_tablas = excel_app.Workbooks.Open(ruta_origen_ebss)

        for ruta_doc in docs_ebss_destino:
            print(f"   📄 Procesando Word: {os.path.basename(ruta_doc)}")
            doc = word.Documents.Open(ruta_doc)
            
            word.Selection.Find.ClearFormatting()
            if word.Selection.Find.Execute(FindText="[[TABLAS_OFICIOS]]"):
                word.Selection.Delete() 

                primer_paso = True
                for oficio in oficios_seleccionados:
                    try:
                        ws = wb_tablas.Sheets(oficio)
                        ws.Activate() 
                        
                        if not primer_paso: word.Selection.InsertBreak(Type=7)
                            
                        if exportar_rango_a_imagen(ws, RANGO_CUADRO_1, ruta_temp) and os.path.exists(ruta_temp):
                            shape1 = word.Selection.InlineShapes.AddPicture(FileName=ruta_temp, LinkToFile=False, SaveWithDocument=True)
                            shape1.LockAspectRatio = False 
                            shape1.Width = 18 * 28.35  
                            shape1.Height = 20 * 28.35 
                            shape1.Range.ParagraphFormat.Alignment = 1
                            word.Selection.Collapse(Direction=0)
                            word.Selection.TypeParagraph()
                            print(f"      ✔ {oficio} - Cuadro 1 OK")
                            os.remove(ruta_temp) 

                        time.sleep(1) 
                        word.Selection.InsertBreak(Type=7)
                        
                        if exportar_rango_a_imagen(ws, RANGO_CUADRO_2, ruta_temp) and os.path.exists(ruta_temp):
                            shape2 = word.Selection.InlineShapes.AddPicture(FileName=ruta_temp, LinkToFile=False, SaveWithDocument=True)
                            shape2.LockAspectRatio = False 
                            shape2.Width = 18 * 28.35  
                            shape2.Height = 20 * 28.35 
                            shape2.Range.ParagraphFormat.Alignment = 1
                            word.Selection.Collapse(Direction=0)
                            word.Selection.TypeParagraph()
                            print(f"      ✔ {oficio} - Cuadro 2 OK")
                            os.remove(ruta_temp)
                        
                        primer_paso = False
                    except Exception as e:
                        print(f"      ❌ Error en oficio '{oficio}': {e}")
            else:
                print(f"   ⚠️ No se encontró la etiqueta [[TABLAS_OFICIOS]]")

            doc.Save(); doc.Close()

        wb_tablas.Close(False); excel_app.Quit(); word.Quit()

    except Exception as e:
        print(f"❌ Error crítico Fase 4: {e}")
        try: excel_app.Quit(); word.Quit()
        except: pass

# ==========================================
# FASE 5: EXPORTACIÓN MASIVA A PDF
# ==========================================
def fase5_exportar_pdf():
    pyautogui.alert("Selecciona la carpeta donde están los documentos Word que deseas convertir.\n(Ej: la carpeta '01 Documentos')", "Conversor a PDF")
    
    root = tk.Tk(); root.withdraw()
    carpeta_origen = filedialog.askdirectory(title="Selecciona la carpeta de los Documentos Word")
    
    if not carpeta_origen:
        print("❌ Operación cancelada. No se seleccionó ninguna carpeta.")
        return

    carpeta_origen = os.path.normpath(carpeta_origen)
    carpeta_padre = os.path.dirname(carpeta_origen)
    carpeta_pdf = os.path.join(carpeta_padre, "03 PDF")

    if not os.path.exists(carpeta_pdf):
        os.makedirs(carpeta_pdf)
        print(f"📂 Creada nueva carpeta destino: {carpeta_pdf}")

    archivos_word = [f for f in os.listdir(carpeta_origen) if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
    
    if not archivos_word:
        pyautogui.alert("No se encontraron documentos Word en la carpeta seleccionada.", "Sin archivos")
        return

    print("\n" + "="*50)
    print(f"🚀 INICIANDO FASE 5: CONVERSIÓN A PDF")
    print("="*50)
    print(f"   📂 Origen:  {carpeta_origen}")
    print(f"   🎯 Destino: {carpeta_pdf}\n")

    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        for archivo in archivos_word:
            ruta_doc = os.path.join(carpeta_origen, archivo)
            nombre_pdf = os.path.splitext(archivo)[0] + ".pdf"
            ruta_pdf = os.path.join(carpeta_pdf, nombre_pdf)

            print(f"   ⏳ Convirtiendo: {archivo} ...")
            try:
                doc = word.Documents.Open(ruta_doc)
                # 17 = Formato Constante para PDF en Word
                doc.ExportAsFixedFormat(OutputFileName=ruta_pdf, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0, CreateBookmarks=1)
                doc.Close(False)
                print(f"      ✔ Guardado: {nombre_pdf}")
            except Exception as e:
                print(f"      ❌ Error al convertir {archivo}: {e}")

        word.Quit()
        
        print("\n" + "="*50)
        print("🏁 ¡CONVERSIÓN COMPLETADA!")
        print("="*50 + "\n")
        pyautogui.alert(f"¡Todos los documentos han sido convertidos!\n\nSe han guardado en:\n{carpeta_pdf}", "PDFs Generados")

    except Exception as e:
        print(f"❌ Error crítico en Fase 5: {e}")
        try: word.Quit()
        except: pass

# ==========================================
# INTERFAZ GRÁFICA DEL MENÚ (TKINTER)
# ==========================================
def lanzar_interfaz_principal():
    root = tk.Tk()
    root.title("RPA Proyectos - Panel de Control")
    root.geometry("450x420")
    root.configure(bg="#f4f4f4")
    
    # Centrar ventana en la pantalla
    window_width, window_height = 450, 420
    screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
    position_top, position_right = int(screen_height/2 - window_height/2), int(screen_width/2 - window_width/2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    eleccion_usuario = tk.StringVar()
    
    def seleccionar(opcion):
        eleccion_usuario.set(opcion)
        root.destroy()

    # Textos de la cabecera
    tk.Label(root, text="Asistente RPA de Proyectos", font=("Segoe UI", 16, "bold"), bg="#f4f4f4", fg="#333333").pack(pady=(20, 5))
    tk.Label(root, text="Selecciona el módulo que deseas ejecutar:", font=("Segoe UI", 11), bg="#f4f4f4", fg="#666666").pack(pady=(0, 15))

    # Definición de botones (Texto, Color de Fondo, Color de Letra)
    botones = [
        ("1. Crear Estructura de Proyecto", "#4CAF50", "white"),
        ("2. Actualizar Textos (Word)", "#2196F3", "white"),
        ("3. Insertar Imágenes (Memoria)", "#2196F3", "white"),
        ("4. Generar Cuadros EBSS", "#2196F3", "white"),
        ("5. Convertir Documentos a PDF", "#FF9800", "white"),
        ("6. Ejecución Completa (Fases 2, 3 y 4)", "#9C27B0", "white")
    ]

    for texto, bg, fg in botones:
        tk.Button(root, text=texto, bg=bg, fg=fg, font=("Segoe UI", 11, "bold"), 
                  activebackground="#e0e0e0", relief="flat", cursor="hand2",
                  command=lambda t=texto: seleccionar(t), width=35, pady=6).pack(pady=4)
        
    root.mainloop()
    return eleccion_usuario.get()

# ==========================================
# MOTOR PRINCIPAL
# ==========================================
if __name__ == "__main__":
    excel_maestro = r"C:\Users\joleh\Desktop\MII\SEGUNDO\2Cuatrimestre\IIA\RPA_Jolehisy_Acevedo_Medina\Configuración_memoria.xlsx"
    os.system('cls' if os.name == 'nt' else 'clear')

    opcion_elegida = lanzar_interfaz_principal()

    if opcion_elegida == '1. Crear Estructura de Proyecto':
        fase1_crear_entorno()
        
    elif opcion_elegida == '5. Convertir Documentos a PDF':
        # La Fase 5 tiene su propia lógica de selección de carpetas y no necesita el Excel Maestro
        fase5_exportar_pdf()
        
    elif opcion_elegida in ['2. Actualizar Textos (Word)', '3. Insertar Imágenes (Memoria)', '4. Generar Cuadros EBSS', '6. Ejecución Completa (Fases 2, 3 y 4)']:
        root = tk.Tk(); root.withdraw()
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta (Raíz o Subcarpeta) a procesar")
        if carpeta:
            carpeta = os.path.normpath(carpeta)
            
            # Autocompletado del código y la empresa
            ruta_analisis = carpeta
            cod_sug, emp_sug = "", ""
            while ruta_analisis != os.path.dirname(ruta_analisis):
                nombre_c = os.path.basename(ruta_analisis)
                partes = nombre_c.split(' ', 1)
                if len(partes) > 1 and partes[0].isdigit():
                    cod_sug = partes[0]
                    emp_sug = partes[1]
                    break
                ruta_analisis = os.path.dirname(ruta_analisis)

            cod = pyautogui.prompt("Confirma CÓDIGO del proyecto:", default=cod_sug)
            emp = pyautogui.prompt("Confirma EMPRESA:", default=emp_sug)
            
            if cod and emp:
                if opcion_elegida == '2. Actualizar Textos (Word)':
                    fase2_reemplazar_textos(carpeta, excel_maestro, cod, emp)
                elif opcion_elegida == '3. Insertar Imágenes (Memoria)':
                    fase3_insertar_imagenes(carpeta, excel_maestro, cod, emp)
                elif opcion_elegida == '4. Generar Cuadros EBSS':
                    fase4_insertar_ebss(carpeta, excel_maestro, cod, emp)
                elif opcion_elegida == '6. Ejecución Completa (Fases 2, 3 y 4)':
                    fase2_reemplazar_textos(carpeta, excel_maestro, cod, emp)
                    fase3_insertar_imagenes(carpeta, excel_maestro, cod, emp)
                    fase4_insertar_ebss(carpeta, excel_maestro, cod, emp)
                
                print("\n" + "="*50)
                print("🏁 ¡PROCESO TOTALMENTE COMPLETADO!")
                print("="*50 + "\n")
                pyautogui.alert("¡Módulo finalizado con éxito!", "RPA Completado")
        else:
            print("❌ Operación cancelada. No se seleccionó ninguna carpeta.")
    else:
        print("❌ Operación cancelada o ventana cerrada.")