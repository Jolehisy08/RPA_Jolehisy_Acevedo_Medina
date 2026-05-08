import os
import shutil
import pyautogui

def fase1_crear_entorno():

    # Fase 1: Clona la carpeta PLANTILLA y renombra todos los archivos y subcarpetas.

    # 1. Interfaz para pedir los datos al usuario (Uso de alertas)
    pyautogui.alert("Iniciando Fase 1: Creación de entorno de proyecto.", "Súper RPA")
    
    codigo_proyecto = pyautogui.prompt("Introduce el CÓDIGO del proyecto (Sustituirá a 'XXXX'):\nEjemplo: P-045", "Datos del Proyecto")
    if not codigo_proyecto: 
        return # Si el usuario cancela, salimos
        
    nombre_empresa = pyautogui.prompt("Introduce el NOMBRE DE LA EMPRESA (Sustituirá a '[[EMPRESA]]'):\nEjemplo: HIDROELÉCTRICA S.L.", "Datos del Proyecto")
    if not nombre_empresa: 
        return

    # 2. Definición de rutas (¡OJO! Cambia estas rutas por las tuyas reales de Dropbox)
    # Ejemplo basado en tu captura: C:\Users\TuUsuario\Dropbox\INGENIERIA\PLANTILLA
    ruta_plantilla = r"C:\Users\joleh\Dropbox\INGENIERIA\PLANTILLA"
    
    # La nueva carpeta se creará en el mismo directorio padre que la plantilla
    directorio_padre = os.path.dirname(ruta_plantilla)
    nombre_nueva_carpeta = f"{codigo_proyecto} {nombre_empresa}"
    carpeta_destino = os.path.join(directorio_padre, nombre_nueva_carpeta)

    try:
        # 3. Clonado de la carpeta completa
        shutil.copytree(ruta_plantilla, carpeta_destino)
        
        # 4. Renombrado masivo (Magia pura)
        # Usamos topdown=False para renombrar primero los archivos que están más al fondo.
        # Si no lo hacemos así, al cambiar el nombre de una carpeta padre, la ruta de los hijos se rompe.
        for raiz, carpetas, archivos in os.walk(carpeta_destino, topdown=False):
            
            # Renombramos archivos
            for nombre_archivo in archivos:
                if "XXXX" in nombre_archivo or "[[EMPRESA]]" in nombre_archivo:
                    nuevo_nombre = nombre_archivo.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    ruta_vieja = os.path.join(raiz, nombre_archivo)
                    ruta_nueva = os.path.join(raiz, nuevo_nombre)
                    os.rename(ruta_vieja, ruta_nueva)
            
            # Renombramos subcarpetas (por si alguna carpeta también tiene las etiquetas)
            for nombre_carpeta in carpetas:
                if "XXXX" in nombre_carpeta or "[[EMPRESA]]" in nombre_carpeta:
                    nuevo_nombre = nombre_carpeta.replace("XXXX", codigo_proyecto).replace("[[EMPRESA]]", nombre_empresa)
                    ruta_vieja = os.path.join(raiz, nombre_carpeta)
                    ruta_nueva = os.path.join(raiz, nuevo_nombre)
                    os.rename(ruta_vieja, ruta_nueva)

        # 5. Feedback final
        pyautogui.alert(f"Fase 1 Completada con éxito!\nSe ha creado y configurado el proyecto:\n{nombre_nueva_carpeta}", "Éxito")

    except FileExistsError:
        pyautogui.alert("¡Error! Ya existe una carpeta con ese código y nombre de empresa.", "Error")
    except Exception as e:
        pyautogui.alert(f"Ha ocurrido un error inesperado:\n{str(e)}", "Error Crítico")

if __name__ == "__main__":
    fase1_crear_entorno()