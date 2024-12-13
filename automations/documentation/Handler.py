from dotenv import load_dotenv
import pythoncom
import os
import time
import win32com.client as win32
from watchdog.events import FileSystemEventHandler
import chardet
from TemplateFile import TemplateFile

EXCEL_DIRERCTORY = os.getenv("EXCEL") 
MACRO  = os.getenv("MACRO") 

class Handler(FileSystemEventHandler):
    data_count = 1
    values = []

    def __init__(self):
        self.screenshot_count = 1 
        self.tmeplate_path = ''
        self.template_file = ''

    def on_created(self, event):
        if event.is_directory:
            return None
        print(f"Evento detectado: {event.src_path}")
        
        if event.src_path.endswith(".vbs"):
            print(f"Archivo .vbs creado: {event.src_path}")
            self.change_extension(event.src_path)
        elif event.src_path.endswith(".txt"):
            print(f"Archivo .txt creado: {event.src_path}")
            self.process_file(event.src_path)

    def change_extension(self, filepath):
        new_filepath = filepath.replace(".vbs", ".txt")
        for attempt in range(5):
            try:
                os.rename(filepath, new_filepath)
                print(f"Archivo renombrado de .vbs a .txt: {new_filepath}")
                # Pausar brevemente para que el sistema registre el cambio
                time.sleep(1)
                self.process_file(new_filepath)
                return
            except PermissionError as e:
                print(f"Intento {attempt + 1}: Error al renombrar el archivo: {e}")
                time.sleep(1)  # Esperar un segundo antes de volver a intentar
        print("No se pudo renombrar el archivo después de varios intentos.")

    def process_file(self, filepath):
        print(f"Procesando archivo: {filepath}")
        
        # Reintentar hasta que el archivo exista
        for _ in range(5):
            if os.path.isfile(filepath):    
                try:
                    with open(filepath, 'r', encoding='ansi') as file:
                        content = file.read()
                        print(content)
                        # Modificar el archivo
                        self.modificar_archivo(filepath)                        

                except Exception as e:
                    print(f"Error al leer el archivo: {e}")
                break
            else:
                print(f"Esperando que el archivo esté disponible: {filepath}")
                time.sleep(1)
        else:
            print(f"El archivo no existe después de varios intentos: {filepath}")
            return

        # Ejecutar el script VBA en Excel
        try:
            pythoncom.CoInitialize()
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = True
            wb = excel.Workbooks.Open(EXCEL_DIRERCTORY)
            
            # Acceder a una hoja específica
            sheet = wb.Worksheets(1) # Asegúrate de estar accediendo a la hoja correcta 
                       
            # Editar una celda específica
            sheet.Range("E6").Value = self.template_path  # Cambia (1, 1) por la fila y columna que deseas editar
            sheet.Range("E7").Value = self.template_file  # Cambia (1, 1) por la fila y columna que deseas editar
            sheet.Range("C11").Value = 0  # Cambia (1, 1) por la fila y columna que deseas editar
            sheet.Range("B11").Value = "Default"   # Cambia (1, 1) por la fila y columna que deseas editar
            
            # Guardar y cerrar el libro
            wb.Save()
            print("Columnas editadas correctamente.")
        except Exception as e:
            print(f"Error al ejecutar la macro y editar las columnas: {e}")


    def modificar_archivo(self, file_path):   
        print("Modificando archivo...")
        self.screenshot_count = 1

        try:
            # Detectar la codificación del archivo
            with open(file_path, 'rb') as file:
                raw_data = file.read()
                result = chardet.detect(raw_data)
                encoding = result['encoding']

            # Lee todas las líneas del archivo usando la codificación detectada
            with open(file_path, 'r', encoding=encoding) as file:
                lines = file.readlines()

            # Buscar e insertar el código después de cada 'sendVKey 0'
            for i, line in enumerate(lines):
                if "session.findById(\"wnd[0]\").sendVKey 0" in line:
                    # Código para tomar la captura de pantalla       
                    screenshot_code = f"\nresponse = Doc.TakeScreenshot('#SCREEN{self.screenshot_count}#')\n" \
                                    "If response <> '' Then\n" \
                                    "    objSheet.Cells(iRow, 5) = response\n" \
                                    "    GoTo myerr\n" \
                                    "End If\n"
                    lines.insert(i + 1, screenshot_code.replace("'", '"'))
                    self.screenshot_count += 1
                    # self.values[i] = screenshot_code  
                if "session.findById('wnd[0]/usr/ctxtRMMG1-MATNR').text" in line:
                    # Código para tomar la captura de pantalla
                    data_code = f"\nresponse = Doc.TakeScreenshot('#DATA{self.data_count}#')\n" \
                                    "If response <> '' Then\n" \
                                    "    objSheet.Cells(iRow, 5) = response\n" \
                                    "    GoTo myerr\n" \
                                    "End If\n"
                

            # Elimina las primeras 13 líneas
            lines_to_keep = lines[14:]

            # Sobrescribe el archivo con las líneas restantes usando la misma codificación
            with open(file_path, 'w', encoding=encoding) as file:
                file.writelines(lines_to_keep)

            creator = TemplateFile("Workcenters", "Creacion", self.screenshot_count - 1)  
            self.template_path, self.template_file = creator.create_template()

            print("Archivo modificado exitosamente.")
        except FileNotFoundError:
            print(f"Error: El archivo '{file_path}' no existe.")
        except Exception as e:
            print(f"Error al modificar el archivo: {e}")


    def on_modified(self, event):
        if event.is_directory:
            return None
        print(f"Evento detectado: {event.src_path}")
        
        if event.src_path.endswith(".vbs"):
            print(f"Archivo .vbs modificado: {event.src_path}")
            self.change_extension(event.src_path)
        elif event.src_path.endswith(".txt"):
            print(f"Archivo .txt modificado: {event.src_path}")
            self.process_file(event.src_path)

    def create_template (self):
        print("Creando template...")
        try:
            with open("template.txt", "w") as file:
                file.write("session.findById(\"wnd[0]\").sendVKey 0")
            print("Template creado exitosamente.")
        except Exception as e:
            print(f"Error al crear el template: {e}")






