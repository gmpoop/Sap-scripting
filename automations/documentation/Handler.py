from dotenv import load_dotenv
import os
import time
import win32com.client as win32
from watchdog.events import FileSystemEventHandler

EXCEL_DIRERCTORY = os.getenv("EXCEL_DIRERCTORY") 
MACRO  = os.getenv("MACRO") 

class Handler(FileSystemEventHandler):
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
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Open(EXCEL_DIRERCTORY)
            # excel.Application.Run(MACRO)
            # wb.Close(SaveChanges=False)
            # excel.Quit()
            print("Macro ejecutada correctamente.")
        except Exception as e:
            print(f"Error al ejecutar la macro: {e}")   


    def modificar_archivo(self, file_path):
        print("Modificando archivo...")

        try:
            # Lee todas las líneas del archivo
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            # Elimina las primeras 13 líneas
            lines_to_keep = lines[14:]

            # Sobrescribe el archivo con las líneas restantes
            with open(file_path, 'w', encoding='utf-8') as file:
                file.writelines(lines_to_keep)

            print("Archivo modificado exitosamente.")
        except FileNotFoundError:
            print(f"Error: El archivo '{file_path}' no existe.")
        except Exception as e:
            print(f"Error al modificar el archivo: {e}")




