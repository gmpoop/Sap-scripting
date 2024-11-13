import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32com.client as win32

class Watcher:
    DIRECTORY_TO_WATCH = "H:/develop/Sap-scripting/automations/documentation/txts"

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DIRECTORY_TO_WATCH, recursive=False)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            self.observer.stop()
            print("Observer Stopped")

        self.observer.join()

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
        print(f"Archivo modificado de .vbs a .txt: {filepath}")
        new_filepath = filepath.replace(".vbs", ".txt")
        self.process_file(new_filepath)

    def process_file(self, filepath):
        print(f"Procesando archivo: {filepath}")
        # Verificar si el archivo existe
        if os.path.isfile(filepath):
            # Leer el archivo en formato ANSI (Western Europe)
            try:
                with open(filepath, 'r', encoding='ansi') as file:
                    content = file.read()
                    print(content)
            except Exception as e:
                print(f"Error al leer el archivo: {e}")
        else:
            print(f"El archivo no existe: {filepath}")

        # Ejecutar el script VBA en Excel
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Open(r"ruta/a/tu/macro.xlsm")
            excel.Application.Run("NombreDeTuMacro")
            wb.Close(SaveChanges=False)
            excel.Quit()
            print("Macro ejecutada correctamente.")
        except Exception as e:
            print(f"Error al ejecutar la macro: {e}")

if __name__ == '__main__':
    w = Watcher()
    w.run()
