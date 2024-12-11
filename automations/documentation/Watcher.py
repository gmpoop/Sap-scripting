import os
import time
from watchdog.observers import Observer 
from Handler import Handler

DIRECTORY_TO_WATCH = os.getenv("DIRECTORY_TO_WATCH")
path_to_watch = r"H:\develop\Sap-scripting\automations\documentation\txts"


class Watcher:

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        print(f"Observando el directorio: {path_to_watch}")    
        self.observer.schedule(event_handler, path_to_watch, recursive=False)
        self.observer.start()   
        try:
            while True:
                time.sleep(5)
                print("Observing...")
        except:
            self.observer.stop()
            print("Observer Stopped") 

        self.observer.join()



