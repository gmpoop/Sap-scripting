import os
import time
from watchdog.observers import Observer 
from .Handler import Handler

DIRECTORY_TO_WATCH = os.getenv("DIRECTORY_TO_WATCH")

class Watcher:

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, DIRECTORY_TO_WATCH, recursive=False)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
                print("Observing...")
        except:
            self.observer.stop()
            print("Observer Stopped") 

        self.observer.join()



