from dotenv import load_dotenv
import os
import time
from .Watcher import Watcher


if __name__ == '__main__':
    w = Watcher()
    w.run()

