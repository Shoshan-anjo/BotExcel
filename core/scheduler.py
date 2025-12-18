# scheduler.py
import schedule
import time
from datetime import datetime

class Scheduler:
    def schedule_task(self, horario, func):
        """
        Recibe horario en 'HH:MM' y agenda la función
        """
        h, m = map(int, horario.split(":"))
        schedule.every().day.at(f"{h:02d}:{m:02d}").do(func)

    def run_forever(self):
        while True:
            schedule.run_pending()
            time.sleep(1)
