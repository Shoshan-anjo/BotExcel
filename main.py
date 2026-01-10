# main.py

import sys
from application.execute_refresh_uc import execute_refresh
from infrastructure.config_loader import ConfigLoader
from infrastructure.logger_service import LoggerService
from infrastructure.scheduler_service import SchedulerService

def main():
    config = ConfigLoader()
    logger = LoggerService(config.get_log_level()).get_logger()

    if "--scheduler" in sys.argv:
        scheduler = SchedulerService(
            logger=logger,
            execute_fn=execute_refresh
        )
        scheduler.start()
    elif "--refresh" in sys.argv:
        execute_refresh()
    else:
        from gui.app import run
        run()

if __name__ == "__main__":
    main()
