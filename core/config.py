# core/config.py
from dotenv import load_dotenv
import os

load_dotenv()

class Config:
    THEME = os.getenv("THEME", "light")
    MAIL_ENABLED = os.getenv("MAIL_ENABLED") == "true"
    MAIL_SENDER = os.getenv("MAIL_SENDER", "")
    MAIL_RECIPIENTS = os.getenv("MAIL_RECIPIENTS", "").split(",")
    USE_OUTLOOK_DESKTOP = os.getenv("USE_OUTLOOK_DESKTOP") == "true"
