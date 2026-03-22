# config.py
import os
from dotenv import load_dotenv

load_dotenv()

TRIGGER_WORD   = "agent:"
GEMINI_API_KEY = os.getenv("AIzaSyBOOobs8lDr46BYrubJcdFVh3zY34HQswo", "")
LOG_FILE       = "office_agent.log"
TOAST_ENABLED  = True

OFFICE_PATHS = {
    "excel":      r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "word":       r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
}
