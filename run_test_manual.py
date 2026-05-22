from pathlib import Path
import subprocess
import sys

INPUT_FILE = Path(r"C:\Users\Amintas\Downloads\Desmatamento na Amazônia.docx")

if not INPUT_FILE.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {INPUT_FILE}")

subprocess.run([sys.executable, "ai_word_editor.py", str(INPUT_FILE), "--manual"], check=True)
