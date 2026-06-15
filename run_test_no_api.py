#!/usr/bin/env python3
"""
Quick test: generate manual prompts without using API credits.
"""
import subprocess
import sys
from pathlib import Path

input_file = Path.home() / "Downloads" / "Desmatamento na Amazônia.docx"

if not input_file.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {input_file}")

subprocess.run(
    [sys.executable, "ai_word_editor.py", str(input_file), "--manual"],
    check=True,
)
