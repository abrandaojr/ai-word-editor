#!/usr/bin/env python3
"""
Quick test: revise a Word file in your Downloads folder via the API.
Adjust the filename as needed.
"""
import subprocess
import sys
from pathlib import Path

test_file = Path.home() / "Downloads" / "Desmatamento na Amazônia.docx"

if not test_file.exists():
    print(f"Test file not found: {test_file}")
    print("Update the filename in run_test.py to match your test document.")
    sys.exit(1)

subprocess.run(
    [sys.executable, "ai_word_editor.py", str(test_file), "--audit"],
    check=True,
)
