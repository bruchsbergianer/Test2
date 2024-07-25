# -*- coding: utf-8 -*-
"""
Created on Fri Jul  5 10:10:25 2024

@author: harderd
"""

import zipfile
import os
import tempfile
import subprocess

def print_documents_from_zip(zip_path, sumatra_pdf_path):
    # Erstellen eines temporären Verzeichnisses zum Entpacken der ZIP-Datei
    with tempfile.TemporaryDirectory() as temp_dir:
        # Entpacken der ZIP-Datei in das temporäre Verzeichnis
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Durchsuchen des temporären Verzeichnisses nach Dateien
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Drucken der Datei, hier wird angenommen, dass es sich um PDFs handelt
                if file_path.endswith('.pdf'):
                    subprocess.run([sumatra_pdf_path, '-print-to-default', file_path])
                else:
                    print(f"Skipping non-PDF file: {file_path}")

if __name__ == "__main__":
    zip_path = 'C:\\SumatraPDF\\ZIP_FILE'  # Pfad zur ZIP-Datei
    sumatra_pdf_path = 'C:\\SumatraPDF\\SumatraPDF.exe'  # Pfad zur SumatraPDF.exe
    print_documents_from_zip(zip_path, sumatra_pdf_path)

