import zipfile
import os
import tempfile
import subprocess

def print_documents_from_zip(zip_path, sumatra_pdf_path):
    # Überprüfen, ob der SumatraPDF-Pfad existiert
    if not os.path.isfile(sumatra_pdf_path):
        print(f"SumatraPDF.exe nicht gefunden: {sumatra_pdf_path}")
        return

    # Überprüfen, ob die ZIP-Datei existiert
    if not os.path.isfile(zip_path):
        print(f"ZIP-Datei nicht gefunden: {zip_path}")
        return

    # Erstellen eines temporären Verzeichnisses zum Entpacken der ZIP-Datei
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Entpacken der ZIP-Datei in das temporäre Verzeichnis
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
                print(f"ZIP-Datei erfolgreich entpackt: {zip_path}")
        except zipfile.BadZipFile:
            print("Fehler: Ungültige ZIP-Datei.")
            return

        # Durchsuchen des temporären Verzeichnisses nach Dateien
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Drucken der Datei, hier wird angenommen, dass es sich um PDFs handelt
                if file_path.endswith('.pdf'):
                    try:
                        print(f"Drucken der Datei: {file_path}")
                        subprocess.run([sumatra_pdf_path, '-print-to-default', file_path], check=True)
                        print(f"Datei erfolgreich gedruckt: {file_path}")
                    except subprocess.CalledProcessError as e:
                        print(f"Fehler beim Drucken der Datei: {file_path}. Fehler: {e}")
                else:
                    print(f"Überspringen der nicht-PDF-Datei: {file_path}")

if __name__ == "__main__":
    zip_path = 'C:\\SumatraPDF\\ZIP_FILE\\Dokumente.zip'  # Pfad zur ZIP-Datei
    sumatra_pdf_path = 'C:\\SumatraPDF\\SumatraPDF.exe'  # Pfad zur SumatraPDF.exe
    print_documents_from_zip(zip_path, sumatra_pdf_path)
