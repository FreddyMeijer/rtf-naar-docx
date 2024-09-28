import pypandoc
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

def select_directory():
    # Maak een root window, maar verberg het
    root = tk.Tk()
    root.withdraw()  # Verberg het hoofdvenster

    # Open dialoog voor mapselectie
    directory = filedialog.askdirectory(title="Kies de map met RTF bestanden")
    
    # Controleer of een map is geselecteerd
    if directory:
        return Path(directory)
    else:
        print("Geen map geselecteerd. Script wordt afgesloten.")
        exit()

def convert_rtf_to_docx(directory):
    # Loop door alle bestanden in de opgegeven map
    for file_path in directory.iterdir():
        if file_path.suffix == ".rtf":
            docx_file = file_path.with_suffix(".docx")
            
            try:
                # Converteer rtf naar docx met pypandoc
                pypandoc.convert_file(str(file_path), 'docx', outputfile=str(docx_file))
                print(f"Succesvol geconverteerd: {file_path.name} -> {docx_file.name}")
            except Exception as e:
                print(f"Fout bij het converteren van {file_path.name}: {e}")

if __name__ == "__main__":
    # Laat de gebruiker een map kiezen
    selected_directory = select_directory()

    # Roep de functie aan om de bestanden te converteren
    convert_rtf_to_docx(selected_directory)
