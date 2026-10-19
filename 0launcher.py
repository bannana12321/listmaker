import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox


def run_script(script_name):
  """Führt ein Python-Skript im selben Verzeichnis aus."""
  try:
    # Pfad des aktuellen Ordners ermitteln
    current_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(current_dir, script_name)

    # Überprüfen, ob die Datei existiert
    if not os.path.exists(script_path):
      messagebox.showerror(
          "Fehler", f"Die Datei '{script_name}' wurde im Ordner nicht gefunden!"
      )
      return

    # Skript starten (ohne das GUI zu blockieren)
    subprocess.Popen([sys.executable, script_path])
  except Exception as e:
    messagebox.showerror(
        "Fehler", f"Ein unerwarteter Fehler ist aufgetreten:\n{e}"
    )


# Hauptfenster erstellen
root = tk.Tk()
root.title("Skript Launcher")
root.geometry("320x220")
root.config(bg="#f7f9fa")

# Überschrift
title_label = tk.Label(
    root,
    text="Skript-Steuerung",
    font=("Arial", 14, "bold"),
    bg="#f7f9fa",
    fg="#333333",
)
title_label.pack(pady=20)

# Button 1: AwListen Ersteller
btn_listgen = tk.Button(
    root,
    text="AwListen Ersteller",
    command=lambda: run_script("listgenerator.py"),
    font=("Arial", 11),
    bg="#4CAF50",
    fg="white",
    activebackground="#45a049",
    activeforeground="white",
    width=22,
    height=2,
    relief="flat",
    cursor="hand2",
)
btn_listgen.pack(pady=5)

# Button 2: Klassenlisten
btn_classmgr = tk.Button(
    root,
    text="Klassenlisten",
    command=lambda: run_script("sclasslistmanager.py"),
    font=("Arial", 11),
    bg="#2196F3",
    fg="white",
    activebackground="#1e88e5",
    activeforeground="white",
    width=22,
    height=2,
    relief="flat",
    cursor="hand2",
)
btn_classmgr.pack(pady=5)

# Fenster im Vordergrund halten und starten
root.mainloop()