#-------V4--------
import os
import tkinter as tk
from tkinter import messagebox, ttk

# ReportLab für PDF-Erstellung im Querformat
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

from pclasses import cstoragemanager as csm
from pclasses import c_datemanager as datt

# ==========================================
# KONFIGURATION (Hier anpassbar)
# ==========================================
PDF_TITLE_TEXT = (  # Erste Zeile der Überschrift
    "test1"
)
PDF_OUTPUT_DIR = "../testoutput"  # Ordner, in dem die PDFs gespeichert werden
STORAGE_DIR = "../teststorage"  # Ordner mit den Excel-Dateien

# Zuordnung der Monatsnummern zu deutschen Namen
GERMAN_MONTHS = {
    1: "Januar",
    2: "Februar",
    3: "März",
    4: "April",
    5: "Mai",
    6: "Juni",
    7: "Juli",
    8: "August",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Dezember",
}


# Übersetzungen für Wochentage
ENGLISH_TO_GERMAN_WEEKDAYS = {
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag",
}

GERMAN_TO_ENGLISH_WEEKDAYS = {
    "Montag": "Monday",
    "Dienstag": "Tuesday",
    "Mittwoch": "Wednesday",
    "Donnerstag": "Thursday",
    "Freitag": "Friday",
    "Samstag": "Saturday",
    "Sonntag": "Sunday",
}


class CalendarPdfApp:

  def __init__(self, root, storage_dir=STORAGE_DIR):
    self.root = root
    self.root.title("Excel & Kalender PDF Generator")
    self.root.geometry("650x880")

    self.storage_dir = storage_dir
    self.date_manager = datt.datetabler()

    # Aktuelles Monats-Dictionary
    self.current_month_dict = {}

    # UI-Elemente aufbauen
    self.create_widgets()

  def create_widgets(self):
    # 1. Datei-Auswahl
    frame_file = ttk.LabelFrame(self.root, text="1. Excel-Datei auswählen")
    frame_file.pack(fill="x", padx=10, pady=5)

    ttk.Label(frame_file, text="Gefundene Dateien:").pack(
        anchor="w", padx=5, pady=2
    )

    try:
      filenames = csm.getFileNames(self.storage_dir)
    except Exception as e:
      filenames = []
      messagebox.showwarning(
          "Hinweis", f"Ordner '{self.storage_dir}' nicht gefunden oder leer.\n{e}"
      )

    self.combo_files = ttk.Combobox(
        frame_file, values=filenames, state="readonly"
    )
    self.combo_files.pack(fill="x", padx=5, pady=5)

    # Event beim Ändern der Dateiauswahl binden -> Namensliste automatisch laden
    self.combo_files.bind(
        "<<ComboboxSelected>>", lambda event: self.load_names_list()
    )

    # 2. Datum-Auswahl (Jahr & Monat)
    frame_date = ttk.LabelFrame(self.root, text="2. Zeitraum wählen")
    frame_date.pack(fill="x", padx=10, pady=5)

    ttk.Label(frame_date, text="Monat (1-12):").grid(
        row=0, column=0, padx=5, pady=5
    )
    self.spin_month = ttk.Spinbox(frame_date, from_=1, to=12, width=5)
    self.spin_month.set(10)
    self.spin_month.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(frame_date, text="Jahr:").grid(row=0, column=2, padx=5, pady=5)
    self.spin_year = ttk.Spinbox(frame_date, from_=2000, to=2035, width=8)
    self.spin_year.set(2026)
    self.spin_year.grid(row=0, column=3, padx=5, pady=5)

    btn_load_month = ttk.Button(
        frame_date, text="Monat laden", command=self.load_month
    )
    btn_load_month.grid(row=0, column=4, padx=10, pady=5)

    # 3. Wochentag-Schnellauswahl (jetzt auf Deutsch)
    frame_weekday = ttk.LabelFrame(self.root, text="3. Ganzen Wochentag markieren")
    frame_weekday.pack(fill="x", padx=10, pady=5)

    self.combo_weekdays = ttk.Combobox(
        frame_weekday,
        values=[
            "Montag",
            "Dienstag",
            "Mittwoch",
            "Donnerstag",
            "Freitag",
            "Samstag",
            "Sonntag",
        ],
        state="readonly",
    )
    self.combo_weekdays.current(6)  # Standard: Sonntag
    self.combo_weekdays.pack(side="left", padx=5, pady=5)

    btn_toggle_weekday = ttk.Button(
        frame_weekday,
        text="Alle dieses Wochentags umschalten",
        command=self.toggle_weekday_action,
    )
    btn_toggle_weekday.pack(side="left", padx=5, pady=5)

    # 4. Tage-Übersicht & Einzel-Markierung
    frame_days = ttk.LabelFrame(
        self.root, text="4. Tage Übersicht (Doppelklick zum Umschalten)"
    )
    frame_days.pack(fill="x", padx=10, pady=5)

    self.tree_days = ttk.Treeview(
        frame_days,
        columns=("Tag", "Wochentag", "Markiert"),
        show="headings",
        height=5,
    )
    self.tree_days.heading("Tag", text="Tag")
    self.tree_days.heading("Wochentag", text="Wochentag")
    self.tree_days.heading("Markiert", text="Markiert")

    self.tree_days.column("Tag", width=60, anchor="center")
    self.tree_days.column("Wochentag", width=120, anchor="center")
    self.tree_days.column("Markiert", width=100, anchor="center")

    self.tree_days.pack(fill="x", padx=5, pady=5)
    self.tree_days.bind("<Double-1>", self.on_day_double_click)

    # 5. Namensliste-Vorschau aus Excel
    frame_names = ttk.LabelFrame(
        self.root, text="5. Vorschau der Namensliste aus Excel (Spalte A)"
    )
    frame_names.pack(fill="both", expand=True, padx=10, pady=5)

    frame_names_top = ttk.Frame(frame_names)
    frame_names_top.pack(fill="x", padx=5, pady=2)

    btn_refresh_list = ttk.Button(
        frame_names_top,
        text="Liste laden/aktualisieren",
        command=self.load_names_list,
    )
    btn_refresh_list.pack(side="left")

    self.lbl_names_count = ttk.Label(frame_names_top, text="Anzahl Namen: 0")
    self.lbl_names_count.pack(side="right")

    # Listbox mit Scrollbar zur Anzeige der Namen
    frame_listbox = ttk.Frame(frame_names)
    frame_listbox.pack(fill="both", expand=True, padx=5, pady=5)

    self.scrollbar_names = ttk.Scrollbar(frame_listbox)
    self.scrollbar_names.pack(side="right", fill="y")

    self.listbox_names = tk.Listbox(
        frame_listbox, yscrollcommand=self.scrollbar_names.set, height=5
    )
    self.listbox_names.pack(side="left", fill="both", expand=True)
    self.scrollbar_names.config(command=self.listbox_names.yview)

    # 6. PDF-Einstellungen (Zusatztext & Dateiname)
    frame_pdf_config = ttk.LabelFrame(
        self.root, text="6. PDF-Einstellungen (Überschrift & Dateiname)"
    )
    frame_pdf_config.pack(fill="x", padx=10, pady=5)

    ttk.Label(frame_pdf_config, text="Zusatztext (Mittlere Zeile):").grid(
        row=0, column=0, sticky="w", padx=5, pady=2
    )
    self.entry_subtitle = ttk.Entry(frame_pdf_config, width=40)
    self.entry_subtitle.insert(0, "Abteilung / Zusätzliche Infos")
    self.entry_subtitle.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

    ttk.Label(frame_pdf_config, text="PDF-Dateiname:").grid(
        row=1, column=0, sticky="w", padx=5, pady=2
    )
    self.entry_filename = ttk.Entry(frame_pdf_config, width=40)
    self.entry_filename.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
    ttk.Label(
        frame_pdf_config,
        text="(leer lassen für Standard: Monatsplan_JJJJ_MM.pdf)",
        font=("Arial", 8),
    ).grid(row=2, column=1, sticky="w", padx=5, pady=0)

    frame_pdf_config.columnconfigure(1, weight=1)

    # 7. PDF Generieren Button
    btn_pdf = ttk.Button(
        self.root,
        text="📄 PDF erstellen & speichern",
        command=self.generate_pdf,
    )
    btn_pdf.pack(fill="x", padx=10, pady=10)

    # Initiales Laden
    if filenames:
      self.combo_files.current(0)
      self.load_names_list()

    self.load_month()

  def load_names_list(self):
    """Liest die Namen aus der gewählten Excel-Datei und zeigt sie in der Listbox an."""
    self.listbox_names.delete(0, tk.END)
    selected_file = self.combo_files.get()

    if not selected_file:
      self.lbl_names_count.config(text="Anzahl Namen: 0")
      return

    try:
      col_manager = csm.ExcelColumnManager(selected_file, self.storage_dir)
      names = col_manager.getList("A")
      names = [n for n in names if n is not None]

      for idx, name in enumerate(names, start=1):
        self.listbox_names.insert(tk.END, f"{idx}. {name}")

      self.lbl_names_count.config(text=f"Anzahl Namen: {len(names)}")
    except Exception as e:
      self.lbl_names_count.config(text="Fehler beim Laden")
      messagebox.showerror("Fehler beim Lesen der Excel-Datei", str(e))

  def load_month(self):
    """Erstellt das Monat-Dict über d_datemanager."""
    try:
      m = int(self.spin_month.get())
      y = int(self.spin_year.get())
      self.current_month_dict = self.date_manager.month_dict(m, y)
      self.refresh_tree()
    except ValueError:
      messagebox.showerror(
          "Fehler", "Bitte gültiges Jahr und Monat eingeben."
      )

  def refresh_tree(self):
    """Aktualisiert die Darstellung der Tage in der Tabelle (auf Deutsch übersetzt)."""
    for item in self.tree_days.get_children():
      self.tree_days.delete(item)

    for day_str, info in self.current_month_dict.items():
      marked_str = "JA (Gelb)" if info["marked"] else "Nein"
      # Wochentag aus dem Dictionary ins Deutsche übersetzen
      german_weekday = ENGLISH_TO_GERMAN_WEEKDAYS.get(
          info["weekday"], info["weekday"]
      )
      self.tree_days.insert(
          "", "end", iid=day_str, values=(day_str, german_weekday, marked_str)
      )

  def on_day_double_click(self, event):
    """Umschalten eines einzelnen Tages per Doppelklick."""
    selected = self.tree_days.selection()
    if not selected:
      return
    day_num = int(selected[0])
    self.current_month_dict = self.date_manager.toggle_single_day(
        self.current_month_dict, day_num
    )
    self.refresh_tree()

  def toggle_weekday_action(self):
    """Schaltet alle Tage eines bestimmten Wochentags um (Mapping von Deutsch auf Englisch)."""
    selected_german_day = self.combo_weekdays.get()
    selected_english_day = GERMAN_TO_ENGLISH_WEEKDAYS.get(selected_german_day)

    if selected_english_day and self.current_month_dict:
      self.current_month_dict = self.date_manager.toggle_days(
          self.current_month_dict, selected_english_day
      )
      self.refresh_tree()

  def generate_pdf(self):
    """Erstellt die PDF-Tabelle im Querformat mit linksbündigem, dreizeiligem Titel."""
    selected_file = self.combo_files.get()
    if not selected_file:
      messagebox.showerror(
          "Fehler", "Bitte wähle zuerst eine Excel-Datei aus."
      )
      return

    # Namen aus Spalte A der Excel-Datei abrufen
    try:
      col_manager = csm.ExcelColumnManager(selected_file, self.storage_dir)
      names = col_manager.getList("A")
      names = [n for n in names if n is not None]
    except Exception as e:
      messagebox.showerror("Fehler beim Lesen der Excel-Datei", str(e))
      return

    if not names:
      messagebox.showwarning(
          "Warnung", "Keine Namen in Spalte A der gewählten Datei gefunden."
      )

    # Zielverzeichnis für PDFs erstellen, falls es nicht existiert
    if not os.path.exists(PDF_OUTPUT_DIR):
      os.makedirs(PDF_OUTPUT_DIR)

    month_num = int(self.spin_month.get())
    year_num = int(self.spin_year.get())
    month_name_german = GERMAN_MONTHS.get(month_num, str(month_num))

    # Benutzerdefinierten Dateinamen auswerten oder Standard nehmen
    custom_filename = self.entry_filename.get().strip()
    if custom_filename:
      if not custom_filename.endswith(".pdf"):
        custom_filename += ".pdf"
      pdf_filename_full = os.path.join(PDF_OUTPUT_DIR, custom_filename)
    else:
      pdf_filename_full = os.path.join(
          PDF_OUTPUT_DIR, f"Monatsplan_{year_num}_{month_num:02d}.pdf"
      )

    # PDF Dokument im Querformat (A4 Landscape)
    doc = SimpleDocTemplate(
        pdf_filename_full,
        pagesize=landscape(A4),
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15,
    )

    story = []

    # Styles für die linksbündige Überschrift definieren
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Heading1"],
        fontSize=13,
        alignment=0,  # Linksbündig
        spaceAfter=2,
    )
    middle_style = ParagraphStyle(
        "CustomMiddle",
        parent=styles["Normal"],
        fontSize=10,
        alignment=0,  # Linksbündig
        spaceAfter=2,
    )
    subtitle_style = ParagraphStyle(
        "CustomSubtitle",
        parent=styles["Normal"],
        fontSize=10,
        alignment=0,  # Linksbündig
        spaceAfter=10,
    )

    # 1. Zeile (Haupttitel)
    story.append(Paragraph(PDF_TITLE_TEXT, title_style))

    # 2. Zeile (Zusatztext aus dem GUI-Eingabefeld dazwischen)
    custom_subtext = self.entry_subtitle.get().strip()
    if custom_subtext:
      story.append(Paragraph(custom_subtext, middle_style))

    # 3. Zeile (Monat und Jahr)
    story.append(
        Paragraph(f"{month_name_german} {year_num}", subtitle_style)
    )

    # Tage bestimmen
    num_days = len(self.current_month_dict)
    day_keys = [str(i) for i in range(1, num_days + 1)]

    # Kopfzeile der Tabelle
    header_row = ["Name / Tag"] + day_keys
    table_data = [header_row]

    # Zeilen für jeden Namen
    for name in names:
      row = [str(name)] + [""] * num_days
      table_data.append(row)

    # Tabellen-Styling
    t_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]

    # Gelbe Markierung für markierte Tage anwenden
    for col_idx, day_str in enumerate(day_keys, start=1):
      if self.current_month_dict.get(day_str, {}).get("marked", False):
        t_style.append(
            ("BACKGROUND", (col_idx, 0), (col_idx, -1), colors.yellow)
        )

    # Spaltenbreiten berechnen
    total_width = landscape(A4)[0] - 30
    name_col_width = 100
    day_col_width = (total_width - name_col_width) / max(num_days, 1)
    col_widths = [name_col_width] + [day_col_width] * num_days

    pdf_table = Table(table_data, colWidths=col_widths)
    pdf_table.setStyle(TableStyle(t_style))

    story.append(pdf_table)

    try:
      doc.build(story)
      messagebox.showinfo(
          "Erfolg",
          f"Die PDF wurde erfolgreich im Ordner gespeichert:\n{os.path.abspath(pdf_filename_full)}",
      )
    except Exception as e:
      messagebox.showerror("Fehler bei PDF-Erstellung", str(e))


if __name__ == "__main__":
  root = tk.Tk()
  app = CalendarPdfApp(root)
  root.mainloop()