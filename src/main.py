import camelot
import fitz
import cv2
import numpy as np
import locale
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import shutil

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime

import traceback


def show_error_popup(title, error_message):
    """Display an error in a popup window."""
    popup = tk.Toplevel() if tk._default_root else tk.Tk()
    popup.withdraw()

    messagebox.showerror(title, error_message, parent=popup)

    popup.destroy()


def handle_exception(exc_type, exc_value, exc_traceback):
    """Display uncaught exceptions in a popup."""
    error_details = "".join(
        traceback.format_exception(
            exc_type, exc_value, exc_traceback
        )
    )

    show_error_popup(
        "Unexpected error",
        f"An unexpected error occurred:\n\n"
        f"{exc_value}\n\n"
        f"Details:\n{error_details}"
    )


# Redirect uncaught exceptions to the popup
import sys
sys.excepthook = handle_exception


class SimpleDatePicker(tk.Tk):
    def __init__(self):
        super().__init__()

        # --------------------------------------------------
        # Window configuration
        # --------------------------------------------------
        self.title("Planning Auto")
        self.resizable(True, True)
        self.configure(bg="#F3F6FA")

        self.selected_pdf = tk.StringVar(value="")
        self.process_all = tk.BooleanVar(value=True)

        self.today = datetime.date.today()

        # --------------------------------------------------
        # Colors
        # --------------------------------------------------
        self.COLORS = {
            "background": "#F3F6FA",
            "white": "#FFFFFF",
            "primary": "#17365D",
            "primary_hover": "#244B78",
            "accent": "#3976C5",
            "text": "#253247",
            "muted": "#718096",
            "border": "#DCE4EF",
            "success": "#23845B",
            "error": "#C0392B",
        }

        # --------------------------------------------------
        # ttk styling
        # --------------------------------------------------
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "Modern.TCombobox",
            padding=(12, 10),
            font=("Segoe UI", 11),
            foreground=self.COLORS["text"],
            fieldbackground=self.COLORS["white"],
            background=self.COLORS["white"],
            bordercolor=self.COLORS["border"],
            arrowcolor=self.COLORS["primary"],
        )

        style.map(
            "Modern.TCombobox",
            fieldbackground=[("readonly", self.COLORS["white"])],
            foreground=[("readonly", self.COLORS["text"])],
            bordercolor=[("focus", self.COLORS["accent"])],
        )

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(18, 12),
            foreground=self.COLORS["white"],
            background=self.COLORS["primary"],
            borderwidth=0,
        )

        style.map(
            "Primary.TButton",
            background=[
                ("disabled", "#AAB7C8"),
                ("active", self.COLORS["primary_hover"]),
                ("pressed", self.COLORS["primary_hover"]),
            ],
            foreground=[
                ("disabled", "#F0F2F5"),
                ("!disabled", self.COLORS["white"]),
            ],
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            padding=(12, 8),
        )

        # --------------------------------------------------
        # Variables
        # --------------------------------------------------
        self.day_var = tk.StringVar(value=str(self.today.day))
        self.status_var = tk.StringVar(
            value="Prêt à générer votre planning."
        )

        self.processing = False

        # --------------------------------------------------
        # Main container
        # --------------------------------------------------
        self.container = tk.Frame(
            self,
            bg=self.COLORS["background"],
        )
        self.container.pack(fill="both", expand=True)

        # --------------------------------------------------
        # Header
        # --------------------------------------------------
        header = tk.Frame(
            self.container,
            bg=self.COLORS["primary"],
            height=125,
        )
        header.pack(fill="x")
        header.pack_propagate(False)

        header_content = tk.Frame(
            header,
            bg=self.COLORS["primary"],
        )
        header_content.pack(
            fill="both",
            expand=True,
            padx=32,
            pady=22,
        )

        tk.Label(
            header_content,
            text="PLANNING AUTO",
            font=("Segoe UI", 21, "bold"),
            fg=self.COLORS["white"],
            bg=self.COLORS["primary"],
        ).pack(anchor="w")

        tk.Label(
            header_content,
            text="Génération automatique des plannings PDF",
            font=("Segoe UI", 10),
            fg="#D6E3F3",
            bg=self.COLORS["primary"],
        ).pack(anchor="w", pady=(5, 0))

        # --------------------------------------------------
        # Selection card
        # --------------------------------------------------
        card = tk.Frame(
            self.container,
            bg=self.COLORS["white"],
            highlightbackground=self.COLORS["border"],
            highlightthickness=1,
        )
        card.pack(
            fill="x",
            padx=28,
            pady=(24, 16),
        )

        card_content = tk.Frame(
            card,
            bg=self.COLORS["white"],
        )
        card_content.pack(
            fill="x",
            padx=24,
            pady=22,
        )

        # --------------------------------------------------
        # PDF selection
        # --------------------------------------------------
        tk.Label(
            card_content,
            text="FICHIER SOURCE",
            font=("Segoe UI", 9, "bold"),
            fg=self.COLORS["text"],
            bg=self.COLORS["white"],
        ).pack(anchor="w", pady=(0, 7))

        # Process all PDFs by default
        self.all_pdf_radio = tk.Radiobutton(
            card_content,
            text="Traiter tous les PDF du dossier",
            variable=self.process_all,
            value=True,
            font=("Segoe UI", 10),
            fg=self.COLORS["text"],
            bg=self.COLORS["white"],
            activebackground=self.COLORS["white"],
            selectcolor=self.COLORS["white"],
            command=self._toggle_pdf_selection,
        )
        self.all_pdf_radio.pack(anchor="w", pady=3)

        # Process a single PDF
        self.single_pdf_radio = tk.Radiobutton(
            card_content,
            text="Traiter un seul PDF",
            variable=self.process_all,
            value=False,
            font=("Segoe UI", 10),
            fg=self.COLORS["text"],
            bg=self.COLORS["white"],
            activebackground=self.COLORS["white"],
            selectcolor=self.COLORS["white"],
            command=self._toggle_pdf_selection,
        )
        self.single_pdf_radio.pack(anchor="w", pady=3)

        # File selection row
        self.pdf_selection_frame = tk.Frame(
            card_content,
            bg=self.COLORS["white"],
        )
        self.pdf_selection_frame.pack(fill="x", pady=(7, 15))

        self.pdf_path_label = tk.Label(
            self.pdf_selection_frame,
            text="Aucun fichier sélectionné",
            font=("Segoe UI", 9),
            fg=self.COLORS["muted"],
            bg=self.COLORS["white"],
            anchor="w",
            justify="left",
            wraplength=300,
        )
        self.pdf_path_label.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
        )

        self.browse_btn = ttk.Button(
            self.pdf_selection_frame,
            text="Parcourir...",
            command=self._browse_pdf,
            style="Secondary.TButton",
        )
        self.browse_btn.pack(side="right")

        self._toggle_pdf_selection()

        tk.Label(
            card_content,
            text="Configurer le traitement",
            font=("Segoe UI", 14, "bold"),
            fg=self.COLORS["text"],
            bg=self.COLORS["white"],
        ).pack(anchor="w")

        tk.Label(
            card_content,
            text="Choisissez le jour à générer ou traitez le mois entier.",
            font=("Segoe UI", 9),
            fg=self.COLORS["muted"],
            bg=self.COLORS["white"],
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(5, 20))

        tk.Label(
            card_content,
            text="JOUR À TRAITER",
            font=("Segoe UI", 9, "bold"),
            fg=self.COLORS["text"],
            bg=self.COLORS["white"],
        ).pack(anchor="w", pady=(0, 7))

        self.day_menu = ttk.Combobox(
            card_content,
            textvariable=self.day_var,
            values=["Mois complet"] + list(range(1, 32)),
            state="readonly",
            style="Modern.TCombobox",
            font=("Segoe UI", 11),
        )
        self.day_menu.pack(fill="x")

        # Keep the original French value used by _on_run.
        self.day_menu["values"] = (
            ["mois complet"] + list(range(1, 32))
        )

        self.day_menu.set(str(self.today.day))

        # --------------------------------------------------
        # Run button
        # --------------------------------------------------
        self.run_btn = ttk.Button(
            self.container,
            text="Générer le planning",
            command=self._on_run,
            style="Primary.TButton",
        )
        self.run_btn.pack(
            fill="x",
            padx=28,
            pady=(2, 12),
        )

        # --------------------------------------------------
        # Status area
        # --------------------------------------------------
        status_frame = tk.Frame(
            self.container,
            bg=self.COLORS["background"],
        )
        status_frame.pack(
            fill="x",
            padx=30,
            pady=(4, 0),
        )

        self.status_indicator = tk.Label(
            status_frame,
            text="●",
            font=("Segoe UI", 11),
            fg=self.COLORS["success"],
            bg=self.COLORS["background"],
        )
        self.status_indicator.pack(side="left", padx=(0, 8))

        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            fg=self.COLORS["muted"],
            bg=self.COLORS["background"],
            anchor="w",
            justify="left",
            wraplength=470,
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        # --------------------------------------------------
        # Center the window
        # --------------------------------------------------
        self.update_idletasks()

        width = 520
        height = 620

        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2

        self.geometry(f"{width}x{height}+{x}+{y}")

    # ------------------------------------------------------
    # Enable/disable controls during processing
    # ------------------------------------------------------
    def _set_processing_state(self, processing: bool):
        self.processing = processing

        state = "disabled" if processing else "readonly"

        self.day_menu.configure(state=state)

        self.run_btn.configure(
            state="disabled" if processing else "normal",
            text=(
                "Traitement en cours..."
                if processing
                else "Générer le planning"
            ),
        )

        if processing:
            self.status_indicator.configure(
                fg=self.COLORS["accent"]
            )
            self.status_var.set(
                "Traitement des fichiers PDF en cours..."
            )
        else:
            self.status_indicator.configure(
                fg=self.COLORS["success"]
            )

    def _browse_pdf(self):
        pdf_path = filedialog.askopenfilename(
            title="Choisir un planning PDF",
            initialdir=os.path.abspath("plannings_mensuels"),
            filetypes=[
                ("Fichiers PDF", "*.pdf"),
                ("Tous les fichiers", "*.*"),
            ],
            parent=self,
        )

        if pdf_path:
            self.selected_pdf.set(pdf_path)

            self.pdf_path_label.configure(
                text=os.path.basename(pdf_path),
                fg=self.COLORS["text"],
            )

    def _toggle_pdf_selection(self):
        if self.process_all.get():
            self.browse_btn.configure(state="disabled")

            self.pdf_path_label.configure(
                text="Tous les PDF du dossier seront traités",
                fg=self.COLORS["muted"],
            )
        else:
            self.browse_btn.configure(state="normal")

            if self.selected_pdf.get():
                self.pdf_path_label.configure(
                    text=os.path.basename(self.selected_pdf.get()),
                    fg=self.COLORS["text"],
                )
            else:
                self.pdf_path_label.configure(
                    text="Aucun fichier sélectionné",
                    fg=self.COLORS["muted"],
                )

    # ------------------------------------------------------
    # Run the existing processing pipeline
    # ------------------------------------------------------
    def _on_run(self):
        try:
            shutil.rmtree("crop_cell", ignore_errors=True)
            shutil.rmtree("extracted_images", ignore_errors=True)
            # Determine which day to process
            if self.day_var.get() == "mois complet":
                chosen_day = None
            else:
                chosen_day = int(self.day_var.get())

            # Determine which PDF files to process
            if self.process_all.get():
                pdf_directory = "plannings_mensuels"

                if not os.path.isdir(pdf_directory):
                    messagebox.showerror(
                        "Dossier introuvable",
                        f"Le dossier '{pdf_directory}' n'existe pas.",
                        parent=self,
                    )
                    return

                pdf_files = [
                    os.path.join(pdf_directory, filename)
                    for filename in os.listdir(pdf_directory)
                    if filename.lower().endswith(".pdf")
                ]

                if not pdf_files:
                    messagebox.showwarning(
                        "Aucun PDF",
                        "Aucun fichier PDF n'a été trouvé dans le dossier.",
                        parent=self,
                    )
                    return

            else:
                selected_pdf = self.selected_pdf.get()

                if not selected_pdf:
                    messagebox.showwarning(
                        "Aucun fichier sélectionné",
                        "Veuillez sélectionner un fichier PDF.",
                        parent=self,
                    )
                    return

                if not os.path.isfile(selected_pdf):
                    messagebox.showerror(
                        "Fichier introuvable",
                        "Le fichier sélectionné n'existe plus.",
                        parent=self,
                    )
                    return

                pdf_files = [selected_pdf]

            # Begin processing
            self._set_processing_state(True)
            self.update_idletasks()
            self.update()

            locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")

            total = len(pdf_files)

            for index, pdf_file in enumerate(pdf_files, start=1):
                filename = os.path.basename(pdf_file)

                self.status_var.set(
                    f"Traitement du fichier {index}/{total} : {filename}"
                )
                self.update_idletasks()

                main(pdf_file, chosen_day=chosen_day)

            # Preserve your existing cleanup behavior
            shutil.rmtree("crop_cell")
            shutil.rmtree("extracted_images")

        except Exception as e:
            print(e)

            self._set_processing_state(False)

            self.status_indicator.configure(
                fg=self.COLORS["error"]
            )
            self.status_var.set(
                "Une erreur est survenue pendant le traitement."
            )

            messagebox.showerror(
                "Erreur de traitement",
                f"Impossible de terminer le traitement.\n\n{e}",
                parent=self,
            )
            return

        self._set_processing_state(False)

        self.status_indicator.configure(
            fg=self.COLORS["success"]
        )
        self.status_var.set(
            f"Traitement terminé : {len(pdf_files)} fichier(s) traité(s)."
        )

        messagebox.showinfo(
            "Traitement terminé",
            f"Les plannings ont été générés avec succès.\n\n"
            f"Nombre de fichiers traités : {len(pdf_files)}",
            parent=self,
        )
        shutil.rmtree("crop_cell", ignore_errors=True)
        shutil.rmtree("extracted_images", ignore_errors=True)


def is_right_color(cell_image_path, reference_color):
    cell_image = cv2.imread(cell_image_path)
    distance = np.linalg.norm(cell_image - reference_color, axis=2)
    mask = np.uint8((distance < 40) * 255)


    return 23500 < mask.sum(), mask.sum()


def extract_images(pdf_file):
    doc = fitz.open(pdf_file)

    for i, page in enumerate(doc):
        for img_index, info in enumerate(page.get_image_info(xrefs=True)):
            xref = info['xref']
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            if not os.path.exists("extracted_images"):
                os.mkdir("extracted_images")
            with open(f"extracted_images/page{i + 1}_Image{img_index + 1}.{image_ext}", "wb") as f:
                f.write(image_bytes)


def extract_cells():
    for image_path in os.listdir("extracted_images"):
        image = cv2.imread(os.path.join("extracted_images", image_path))
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gray_inv = cv2.bitwise_not(gray)
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 10))
        vertical_lines = cv2.morphologyEx(gray_inv, cv2.MORPH_OPEN, kernel, iterations=2)

        _, thresh = cv2.threshold(vertical_lines, 95, 255, cv2.THRESH_BINARY)

        x_coords, = np.where(thresh[-1] == 255)
        x_coords = np.array(
            [x_coords[i] for i in range(len(x_coords) - 1) if x_coords[i] + 1 != x_coords[i + 1]] + [x_coords[-1]])

        cells = []
        for j in range(len(x_coords) - 1):
            x1, x2 = x_coords[j], x_coords[j + 1]
            cell = image[:, x1:x2]  # crop full height between two lines
            cells.append(cell)
            if not os.path.exists("crop_cell"):
                os.mkdir("crop_cell")
            cv2.imwrite(
                f"crop_cell/page{image_path.split('page')[-1].split('_')[0]}_cell_{image_path.split('page')[-1].split('Image')[1].split('.')[0]}_{j}.png",
                cell)


def add_to_table(table_data, story, rowHeights):
    table = Table(table_data, colWidths=[110, 150, 70, 120], rowHeights=rowHeights)
    table.setStyle(TableStyle([
        # Outer borders
        ('BOX', (0, 0), (-1, -1), 1, colors.black),

        # Vertical lines between columns
        ('LINEBEFORE', (1, 0), (1, -1), 1, colors.black),
        ('LINEBEFORE', (2, 0), (2, -1), 1, colors.black),

        # Horizontal lines only inside right part (time + name)
        ('LINEABOVE', (1, 1), (-1, -1), 0.5, colors.black),

        # Merge unit cells vertically
        ('SPAN', (0, 0), (0, len(table_data) - 1)),

        # Center vertically & add padding
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(table)
    story.append(Spacer(1, 12))


def build_document(story, day, units_list, styles, units, rowHeights):
    story.append(Paragraph(f"<b>{day}</b>", styles["Title"]))
    story.append(Spacer(1, 12))

    for unit in units_list:
        if unit not in units:
            continue
        shifts = units[unit]
        rows = []
        for i, [[hour, color_code], name] in enumerate(shifts):
            if i == 0:
                rows.append([
                    unit,
                    Paragraph(
                        f'<b><font color="rgb({color_code[-1]}, {color_code[1]}, {color_code[0]})">{hour}</font></b>',
                        styles["Normal"]),
                    name,
                    ""
                ])
            else:
                rows.append([
                    "",
                    Paragraph(
                        f'<b><font color="rgb({color_code[-1]}, {color_code[1]}, {color_code[0]})">{hour}</font></b>',
                        styles["Normal"]),
                    name,
                    ""
                ])
        if len(rows) == 0:
            continue
        add_to_table(rows, story, rowHeights)


def main(pdf_file, chosen_day=None):
    extract_images(pdf_file)
    extract_cells()

    planning = {}
    add_to_planning = {}
    name_value = {}
    color_codes = {
        'NUI': [('20:10-7:50', np.array([16, 250, 10]))],
        'REP': [('7~8:30-19~20:30', np.array([220, 242, 153]))],
        'SO': [('21:00-6:30', np.array([8, 244, 255])), ('8:15:20:15', np.array([142, 136, 255])),
               ('20:15:8:15', np.array([67, 125, 251])), ('7:45-19:45', np.array([126, 254, 135])),
               ('8:00-20:00', np.array([220, 242, 153])), #('7:30-13:30', np.array([16, 250, 10])),
               ('20:00-8:00', np.array([255, 250, 0])), ('8:30-16:30', np.array([194, 129, 0])),
               ('9:00-15:30', np.array([181, 127, 127])),
               ('7~8:30-19~20:30 | 19:45-7:45', np.array([255, 131, 255]))],
        'UNA': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'UNB': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'UNC': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'UND': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'UNE': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'UNF': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'UNG': [('7~8:30-19~20:30', np.array([220, 242, 153])), ('7:30-17:30', np.array([125, 253, 120])), ('7:30-19:30', np.array([67, 125, 251]))],
        'VIE': [('8:15-17:15', np.array([125, 253, 120])), ('9:00-17:00', np.array([8, 244, 255])),
                ('9:00-18:00', np.array([160, 255, 243]))],
    }
    mounths = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre",
               "Novembre", "Décembre"]

    tables = camelot.read_pdf(pdf_file, pages='all', line_scale=40)
    title = tables[0].df[0][0]

    dates = title.split('\n')[-2]
    upper_date = dates.split(' ')[-1]
    lower_date = dates.split(' ')[1]
    processed_date_day = int(lower_date.split('/')[0])
    day_number = 0
    last_service = None

    while processed_date_day <= int(upper_date.split('/')[0]):
        date = datetime.datetime(int(dates.split('/')[-1]), int(dates.split('/')[1]), processed_date_day)
        day_name = date.strftime("%A").capitalize()
        planning[f"{day_name} {processed_date_day} {mounths[int(dates.split('/')[1]) - 1]} {dates.split('/')[-1]}"] = {}
        for page_number, table in enumerate(tables):
            title = table.df[0][0]
            service = title.split('\n')[-1].split(' ')[0]
            if service in color_codes.keys():
                name_list = np.array(table.df[0][2::])
                name_list = name_list[name_list != '']
                if len(name_list) == 0:
                    continue
                if title.split('\n')[-1] not in planning[
                    f"{day_name} {processed_date_day} {mounths[int(dates.split('/')[1]) - 1]} {dates.split('/')[-1]}"]:
                    planning[
                        f"{day_name} {processed_date_day} {mounths[int(dates.split('/')[1]) - 1]} {dates.split('/')[-1]}"][
                        title.split('\n')[-1]] = []
                for color_code in color_codes[service]:
                    for i, name in enumerate(name_list):
                        if service == "SO" and len(name.split('\n')) > 2 and name.split('\n')[1] != 'INFIR':
                            continue
                        color_results = is_right_color(f"crop_cell/page{page_number + 1}_cell_{i + 2}_{day_number}.png",
                                          color_code[1])
                        if color_results[0]:
                            if service not in add_to_planning.keys():
                                add_to_planning[service] = {}
                                name_value[service] = {}
                            if name.split('\n')[0] not in add_to_planning[service].keys():
                                add_to_planning[service][name.split('\n')[0]] = color_code
                                name_value[service][name.split('\n')[0]] = color_results[1]
                            else:
                                if name_value[service][name.split('\n')[0]] < color_results[1]:
                                    add_to_planning[service][name.split('\n')[0]] = color_code
                                    name_value[service][name.split('\n')[0]] = color_results[1]
                if service in add_to_planning.keys():
                    planning[
                        f"{day_name} {processed_date_day} {mounths[int(dates.split('/')[1]) - 1]} {dates.split('/')[-1]}"][
                        title.split('\n')[-1]] = [[c, n] for n, c in add_to_planning[service].items()]
        add_to_planning = {}
        name_value = {}

        processed_date_day += 1
        day_number += 1
    if chosen_day:
        doc = SimpleDocTemplate(f"planning_{chosen_day}-{dates.split('/')[1]}-{dates.split('/')[-1]}.pdf",
                                pagesize=A4)
    else:
        doc = SimpleDocTemplate(f"planning_{dates.split('/')[1]}-{dates.split('/')[-1]}.pdf", pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    rowHeights = 15
    units_list = ['NUI - NUIT', 'SO - SOINS', 'UNA - UNITE A', 'UNB - UNITE B', 'UNC - UNITE C',
                  'UND - UNITE D', 'UNE - UNITE E', 'UNF - UNITE F', 'UNG - UNITE G']
    if chosen_day:
        chosen_day_name = datetime.date(int(dates.split('/')[-1]), int(dates.split('/')[1]), chosen_day).strftime("%A").capitalize()
        day = f"{chosen_day_name} {chosen_day} {mounths[int(dates.split('/')[1]) - 1]} {dates.split('/')[-1]}"
        units = planning[day]
        build_document(story, day, units_list, styles, units, rowHeights)

    else:
        for idx, (day, units) in enumerate(planning.items()):
            build_document(story, day, units_list, styles, units, rowHeights)

            if idx < len(planning) - 1:
                story.append(PageBreak())

    doc.build(story)


if __name__ == "__main__":
    app = SimpleDatePicker()
    app.mainloop()

    # for pdf_file in os.listdir('plannings_mensuels'):
    #     main(os.path.join('plannings_mensuels', pdf_file))
