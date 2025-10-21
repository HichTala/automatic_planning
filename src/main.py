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
from tkinter import ttk, messagebox
import datetime
import calendar


class SimpleDatePicker(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Planning Auto")
        self.resizable(False, False)
        pad = {"padx": 8, "pady": 6}

        today = datetime.date.today()
        self.year_range = list(range(today.year - 50, today.year + 11))

        # Variables
        self.day_var = tk.StringVar(value=str(today.day))

        # Day
        ttk.Label(self, text="Jour:").grid(row=2, column=0, **pad, sticky="w")
        self.day_menu = ttk.OptionMenu(self, self.day_var, self.day_var.get(),
                                       *['mois complet'] + list(range(1, 32)))
        self.day_menu.grid(row=2, column=1, **pad, sticky="ew")

        # Run button
        self.run_btn = ttk.Button(self, text="Go", command=self._on_run)
        self.run_btn.grid(row=3, column=0, columnspan=2, pady=(10, 10))

        # Center the window
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws // 2) - (w // 2);
        y = (hs // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def _update_day_menu(self):
        year = int(self.year_var.get())
        month = int(self.month_var.get())
        days = self._days_for(year, month)
        menu = self.day_menu["menu"]
        menu.delete(0, "end")
        for d in days:
            menu.add_command(label=d, command=lambda v=d: self.day_var.set(v))
        if self.day_var.get() > days[-1]:
            self.day_var.set(days[-1])

    def _on_year_or_month_change(self, _=None):
        self._update_day_menu()

    def _set_processing_state(self, processing: bool):
        """Enable/disable controls and change button text."""
        state = "disabled" if processing else "normal"
        try:
            self.year_menu.configure(state=state)
            self.month_menu.configure(state=state)
            self.day_menu.configure(state=state)
        except Exception:
            pass
        self.run_btn.configure(text="Traitement en cours..." if processing else "Run", state=state)

    def _on_run(self):
        """
        Simple (non-threaded) run:
        - disable UI and change label
        - force UI update so the change is visible
        - call user_function (runs in main thread; GUI will be unresponsive while it runs)
        - re-enable UI and show result
        """
        if self.day_var.get() == 'mois complet':
            chosen_day = None
        else:
            chosen_day = int(self.day_var.get())

        # Disable UI and change label
        self._set_processing_state(True)
        # Force the GUI to process the change so user sees the "Processing..." label
        self.update_idletasks()
        self.update()

        # Call the user function (this will block the GUI if it takes long)
        try:
            locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
            for pdf_file in os.listdir('plannings_mensuels'):
                main(os.path.join('plannings_mensuels', pdf_file), chosen_day=chosen_day)
            shutil.rmtree("crop_cell")
            shutil.rmtree("extracted_images")
        except Exception as e:
            # Re-enable UI before showing error
            self._set_processing_state(False)
            return

        # Re-enable UI and show result
        self._set_processing_state(False)


def is_right_color(cell_image_path, reference_color):
    try:
        cell_image = cv2.imread(cell_image_path)
        distance = np.linalg.norm(cell_image - reference_color, axis=2)
        mask = np.uint8((distance < 50) * 255)
    except TypeError:
        breakpoint()

    return 20000 < mask.sum()


def extract_images(pdf_file):
    doc = fitz.open(pdf_file)

    for i, page in enumerate(doc):
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            name = img[7]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            if not os.path.exists("extracted_images"):
                os.mkdir("extracted_images")
            with open(f"extracted_images/page{i + 1}_{name}.{image_ext}", "wb") as f:
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
    table = Table(table_data, colWidths=[160, 70, 140, 120], rowHeights=rowHeights)
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
    color_codes = {
        'ARO': [('7:30-19:30', np.array([180, 120, 10])), ('8:30-20:30', np.array([130, 130, 250])),
                ('9:00-16:00', np.array([250, 255, 120]))],
        'LAV': [('7:30-19:30', np.array([180, 120, 10])), ('8:30-20:30', np.array([130, 130, 250]))],
        'ORA': [('7:30-19:30', np.array([180, 120, 10])), ('8:30-20:30', np.array([130, 130, 250])),
                ('9:00-16:00', np.array([250, 255, 120]))],
        'ROS': [('7:30-19:30', np.array([180, 120, 10])), ('8:30-20:30', np.array([130, 130, 250])),
                ('7:30-14:30', np.array([140, 255, 255])), ('7:30-14:30', np.array([0, 130, 250]))],
        'TUL': [('7:30-19:30', np.array([180, 120, 10])), ('8:30-20:30', np.array([130, 130, 250])),
                ('9:00-16:00', np.array([250, 255, 120]))],
        'VLT': [('7:30-14:30', np.array([180, 120, 10]))],
        'INF': [('8:00-20:00', np.array([45, 240, 230])), ('20:00-8:00', np.array([120, 20, 200])),
                ('9:00-16:00', np.array([250, 255, 120]))],
        'NUI': [('9:00-12:00', np.array([120, 120, 70])), ('13:00-16:30', np.array([140, 255, 50])),
                ('8:30-12:00', np.array([0, 130, 250])), ('20:00-8:00', np.array([120, 20, 200])),
                ('20:00-8:00', np.array([180, 120, 130]))],
        'ASI': [('h1-h4/lc-lg', np.array([180, 120, 10])), ('h2-h5', np.array([250, 5, 250])),
                ('la-lh', np.array([180, 125, 255]))]
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
                        if is_right_color(f"crop_cell/page{page_number + 1}_cell_{i + 2}_{day_number}.png",
                                          color_code[1]):
                            planning[
                                f"{day_name} {processed_date_day} {mounths[int(dates.split('/')[1]) - 1]} {dates.split('/')[-1]}"][
                                title.split('\n')[-1]].append([color_code, name.split('\n')[0]])

        processed_date_day += 1
        day_number += 1
    if chosen_day:
        doc = SimpleDocTemplate(f"planning_{chosen_day}-{dates.split('/')[1]}-{dates.split('/')[-1]}.pdf",
                                pagesize=A4)
    else:
        doc = SimpleDocTemplate(f"planning_{dates.split('/')[1]}-{dates.split('/')[-1]}.pdf", pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    rowHeights = 18
    units_list = ['INF - Infirmerie Jour', 'ARO - Unité de vie Aromates', 'LAV - Unité de vie Lavande',
                  'ORA - Unité de vie Orangeraie', 'ROS - Unité de vie Rose', 'TUL - Unité de vie Tulipe',
                  'VLT - Volants', 'NUI - Nuit', 'ASI - Hôtellerie']
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
