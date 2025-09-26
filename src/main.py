import camelot
import fitz
import cv2
import numpy as np
import locale
from datetime import datetime
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import shutil


def is_right_color(cell_image_path, reference_color):
    try:
        cell_image = cv2.imread(cell_image_path)
        distance = np.linalg.norm(cell_image - reference_color, axis=2)
        mask = np.uint8((distance < 50) * 255)
    except TypeError:
        breakpoint()

    return 25000 < mask.sum()


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
    ('BOX', (0,0), (-1,-1), 1, colors.black),

    # Vertical lines between columns
    ('LINEBEFORE', (1,0), (1,-1), 1, colors.black),
    ('LINEBEFORE', (2,0), (2,-1), 1, colors.black),

    # Horizontal lines only inside right part (time + name)
    ('LINEABOVE', (1,1), (-1,-1), 0.5, colors.black),

    # Merge unit cells vertically
    ('SPAN', (0,0), (0,len(table_data)-1)),

    # Center vertically & add padding
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(table)
    story.append(Spacer(1, 12))


def main(pdf_file):
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
                ('20:00-8:00', np.array([180, 120, 130]))]
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
        date = datetime(int(dates.split('/')[-1]), int(dates.split('/')[1]), processed_date_day)
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

    doc = SimpleDocTemplate(f"planning_{dates.split('/')[1]}-{dates.split('/')[-1]}.pdf", pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    rowHeights = 18
    units_list = ['INF - Infirmerie Jour', 'ARO - Unité de vie Aromates', 'LAV - Unité de vie Lavande',
                  'ORA - Unité de vie Orangeraie',
                  'ROS - Unité de vie Rose', 'TUL - Unité de vie Tulipe', 'VLT - Volants', 'NUI - Nuit']

    for idx, (day, units) in enumerate(planning.items()):
        story.append(Paragraph(f"<b>{day}</b>", styles["Title"]))
        story.append(Spacer(1, 12))

        table_data = []
        add_to_table(table_data, story, rowHeights)

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

        table_data = [["ASI - Hôtellerie", "", "", ""], ["", "", "", ""], ["", "", "", ""], ["", "", "", ""],
                      ["", "", "", ""]]
        add_to_table(table_data, story, rowHeights)

        table_data = [["LINGERES", "", "", ""], ["", "", "", ""]]
        add_to_table(table_data, story, rowHeights)

        if idx < len(planning) - 1:
            story.append(PageBreak())

    doc.build(story)


if __name__ == "__main__":
    locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
    for pdf_file in os.listdir('plannings_mensuels'):
        main(os.path.join('plannings_mensuels', pdf_file))
    shutil.rmtree("crop_cell")
    shutil.rmtree("extracted_images")
