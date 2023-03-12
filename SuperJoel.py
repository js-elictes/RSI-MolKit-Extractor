#!/usr/bin/env python3
# SuperJoel is a program which extracts certain data from .log files
import os
import logging
import re
import math
import plotly.graph_objects as go
import plotly.io as pio
import io
import openpyxl
from openpyxl import Workbook
import docx  # pip
from docx.shared import *

logging.basicConfig(level=logging.INFO)
__version__ = 0.9
Hartree_to_kJ = 2625.4996394799


def do_not_overwrite(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


def input_prompt():
    il, il2, il3, il4 = False, False, False, False
    verstr = "SuperJoel ver. {} by Jonáš Schröder".format(__version__)
    print(verstr + "\n" + "-" * len(verstr))

    while not il:
        tot_input = input(
            "Do you want an Excel or a Docs file? [Excel/Docs] : ").lower()
        if tot_input in ["excel", "e"]:
            tort, il = True, True
        elif tot_input in ["docs", "d"]:
            tort, il = False, True
        else:
            logging.error("Select a valid option !!!")

    while not il2:
        tot_input = input(
            "Do you want images in your file? [Yes/No] : ").lower()
        if tot_input in ["yes", "y"]:
            img, il2 = True, True
            while not il3:
                tot_input = input(
                    "\\Do you want to automatically or manually?\n"
                    "(Automatic results may be badly rotated.) "
                    "[Auto/Manual] : ").lower()
                if tot_input in ["auto", "a"]:
                    autoimg, il3 = True, True
                elif tot_input in ["manual", "m"]:
                    autoimg, il3 = False, True
                else:
                    logging.error("Select a valid option !!!")
        elif tot_input in ["no", "n"]:
            img, autoimg, il2 = False, False, True
        else:
            logging.error("Select a valid option !!!")

    while not il4:
        tot_input = input(
            "All or only selected .log files : [All/Selected] : ").lower()
        if tot_input in ["all", "a"]:
            all_or_selected, il4 = True, True
        elif tot_input in ["selected", "s"]:
            all_or_selected, il4 = False, True
        else:
            logging.error("Select a valid option !!!")
    return tort, img, autoimg, all_or_selected


def visualisation(geometry, filename, Manual=False):
    # function to calculate distance between two atoms
    def distance(x1, y1, z1, x2, y2, z2):
        return math.sqrt((x2 - x1)**2 + (y2 - y1)**2 + (z2 - z1)**2)

    with io.StringIO(geom) as f:
        data = [line.split() for line in f]
        coordinates = [list(map(float, i[1:])) for i in data]
        atom_type = [int(item[0]) for item in data]
    close_pairs = []

    for i in range(len(coordinates)):
        for j in range(len(coordinates)):
            if i == j:
                continue
            x1, y1, z1 = coordinates[i]
            x2, y2, z2 = coordinates[j]
            dist = distance(x1, y1, z1, x2, y2, z2)
            if atom_type[i] == 1 and dist > 0.5:
                continue
            if atom_type[i] >= 18 and dist > 1.7:
                continue
            if dist < 1.95:
                close_pairs.append([[x1, y1, z1], [x2, y2, z2]])

    fig = go.Figure(
        data=[
            go.Scatter3d(
                x=[
                    p[0][0], p[1][0]], y=[
                    p[0][1], p[1][1]], z=[
                        p[0][2], p[1][2]], mode='lines', line=dict(
                            color='black', width=2)) for p in close_pairs])

    colors = {
        6: 'darkgray',
        1: 'lightgray',
        8: 'red',
        7: 'green',
        17: 'blue',
        9: 'blue',
        14: 'yellow',
        11: "pink",
        19: "pink",
        11: "pink",
        3: "pink",
        55: "pink",
        21: "gold",
        72: "gold",
        35: "blue",
        53: "blue"}

    sizes = {
        1: 5,
        2: 5,
        3: 7,
        4: 7,
        5: 7,
        6: 7,
        8: 7,
        9: 7,
        10: 7,
        11: 7,
        12: 7,
        13: 7,
        14: 7,
        15: 7,
        16: 7,
        17: 7,
        18: 7,
        19: 7,
        20: 7,
        21: 7,
        22: 7,
        23: 7,
        24: 7,
        25: 7,
        26: 7,
        27: 7,
        28: 7,
        29: 7,
        30: 7,
        31: 7,
        32: 7,
        33: 7,
        34: 7,
        35: 7,
        36: 7}
    fig.add_trace(
        go.Scatter3d(
            x=[
                coordinates[i][0] for i in range(
                    len(coordinates))], y=[
                coordinates[i][1] for i in range(
                    len(coordinates))], z=[
                coordinates[i][2] for i in range(
                    len(coordinates))], mode='markers', marker=dict(
                size=[
                    sizes.get(
                        atom_type[i], 9) for i in range(
                        len(coordinates))], color=[
                    colors.get(
                        atom_type[i], 'black') for i in range(
                        len(coordinates))], line=dict(
                    color='black', width=1))))

    noax = dict(visible=False, showgrid=False, backgroundcolor="white")
    fig.update_layout(
        scene=dict(
            xaxis=noax,
            yaxis=noax,
            zaxis=noax),
        showlegend=False)

    if not Manual:
        img_data = pio.to_image(fig, format='png', width=1000, height=1000)
        logging.info(f"The image of {filename} Exported")
        return (img_data)
    else:
        fig.show()
        logging.info(f"The image of {filename} was shown")
        input("Press Enter to continue: ")


def export_relevant(log_file):
    frqheader = None
    ngeom = ""
    logging.info("Processing file {}".format(log_file))
    with open(log_file, 'r') as imported_file:
        input = imported_file.read()

    # Search for the last frequency header
    for i in re.finditer(r'---*\n (#.*?)---*', input, re.DOTALL):
        if "freq" in "".join(j.strip() for j in i.group(0)).lower():
            frqheader = i.group(1).replace("\n ", "")
            frqheaderpos = i.span()

    if not frqheader:
        logging.error("File {} does not contain frequencies. Skipping ..."
                      .format(log_file))
        if text_or_table:
            return None
        else:
            outstr = ("File \"{}\" skipped. \n"
                      "No frequency calculation Found!!!"
                      .format(log_file))
            return outstr, None
    else:
        endfrqpos = (re.search(r'Normal termination', input[frqheaderpos[1]:])
                     .span()[1] + frqheaderpos[1])
        frqcalc = input[frqheaderpos[0]:endfrqpos]
        thermochem = " " + re.search(r'(Zero-point correction= .*?\n) \n',
                                     frqcalc, re.DOTALL).group(1)
        geom = re.findall(
            r' *Standard orientation: *\n -*\n.*?-*\n -*\n(.*?) -{10}',
            frqcalc, re.DOTALL)[-1]
        for i in geom.splitlines():
            num, atom, atype, x, y, z = i.split()
            ngeom = ngeom + " ".join([atom, x, y, z]) + "\n"

        if text_or_table:
            charge = int(re.search(r'-?\d+', re.search(
                r'Charge = .*?(?= Multiplicity)', frqcalc).group(0)).group())
            mult = int(re.search(r'-?\d+', re.search(
                r'Multiplicity = .*?\n', frqcalc).group(0)).group())
            thermochem = [val for i, val in enumerate(
                [float(x) for x in re.findall(r'-?\d*\.\d+|-?\d+',
                                              thermochem)])
                if i not in (1, 2, 3, 5)]
            imag = [float(x) for x in re.findall(
                r'-?\d*\.\d+|-?\d+', re.search(r'Low frequencies ---.*?\n',
                                               frqcalc).group(0))]
            imag = "OK" if all(abs(val) < 30 for val in imag) else imag[0]
            E_tot = thermochem[1] - thermochem[0]
            E_ok, H_298k, G_298k = thermochem[1:]
            return frqheader, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, ngeom
        else:
            chrgandmult = re.search(
                r'Charge = .*? Multiplicity = .*?\n', frqcalc).group(0)
            lowfrqs = "".join(re.findall(r'Low frequencies ---.*?\n', frqcalc))
            geomheader = ("Atomic  Coordinates (Angstroms)"
                          "\nAtomic#  X      Y         Z")

            outstr = "\n\n".join([log_file, frqheader, thermochem, lowfrqs,
                                  chrgandmult, geomheader, ngeom])
            return outstr, ngeom


if __name__ == "__main__":
    text_or_table, images, autoimg, all_or_selected = input_prompt()
    if text_or_table:
        export_file = "SuperJoel Excel Output.xlsx"
    else:
        export_file = "SuperJoel Word Output.docx"
    export_file = do_not_overwrite(export_file)
    pre_log_files = [f for f in os.listdir() if f.endswith('.log')]
    log_files = []
    if not all_or_selected:
        print(f"These are all the .log files: {log_files}")
        for file in pre_log_files:
            tot_input = input(
                f"Do you want to process the file: {file}? "
                "[Yes/No] : ").lower()
            if tot_input in ["yes", "y"]:
                log_files.append(file)
            elif tot_input in ["no", "n"]:
                pass
            else:
                log_files.append(file)
                logging.error(
                    "This was not a valid choice!!! "
                    "The file will be processed.")
    else:
        log_files = pre_log_files
    print("")
    logging.info(f"This program will export data to {export_file}\n"
                 f"{os.getcwd()}")

    datarows = []
    image_data = []
    if not log_files:
        logging.error("{} does not contain any .log files, "
                      "nothing to do, Quitting now.".format(os.getcwd()))
        quit()
    logging.info("Exporting data to ./{}".format(export_file))
    if text_or_table:
        wb = Workbook()
        ws = wb.active
        header = ["File name", "Header", "Charge", "Multiplicity",
                  "Imag", "E-tot (Hartree)", "E-tot / rel (kJ/mol)",
                  "E-ok (Hartree)", "E-ok / rel (kJ/mol)",
                  "H-298k (Hartree)", "H-298k / rel (kJ/mol)",
                  "G-298k (Hartree)", "G-298k / rel (kJ/mol)"]
        ws.append(header)

        dataset = [export_relevant(i) for i in log_files]
        Energs = [0 if not i else i[4] for i in dataset]
        smallest = Energs.index(min(Energs))

        for i, data in enumerate(dataset):
            log_file = log_files[i].replace(".log", "")
            if not data:
                ws.append([log_file, 'The Reptiles have infected this'
                           'file... It contains no frequencies!!!'])
                continue
            frqheader, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, geom = data
            if geom:
                if images:
                    if not autoimg:
                        visualisation(geom, log_file, True)
                    else:
                        image_data.append(visualisation(geom, log_file))

            Etotrel, Eokrel, H298rel, G298rel = [round(
                abs(dataset[smallest][i] - data[i]) * Hartree_to_kJ, 1)
                for i in range(4, 8)]
            datarows.append([log_file, frqheader, charge, mult, imag, E_tot,
                            Etotrel, E_ok, Eokrel, H_298k, H298rel, G_298k,
                            G298rel])

            datax = sorted(datarows, key=lambda x: float(x[5]))

        for i in range(len(datax)):
            ws.append(datax[i])
            if autoimg:
                img = openpyxl.drawing.image.Image(io.BytesIO(image_data[i]))
                img.width = 100
                img.height = 100
                img.anchor = ws.cell(
                    row=ws.max_row, column=len(header) + 1).coordinate
                ws.column_dimensions[ws.cell(
                    row=ws.max_row,
                    column=ws.max_column).column_letter].width = 10
                ws.row_dimensions[ws.max_row].height = 80
                ws.add_image(img)

        wb.save(export_file)
        logging.info("Export finished. Normal termination")
        quit()

    else:
        doc = docx.Document()
        with open(export_file, 'w') as f:
            section = doc.sections[0]
            section.page_width = Cm(21)
            section.page_height = Cm(29.7)
            style = doc.styles['Normal']
            style.paragraph_format.space_before = Cm(0)
            style.paragraph_format.space_after = Cm(0)
            font = style.font
            font.name = 'Arial'
            font.size = Pt(11)
            for log_file in log_files:
                outstr, geom = export_relevant(log_file)
                if geom:
                    if images:
                        if not autoimg:
                            visualisation(geom, log_file, True)
                        else:
                            imgdata = visualisation(geom, log_file)
                            picture = doc.add_picture(io.BytesIO(imgdata))
                            picture.height = docx.shared.Mm(140)
                            picture.width = docx.shared.Mm(140)
                lines = outstr.split("\n")
                first_line = lines[0]
                paragraph = doc.add_paragraph(first_line)
                run = paragraph.runs[0]
                run.bold = True
                run.font.size = Pt(14)
                for line in lines[1:]:
                    doc.add_paragraph(line)
                doc.add_page_break()
            doc.save(export_file)
            logging.info("Export finished. Normal termination")
            quit()