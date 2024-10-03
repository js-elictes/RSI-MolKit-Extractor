#!/usr/bin/env python3
# SuperJoel is a program which extracts certain data from Gaussian .log files - Roithova Group
import os
import logging
import re
import numpy as np
from scipy.spatial.distance import cdist
import plotly.graph_objects as go
import plotly.io as pio
import io
import openpyxl
from openpyxl import Workbook
import docx
from docx.shared import Cm, Pt, Mm


# Constants and logging
__version__ = "1.5 -- 03.10.2024"
Hartree_to_kJ = 2625.4996394799
error_rate = 0
logging.basicConfig(level=logging.INFO) # Configure Terminal Info


def do_not_overwrite(path):
    """Appends a counter to a filename to avoid overwriting existing files."""
    n, e = os.path.splitext(path); i = 1
    while os.path.exists(path): path = f"{n} ({i}){e}"; i += 1
    return path


def get_input(prompt, options):
    """Takes user input and matches it against a dictionary of options."""
    while True:
        choice = input(prompt).lower()
        if choice in options:
            return options[choice]
        logging.error(" -Select a valid option-")


def input_prompt():
    """Prompts the user for output preferences and image incorporation."""
    print(f"\n  -- SuperJoel {__version__} by Jonáš Schröder --\n")
    word_or_excel = get_input("Output an Excel or a Docs file? [Excel/Docs] : ",
                              {"excel": True, "e": True, "docs": False, "d": False})
    image = get_input("Incorporate images within the file? [Yes/No] : ",
                      {"yes": True, "y": True, "no": False, "n": False})
    print("\n  -- processing --")
    return word_or_excel, image


def visualisation(geom, filename):
    """Generates a 3D visualization of atomic coordinates."""
    data = np.loadtxt(io.StringIO(geom))
    atom_type, coordinates = data[:, 0].astype(int), data[:, 1:]
    n, dists = len(coordinates), cdist(coordinates, coordinates)
    np.fill_diagonal(dists, np.inf)
    i, j = np.triu_indices(n, 1)
    dist_upper, atom_i = dists[i, j], atom_type[i]
    mask = (dist_upper < 1.95) & ~((atom_i == 1) & (dist_upper > 0.5)) & ~((atom_i >= 18) & (dist_upper > 1.7))
    close_pairs = [(coordinates[a], coordinates[b]) for a, b in zip(i[mask], j[mask])]
    colors = {6: 'darkgray', 1: 'lightgray', 8: 'red', 7: 'green', 17: 'blue', 9: 'blue', 14: 'yellow',
              11: 'pink', 19: 'pink', 3: 'pink', 55: 'pink', 21: 'gold', 72: 'gold', 35: 'blue', 53: 'blue'}
    sizes = {k: 7 if k > 2 else 5 for k in range(1, 37)}
    fig = go.Figure(data=[go.Scatter3d(x=[p[0][0], p[1][0]], y=[p[0][1], p[1][1]], z=[p[0][2], p[1][2]],
                                       mode='lines', line=dict(color='black', width=2)) for p in close_pairs])
    fig.add_trace(go.Scatter3d(x=coordinates[:, 0], y=coordinates[:, 1], z=coordinates[:, 2], mode='markers',
                               marker=dict(size=[sizes.get(a, 9) for a in atom_type],
                                           color=[colors.get(a, 'black') for a in atom_type],
                                           line=dict(color='black', width=1))))
    noax = dict(visible=False, showgrid=False, backgroundcolor='white')
    fig.update_layout(scene=dict(xaxis=noax, yaxis=noax, zaxis=noax), showlegend=False)
    fig.add_annotation(x=0.5, y=0.9, text=filename, showarrow=False, font=dict(family='Arial', size=30, color='black'))
    img_data = pio.to_image(fig, format='png', width=1000, height=1000)
    logging.info(f'  Processing image :  {filename}')
    return img_data


def export_relevant(log_file):
    """Extracts relevant information from a Gaussian .log file."""
    global error_rate
    try:
        frq_header, ngeom = None, ""
        with open(log_file, 'r') as imported_file:
            content = imported_file.read()
        for i in re.finditer(r'---*\n (#.*?)---*', content, re.DOTALL):
            if "freq" in "".join(j.strip() for j in i.group(0)).lower():
                frq_header = i.group(1).replace("\n ", "")
                frq_header_pos = i.span()
        end_frq_pos = (re.search(r'Normal termination', content[frq_header_pos[1]:]).span()[1] + frq_header_pos[1])
        frq_calc = content[frq_header_pos[0]:end_frq_pos]
        thermochem = " " + re.search(r'(Zero-point correction= .*?\n) \n', frq_calc, re.DOTALL).group(1)
        geom = re.findall(r' *Standard orientation: *\n -*\n.*?-*\n -*\n(.*?) -{10}', frq_calc, re.DOTALL)[-1]
        for i in geom.splitlines():
            num, atom, atype, x, y, z = i.split()
            ngeom = ngeom + " ".join([atom," ", x," ", y," ", z]) + "\n"
        logging.info(f"  Processing file  :  {log_file}")
        if word_or_excel:
            charge, mult = int(
                re.search(r'-?\d+', re.search(r'Charge = .*?(?= Multiplicity)', frq_calc).group(0)).group()), \
                int(re.search(r'-?\d+', re.search(r'Multiplicity = .*?\n', frq_calc).group(0)).group())
            thermochem = [val for i, val in enumerate([float(x) for x in re.findall(r'-?\d*\.\d+|-?\d+', thermochem)])
                        if i not in (1, 2, 3, 5)]
            imag = [float(x) for x in
                    re.findall(r'-?\d*\.\d+|-?\d+', re.search(r'Low frequencies ---.*?\n', frq_calc).group(0))]
            imag = "OK" if all(abs(val) < 30 for val in imag) else imag[0]
            E_tot = thermochem[1] - thermochem[0]
            E_ok, H_298k, G_298k = thermochem[1:]
            return frq_header, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, ngeom
        else:
            chrgandmult = re.search(r'Charge = .*? Multiplicity = .*?\n', frq_calc).group(0)
            lowfrqs = "".join(re.findall(r'Low frequencies ---.*?\n', frq_calc))
            geomheader = "Atomic  Coordinates (Angstroms)\nAtomic#  X            Y                Z"

            outstr = "\n\n".join([log_file, frq_header, thermochem, lowfrqs, chrgandmult, geomheader, ngeom])
            return outstr, ngeom
    except:
        logging.error(f" Critical failure :  {log_file}")
        outstr = "\n\n".join([log_file, "This file encountered an Error", "", "", "", "", ""])
        error_rate = error_rate + 1
        if word_or_excel:
            return None
        else:
            return outstr, None


def create_excel_output(log_files, images):
    """Creates an Excel spreadsheet with extracted data and visualizations."""
    wb = Workbook()
    ws = wb.active
    header = [
        "File name", "Header", "Charge", "Multiplicity", "Imag",
        "E-tot (Hartree)", "E-tot / rel (kJ/mol)", "E-ok (Hartree)",
        "E-ok / rel (kJ/mol)", "H-298k (Hartree)", "H-298k / rel (kJ/mol)",
        "G-298k (Hartree)", "G-298k / rel (kJ/mol)"
    ]
    ws.append(header)
    dataset = [export_relevant(f) for f in log_files]
    Energs = [data[4] if data else float('inf') for data in dataset]
    smallest = Energs.index(min(Energs))
    for idx, data in enumerate(dataset):
        log_file = log_files[idx].replace(".log", "")
        if not data:
            ws.append([log_file, 'This file encountered an Error'])
            continue
        frqheader, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, geom = data
        Etotrel, Eokrel, H298rel, G298rel = [
            round(abs(dataset[smallest][i] - data[i]) * Hartree_to_kJ, 1)
            for i in range(4, 8)
        ]
        row = [
            log_file, frqheader, charge, mult, imag, E_tot, Etotrel,
            E_ok, Eokrel, H_298k, H298rel, G_298k, G298rel
        ]
        ws.append(row)
        if images and geom:
            img_data = visualisation(geom, log_file)
            img = openpyxl.drawing.image.Image(io.BytesIO(img_data))
            img.width, img.height = 100, 100
            img.anchor = ws.cell(row=ws.max_row, column=len(header) + 1).coordinate
            ws.column_dimensions[ws.cell(row=ws.max_row, column=ws.max_column).column_letter].width = 10
            ws.row_dimensions[ws.max_row].height = 80
            ws.add_image(img)
    return wb


def create_word_output(log_files, images):
    """Creates a Word document with extracted data and visualizations."""
    doc = docx.Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    style = doc.styles['Normal']
    style.paragraph_format.space_before = style.paragraph_format.space_after = Cm(0)
    style.font.name, style.font.size = 'Arial', Pt(11)
    for log_file in log_files:
        outstr, geom = export_relevant(log_file)
        if geom and images:
            imgdata = visualisation(geom, log_file)
            pic = doc.add_picture(io.BytesIO(imgdata))
            pic.height = pic.width = docx.shared.Mm(140)
        lines = outstr.split("\n")
        para = doc.add_paragraph(lines[0])
        run = para.runs[0]
        run.bold, run.font.size = True, Pt(14)
        for line in lines[1:]:
            doc.add_paragraph(line)
        doc.add_page_break()
    return doc
    

if __name__ == "__main__":
    word_or_excel, images = input_prompt()
    export_file = do_not_overwrite("SuperJoel Excel Output.xlsx" if word_or_excel else "SuperJoel Word Output.docx")
    print(f"\n{os.path.abspath(export_file)}\n")
    log_files = [f for f in os.listdir() if f.endswith('.log')]
    (create_excel_output if word_or_excel else create_word_output)(log_files, images).save(export_file)
    print(f"\n  -- Finished -- {error_rate} out of {len(log_files)} encountered an Error --\n")
    print("""             __..--''``---....___   _..._    __
   /// //_.-'    .-/";  `        ``<._  ``.''_ `. / // /
  ///_.-' _..--.'_    \                    `( ) ) // //
  / (_..-' // (< _     ;_..__               ; `' / ///
   / // // //  `-._,_)' // / ``--...____..-' /// / //\n""")
    quit()
