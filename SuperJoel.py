#!/usr/bin/env python3
# SuperJoel is a program which extracts certain data from Gaussian .log files
import os
import logging
import re
import math
import plotly.graph_objects as go
import plotly.io as pio
import io
import openpyxl
from openpyxl import Workbook
import docx
from docx.shared import *

#  Configure logging
logging.basicConfig(level=logging.INFO)


# Constants
__version__ = "1.4 / 25.08.2023"
Hartree_to_kJ = 2625.4996394799
error_rate = 0


# Appends a counter to a filename to avoid overwriting existing files.
def do_not_overwrite(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1
    return path


# Takes user input and matches it against a dictionary of options.
def get_input(prompt, options):
    while True:
        choice = input(prompt).lower()
        for opt, aliases in options.items():
            if choice in aliases:
                return opt
        logging.error(" -Select a valid option-")


# Prompts the user for output preferences (Excel or Docs) and image incorporation.
def input_prompt():
    print(f"\n  -- SuperJoel {__version__} by Jonáš Schröder --\n")

    word_or_excel = get_input("Output an Excel or a Docs file? [Excel/Docs] : ",
                     {"True": ["excel", "e"], "False": ["docs", "d"]}) == "True"
    image = get_input("Incorporate images within the file? [Yes/No] : ",
                    {"True": ["yes", "y"], "False": ["no", "n"]}) == "True"
    print("\n  -- processing --")
    return word_or_excel, image


# Calculates the distance between two 3D coordinates.
def atom_distance(x1, y1, z1, x2, y2, z2):
    return math.sqrt((x2 - x1)**2 + (y2 - y1)**2 + (z2 - z1)**2)


# Generates a 3D visualization of atomic coordinates.
def visualisation(geom, filename):
    # function to calculate distance between two atoms
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
            dist = atom_distance(x1, y1, z1, x2, y2, z2)
            if atom_type[i] == 1 and dist > 0.5:
                continue
            if atom_type[i] >= 18 and dist > 1.7:
                continue
            if dist < 1.95:
                close_pairs.append([[x1, y1, z1], [x2, y2, z2]])

    fig = go.Figure(data=[go.Scatter3d(x=[p[0][0], p[1][0]], y=[p[0][1], p[1][1]], z=[p[0][2], p[1][2]],
                                       mode='lines', line=dict(color='black', width=2)) for p in close_pairs])

    colors = {6: 'darkgray', 1: 'lightgray', 8: 'red', 7: 'green', 17: 'blue', 9: 'blue', 14: 'yellow',
              11: "pink", 19: "pink", 3: "pink", 55: "pink", 21: "gold", 72: "gold", 35: "blue", 53: "blue"}
    sizes = {k: 7 if k > 2 else 5 for k in range(1, 37)}

    fig.add_trace(go.Scatter3d(x=[coord[0] for coord in coordinates], y=[coord[1] for coord in coordinates],
                               z=[coord[2] for coord in coordinates], mode='markers',
                               marker=dict(size=[sizes.get(atom, 9) for atom in atom_type],
                                           color=[colors.get(atom, 'black') for atom in atom_type],
                                           line=dict(color='black', width=1))))

    noax = dict(visible=False, showgrid=False, backgroundcolor="white")
    fig.update_layout(scene=dict(xaxis=noax, yaxis=noax, zaxis=noax), showlegend=False)
    fig.add_annotation(x=0.5, y=0.9, text=filename, showarrow=False,
                       font=dict(family="Arial", size=30, color="black"))

    img_data = pio.to_image(fig, format='png', width=1000, height=1000)
    logging.info(f"  Processing image :  {filename}")
    return (img_data)


# Extracts relevant information from a Gaussian .log file.
def export_relevant(log_file):
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


# Creates an Excel spreadsheet with extracted data and visualizations.
def create_excel_output(log_files, images):
    datarows = []
    image_data = []
    wb = Workbook()
    ws = wb.active
    header = ["File name", "Header", "Charge", "Multiplicity",
                "Imag", "E-tot (Hartree)", "E-tot / rel (kJ/mol)",
                "E-ok (Hartree)", "E-ok / rel (kJ/mol)",
                "H-298k (Hartree)", "H-298k / rel (kJ/mol)",
                "G-298k (Hartree)", "G-298k / rel (kJ/mol)"]
    ws.append(header)
    
    try:
        dataset = [export_relevant(i) for i in log_files]
        Energs = [0 if not i else i[4] for i in dataset]
        smallest = Energs.index(min(Energs))

        for i, data in enumerate(dataset):
            log_file = log_files[i].replace(".log", "")
            if not data:
                ws.append([log_file, 'This file encountered an Error'])
                continue
            frqheader, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, geom = data
            if geom:
                if images:
                    image_data.append(visualisation(geom, log_file))

            Etotrel, Eokrel, H298rel, G298rel = [round(
                abs(dataset[smallest][i] - data[i]) * Hartree_to_kJ, 1)
                for i in range(4, 8)]
            datarows.append([log_file, frqheader, charge, mult, imag, E_tot,
                            Etotrel, E_ok, Eokrel, H_298k, H298rel, G_298k,
                            G298rel])
            
        for i in range(len(datarows)):
            ws.append(datarows[i])
            if images:
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
    except:
        pass
    return wb


# Creates a Word document with extracted data and visualizations.
def create_word_output(log_files, images):
    doc = docx.Document()
    for log_file in log_files:
        outstr, geom = export_relevant(log_file)
        section = doc.sections[0]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        style = doc.styles['Normal']
        style.paragraph_format.space_before = Cm(0)
        style.paragraph_format.space_after = Cm(0)
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        if geom:
            if images:
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
    return doc
    

# The main execution block that gets user input, processes files, generates output, and saves results.   
if __name__ == "__main__":
    word_or_excel, images = input_prompt()
    export_file = "SuperJoel Excel Output.xlsx" if word_or_excel else "SuperJoel Word Output.docx"
    export_file = do_not_overwrite(export_file)
    print(f"\n{os.getcwd()}/{export_file}\n")

    log_files = [f for f in os.listdir() if f.endswith('.log')]
    log_file_number = sum(1 for file in os.listdir() if file.endswith('.log'))

    if word_or_excel:
        wb = create_excel_output(log_files, images)
        wb.save(export_file)
    else:
        doc = create_word_output(log_files, images)
        doc.save(export_file)
    
    print(f"\n  -- Finished -- {error_rate} out of {log_file_number} encountered an Error --\n")
    print("""             __..--''``---....___   _..._    __
   /// //_.-'    .-/";  `        ``<._  ``.''_ `. / // /
  ///_.-' _..--.'_    \                    `( ) ) // //
  / (_..-' // (< _     ;_..__               ; `' / ///
   / // // //  `-._,_)' // / ``--...____..-' /// / //\n""")
    quit()
