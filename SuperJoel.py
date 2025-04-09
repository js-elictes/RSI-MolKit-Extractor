#!/usr/bin/env python3
# SuperJoel is a program which extracts certain data from Gaussian .log files - Roithova Group
import os
import logging
import re
from openpyxl import Workbook
import docx
from docx.shared import Cm, Pt

# Constants and logging
__version__ = "1.6 -- 09.04.2025"
Hartree_to_kJ = 2625.4996394799
error_rate = 0
logging.basicConfig(level=logging.INFO)  # Configure Terminal Info

# Basic periodic table for converting atomic numbers to symbols (extend as needed)
atomic_symbols = {
    1: 'H',    2: 'He',   3: 'Li',   4: 'Be',   5: 'B',    6: 'C',    7: 'N',    8: 'O',    9: 'F',   10: 'Ne',
   11: 'Na',  12: 'Mg',  13: 'Al',  14: 'Si',  15: 'P',   16: 'S',   17: 'Cl',  18: 'Ar',  19: 'K',   20: 'Ca',
   21: 'Sc',  22: 'Ti',  23: 'V',   24: 'Cr',  25: 'Mn',  26: 'Fe',  27: 'Co',  28: 'Ni',  29: 'Cu',  30: 'Zn',
   31: 'Ga',  32: 'Ge',  33: 'As',  34: 'Se',  35: 'Br',  36: 'Kr',  37: 'Rb',  38: 'Sr',  39: 'Y',   40: 'Zr',
   41: 'Nb',  42: 'Mo',  43: 'Tc',  44: 'Ru',  45: 'Rh',  46: 'Pd',  47: 'Ag',  48: 'Cd',  49: 'In',  50: 'Sn',
   51: 'Sb',  52: 'Te',  53: 'I',   54: 'Xe',  55: 'Cs',  56: 'Ba',  57: 'La',  58: 'Ce',  59: 'Pr',  60: 'Nd',
   61: 'Pm',  62: 'Sm',  63: 'Eu',  64: 'Gd',  65: 'Tb',  66: 'Dy',  67: 'Ho',  68: 'Er',  69: 'Tm',  70: 'Yb',
   71: 'Lu',  72: 'Hf',  73: 'Ta',  74: 'W',   75: 'Re',  76: 'Os',  77: 'Ir',  78: 'Pt',  79: 'Au',  80: 'Hg',
   81: 'Tl',  82: 'Pb',  83: 'Bi',  84: 'Po',  85: 'At',  86: 'Rn',  87: 'Fr',  88: 'Ra',  89: 'Ac',  90: 'Th',
   91: 'Pa',  92: 'U',   93: 'Np',  94: 'Pu',  95: 'Am',  96: 'Cm',  97: 'Bk',  98: 'Cf',  99: 'Es', 100: 'Fm',
  101: 'Md', 102: 'No', 103: 'Lr', 104: 'Rf', 105: 'Db', 106: 'Sg', 107: 'Bh', 108: 'Hs', 109: 'Mt', 110: 'Ds',
  111: 'Rg', 112: 'Cn', 113: 'Nh', 114: 'Fl', 115: 'Mc', 116: 'Lv', 117: 'Ts', 118: 'Og'
}


def do_not_overwrite(path):
    n, e = os.path.splitext(path)
    i = 1
    while os.path.exists(path):
        path = f"{n} ({i}){e}"
        i += 1
    return path


def get_input(prompt, options):
    while True:
        choice = input(prompt).lower()
        if choice in options:
            return options[choice]
        logging.error(" -Select a valid option-")


def input_prompt():
    # Prompts the user for output preferences (Excel, Docs, or XYZ).
    print(f"\n  -- SuperJoel {__version__} by Jonáš Schröder --\n")
    option = get_input("Output an Excel, Docs, or XYZ file? [Excel/Docs/XYZ] : ",
                       {"excel": "excel", "e": "excel",
                        "docs": "docs", "d": "docs",
                        "xyz": "xyz", "x": "xyz"})
    print("\n  -- processing -- \n")
    # table_output is True for Excel, False for Docs and XYZ.
    table_output = True if option == "excel" else False
    return option, table_output


def export_relevant(log_file, table_output):
    # Extracts relevant information from a Gaussian .log file.
    global error_rate
    try:
        frq_header, ngeom = None, ""
        with open(log_file, 'r') as imported_file:
            content = imported_file.read()
        for i in re.finditer(r'---*\n (#.*?)---*', content, re.DOTALL):
            if "freq" in "".join(j.strip() for j in i.group(0)).lower():
                frq_header = i.group(1).replace("\n ", "")
                frq_header_pos = i.span()
        end_frq_pos = (re.search(r'Normal termination', content[frq_header_pos[1]:]).span()[1] +
                       frq_header_pos[1])
        frq_calc = content[frq_header_pos[0]:end_frq_pos]
        thermochem = " " + re.search(r'(Zero-point correction= .*?\n) \n', frq_calc, re.DOTALL).group(1)
        geom = re.findall(r' *Standard orientation: *\n -*\n.*?-*\n -*\n(.*?) -{10}', frq_calc, re.DOTALL)[-1]
        for i in geom.splitlines():
            num, atom, atype, x, y, z = i.split()
            ngeom = ngeom + " ".join([atom, " ", x, " ", y, " ", z]) + "\n"
        logging.info(f"✅ Processing file  :  {log_file}")
        if table_output:
            charge = int(re.search(r'-?\d+', re.search(r'Charge = .*?(?= Multiplicity)', frq_calc).group(0)).group())
            mult = int(re.search(r'-?\d+', re.search(r'Multiplicity = .*?\n', frq_calc).group(0)).group())
            thermochem_vals = [float(x) for x in re.findall(r'-?\d*\.\d+|-?\d+', thermochem)]
            thermochem_vals = [val for i, val in enumerate(thermochem_vals) if i not in (1, 2, 3, 5)]
            imag_values = [float(x) for x in re.findall(r'-?\d*\.\d+|-?\d+', 
                              re.search(r'Low frequencies ---.*?\n', frq_calc).group(0))]
            imag = "OK" if all(abs(val) < 30 for val in imag_values) else imag_values[0]
            E_tot = thermochem_vals[1] - thermochem_vals[0]
            E_ok, H_298k, G_298k = thermochem_vals[1:]
            return frq_header, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, ngeom
        else:
            chrgandmult = re.search(r'Charge = .*? Multiplicity = .*?\n', frq_calc).group(0)
            lowfrqs = "".join(re.findall(r'Low frequencies ---.*?\n', frq_calc))
            geomheader = "Atomic  Coordinates (Angstroms)\nAtomic#  X            Y                Z"
            outstr = "\n\n".join([log_file, frq_header, thermochem, lowfrqs, chrgandmult, geomheader, ngeom])
            return outstr, ngeom
    except Exception as e:
        logging.error(f"⚠️ Critical failure :  {log_file}")
        error_rate += 1
        outstr = "\n\n".join([log_file, "This file encountered an Error", "", "", "", "", ""])
        if table_output:
            return None
        else:
            return outstr, None


def create_excel_output(log_files, table_output):
    # Creates an Excel spreadsheet with extracted data and visualizations.
    wb = Workbook()
    ws = wb.active
    header = [
        "File name", "Header", "Charge", "Multiplicity", "Imag",
        "E-tot (Hartree)", "E-tot / rel (kJ/mol)", "E-ok (Hartree)",
        "E-ok / rel (kJ/mol)", "H-298k (Hartree)", "H-298k / rel (kJ/mol)",
        "G-298k (Hartree)", "G-298k / rel (kJ/mol)"
    ]
    ws.append(header)
    dataset = [export_relevant(f, table_output) for f in log_files]
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
    return wb


def create_word_output(log_files, table_output):
    # Creates a Word document with extracted data and visualizations
    doc = docx.Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    style = doc.styles['Normal']
    style.paragraph_format.space_before = style.paragraph_format.space_after = Cm(0)
    style.font.name, style.font.size = 'Arial', Pt(11)
    for log_file in log_files:
        outstr, geom = export_relevant(log_file, table_output)
        lines = outstr.split("\n")
        para = doc.add_paragraph(lines[0])
        run = para.runs[0]
        run.bold, run.font.size = True, Pt(14)
        for line in lines[1:]:
            doc.add_paragraph(line)
        doc.add_page_break()
    return doc


def extract_second_last_xyz(content):
    # Extracts the second-to-last 'Standard orientation' block using the original regex.
    blocks = re.findall(r' *Standard orientation: *\n -*\n.*?-*\n -*\n(.*?) -{10}', content, re.DOTALL)
    if len(blocks) < 2:
        return None
    block = blocks[-2]
    coords = []
    for line in block.strip().splitlines():
        parts = line.split()
        if len(parts) < 6:
            continue
        try:
            atomic_number = int(parts[1])
        except:
            continue
        symbol = atomic_symbols.get(atomic_number, 'X')
        x, y, z = parts[3:6]
        coords.append((symbol, x, y, z))
    return coords if coords else None


def create_xyz_output(log_files):
    #Merges the second-to-last coordinate blocks from all log files into one XYZ file.
    merged_geometries = []
    for log_file in log_files:
        data = export_relevant(log_file, True)
        if not data:
            logging.error(f"⚠️ Skipping {log_file}: export_relevant did not extract data.")
            continue
        try:
            frq_header, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, ngeom = data
        except Exception as e:
            logging.error(f"⚠️ Skipping {log_file}: error unpacking extraction data.")
            continue
        coord_lines = [line for line in ngeom.splitlines() if line.strip() != ""]
        if not coord_lines:
            logging.error(f"⚠️ Skipping {log_file}: no coordinate lines found in export_relevant output.")
            continue
        try:
            comment = f"{log_file} | Ehf={E_ok:.6f} | E0k={E_tot:.6f} | Imag={imag}"
        except Exception as e:
            comment = f"{log_file} | Imag={imag}"
            logging.warning(f"⚠️ Energy values missing or invalid in {log_file}.")
        merged_geometries.append({'source': log_file, 'atoms': coord_lines, 'comment': comment})
    if not merged_geometries:
        return None

    # Build merged string in XYZ format.
    merged_string = ""
    for geom in merged_geometries:
        count = len(geom['atoms'])
        merged_string += f"{count}\n"
        merged_string += f"{geom['comment']}\n"
        for line in geom['atoms']:
            merged_string += f"{line}\n"
    return merged_string


if __name__ == "__main__":
    option, table_output = input_prompt()
    log_files = [f for f in os.listdir() if f.endswith('.log')]
    if option == "excel":
        export_file = do_not_overwrite("SuperJoel Excel Output.xlsx")
        create_excel_output(log_files, table_output).save(export_file)
        print(f"\n Excel file created: {os.path.abspath(export_file)}")
    elif option == "docs":
        export_file = do_not_overwrite("SuperJoel Word Output.docx")
        create_word_output(log_files, table_output).save(export_file)
        print(f"\n Word file created: {os.path.abspath(export_file)}")
    elif option == "xyz":
        export_file = do_not_overwrite("SuperJoel XYZ Output.xyz")
        merged_xyz = create_xyz_output(log_files)
        if merged_xyz is not None:
            with open(export_file, "w") as f:
                f.write(merged_xyz)
            print(f"\n XYZ file created: {os.path.abspath(export_file)}")
        else:
            print("⚠️ No geometries extracted!")
    print(f"\n  -- Finished -- {error_rate} out of {len(log_files)} files encountered an Error --\n")
    print(r"""             __..--''``---....___   _..._    __
   /// //_.-'    .-/";  `        ``<._  ``.''_ `. / // /
  ///_.-' _..--.'_    \                    `( ) ) // //
  / (_..-' // (< _     ;_..__               ; `' / ///
   / // // //  `-._,_)' // / ``--...____..-' /// / //\
""")
    quit()
