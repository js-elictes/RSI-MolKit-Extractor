#!/usr/bin/env python3
""" RSI MolKit Extractor
A simple program that extracts geometries and thermodynamic data of optimized Gaussian jobs. 
Put all your results into one folder and generate an Excel table with all thermodynamics or an .xyz file for supplementary materials. 
Place in the directory of your files and run. """

import os
import logging
import re
import csv
from datetime import datetime

# Constants and logging
__version__ = "2.1"
Hartree_to_kJ = 2625.4996394799
error_rate = 0
logging.basicConfig(level=logging.INFO)


def do_not_overwrite(path):
    n, e = os.path.splitext(path)
    i = 1
    while os.path.exists(path):
        path = f"{n} ({i}){e}"
        i += 1
    return path


def input_prompt():
    print(f"\n\033[1m   RSI MolKit Extractor v{__version__} ·  Roithová Group 17.12.2025\033[0m")
    options = {"excel": "excel", "e": "excel",
               "docs": "docs", "d": "docs",
               "xyz": "xyz", "x": "xyz",
               "all": "all", "a": "all"}
    while True:
        choice = input("\nOutput -> [E]xcel (.csv) \n          [D]ocs (.rtf) \n          [X]YZ \n          [A]ll\n          ->  ").lower()
        if choice in options:
            print("")
            return options[choice]
        logging.error(" Select a valid option")


def export_relevant(log_file, option):
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
        logging.info(f" Processing  ->  {log_file}")
        if option == "variables":
            charge = int(re.search(r'-?\d+', re.search(r'Charge = .*?(?= Multiplicity)', frq_calc).group(0)).group())
            mult = int(re.search(r'-?\d+', re.search(r'Multiplicity = .*?\n', frq_calc).group(0)).group())
            thermochem_vals = [float(x) for x in re.findall(r'-?\d*\.\d+|-?\d+', thermochem)]
            thermochem_vals = [val for i, val in enumerate(thermochem_vals) if i not in (1, 2, 3, 5)]
            imag_values = [float(x) for x in re.findall(r'-?\d*\.\d+|-?\d+', 
                              re.search(r'Low frequencies ---.*?\n', frq_calc).group(0))]
            imag = "0" if all(abs(val) < 30 for val in imag_values) else imag_values[0]
            E_tot = thermochem_vals[1] - thermochem_vals[0]
            E_ok, H_298k, G_298k = thermochem_vals[1:]
            return frq_header, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, ngeom
        else:
            chrgandmult = re.search(r'Charge = .*? Multiplicity = .*?\n', frq_calc).group(0)
            lowfrqs = "".join(re.findall(r'Low frequencies ---.*?\n', frq_calc))
            outstr = "\n\n".join([log_file, frq_header, chrgandmult, thermochem, lowfrqs])
            return outstr
    except Exception as e:
        logging.error(f"\033[1m Critical failure :  {log_file}\033[0m")
        error_rate += 1
        outstr = "\n\n".join([log_file, "This file encountered an Error\n"])
        if option == "variables":
            return None
        else:
            return outstr


def _escape_rtf(text: str) -> str:
    return text.replace("\\", r"\\").replace("{", r"\{").replace("}", r"\}")


def create_excel_output(log_files: list[str], out_path: str) -> None:
    header = ["File name", "Header", "Charge", "Multiplicity", "Imag",
              "E-tot (Hartree)", "E-tot / rel (kJ/mol)", "E-0K (Hartree)",
              "E-0K / rel (kJ/mol)", "H-298K (Hartree)", "H-298K / rel (kJ/mol)",
              "G-298K (Hartree)", "G-298K / rel (kJ/mol)"]
    dataset = [export_relevant(f, "variables") for f in log_files]
    ref_idx  = min(range(len(dataset)),
                   key=lambda i: dataset[i][4] if dataset[i] else float("inf"))
    rows = []
    for idx, data in enumerate(dataset):
        name = log_files[idx].removesuffix(".log")
        if not data:
            rows.append([name, "⚠️ This file encountered an Error"])
            continue
        frq, chg, mul, imag, Et, E0, H, G, _ = data
        rel = [round(abs(dataset[ref_idx][i] - data[i]) * Hartree_to_kJ, 1)
               for i in range(4, 8)]
        rows.append([name, frq, chg, mul, imag,
                     Et, rel[0], E0, rel[1], H, rel[2], G, rel[3]])
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


def create_word_output(log_files: list[str], out_path: str) -> None:
    rtf = [r"{\rtf1\ansi", r"{\fonttbl\f0 Arial;}",
        r"{\colortbl;\red0\green128\blue129;}", r"\fs22"]
    for f in log_files:
        raw = export_relevant(f, "string")
        if not raw:
            rtf.append(_escape_rtf(f"{f} Error while parsing") + r"\par\par")
            continue
        blocks = [b.strip() for b in re.split(r"\n\s*\n", raw) if b.strip()]
        rtf.append(r"{\b\cf1\fs32 " + _escape_rtf(blocks[0]) + r"}\par")  # header 16 pt, bold, teal
        for line in "\n".join(blocks[1:]).splitlines():
            if line.strip():
                rtf.append(_escape_rtf(line.strip()) + r"\par")
        rtf.append(r"\par")
    rtf.append("}")
    with open(out_path, "w", encoding="utf-8") as file:
        file.write("\n".join(rtf))


def create_xyz_output(log_files):
    merged_geometries = []
    for log_file in log_files:
        data = export_relevant(log_file, "variables")
        if not data:
            #logging.error(f"\033[F\033[1m Critical failure :  {log_file} -> missing data\033[0m")
            logging.error(f"\033[A\033[K\033[1m {log_file} is missing data\033[0m")
            continue
        frq_header, charge, mult, imag, E_tot, E_ok, H_298k, G_298k, ngeom = data
        coord_lines = [line for line in ngeom.splitlines() if line.strip()]
        comment = f"{log_file} | E(HF)={E_tot:.6f} | E(0K)={E_ok:.6f} | Imag={imag} | Charge={charge} | Multiplicity={mult}"
        merged_geometries.append({'atoms': coord_lines, 'comment': comment})
    if not merged_geometries:
        return None
    merged_string = "# You can open this file using our opensource XYZ Viewer: https://js-elictes.github.io/RSI-MolKit-Viewer/\n\n"
    for i, geom in enumerate(merged_geometries):
        count_line = f"{len(geom['atoms'])}\n"
        merged_string += count_line
        merged_string += f"{geom['comment']}\n"
        merged_string += "\n".join(geom['atoms']) + "\n"
    return merged_string


if __name__ == "__main__":
    option = input_prompt()
    log_files = [f for f in os.listdir() if f.endswith('.log')]
    timestamp = datetime.now().strftime("%d%m%Y")
    actions = {
        "excel": ("MolKit_Excel", lambda p: create_excel_output(log_files, p)),
        "docs":  ("MolKit_Word", lambda p: create_word_output(log_files, p)),
        "xyz":   ("MolKit_XYZ",   lambda p: open(p, "w").write(create_xyz_output(log_files)))
    }
    selected = actions if option == "all" else {option: actions[option]}
    for name, (prefix, func) in selected.items():
        ext = {"excel": "csv", "docs": "rtf", "xyz": "xyz"}[name]
        filename = do_not_overwrite(f"{prefix}_{timestamp}.{ext}")
        func(filename)
        print(f"\033[1m\n{ name.capitalize() } file created:\033[0m {os.path.abspath(filename)}")
    if option == "all":
        error_rate = max(1, error_rate // 3)
    print(f"\n  \033[1m   Finished · {error_rate}/{len(log_files)} files unsucessful -> wrong format\033[0m\n")
    print(r"""          __..--''``---....___   _..._    __
     /// //_.-'    .-/";  `        ``<._  ``.''_ `. / // /
     ///_.-' _..--.'_    \                    `( ) ) // //
     / (_..-' // (< _     ;_..__               ; `' / ///
     / // // //  `-._,_)' // / ``--...____..-' /// / //\
        """)
    quit()
