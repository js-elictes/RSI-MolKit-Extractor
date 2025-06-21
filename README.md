# RSI MolKit Extractor

**RSI MolKit Extractor** is a simple tool to extract geometries and thermodynamic data from Gaussian `.log` files, designed to quickly produce clean outputs for publications or analysis, including Excel summaries, Word documents, or `.xyz` structure files.

🔗 For instant web visualization of XYZ files, use my [MíšaXYZ Viewer](https://js-elictes.github.io/MisaXYZ/)

## 📦 Installation

Make sure you have Python 3.6 or later installed, then install the required packages:

```bash
pip install openpyxl python-docx
```

No other setup needed, just clone or download the repo.

## 🧪 How It Works

1. Put all your Gaussian `.log` files into the same folder as the script  
2. Run the script:

   ```bash
   python3 SuperJoel.py
   ```

3. Choose your output type when prompted:

   - **Excel**: outputs a table with thermochemical data  
   - **Docs**: generates a Word document with all key values  
   - **XYZ**: builds a multi-structure `.xyz` file with energy and metadata  

**Example run**:

```text
-- RSI MolKit Extractor 1.8, 21.06.2025 by Jonáš Schröder --

Output an Excel, Docs, or XYZ file? [Excel/Docs/XYZ] : excel

-- processing --

Excel file created: /path/to/MolKit_Excel_Output.xlsx

-- Finished, 0 out of 3 files encountered an error --
```

## 📁 Output Options

### 🟢 Excel

Creates a spreadsheet containing:

| File | Header             | Charge | Mult | Imag | E_tot    | E_rel   | E_0K     | H_298K   | G_298K   |
|------|--------------------|--------|------|------|----------|---------|----------|----------|----------|
| Per file | Thermochemistry | Relative energies in kJ/mol |  |  |  |  |  |  |  |

All energies are converted from Hartree to kJ/mol and displayed both absolute and relative to the most stable structure.

### 🟣 Word

Creates a formatted `.docx` file that includes:

- File name  
- Frequency job header  
- Charge and multiplicity  
- Zero-point and thermal corrections  
- Low frequencies  

### 🔵 XYZ

Combines all final optimized geometries into one `.xyz` file with structure-by-structure comments:

```text
C  0.000000  0.000000  0.000000
H  0.000000  0.000000  1.089000
...
# E(HF)=−312.458232 | E(0K)=−312.392814 | Imag=0 | Charge=0 | Multiplicity=1
```

## 🧪 Example

There is a `testfile.log` in the repo. Run the script in its folder to try it out.

## ⚙️ Notes

- Extracts data from jobs that terminated normally and contain frequency calculations  
- Uses the last standard orientation for optimized coordinates  
- Treats frequencies below 30 cm⁻¹ as non-imaginary  
- Logs and skips files that fail to parse  
- Saves outputs without overwriting existing files unless they are renamed

## 🔧 Troubleshooting

If you see a `ModuleNotFoundError`, run:

```bash
pip install openpyxl python-docx
```

If a file fails, check that it is a complete Gaussian output with a finished frequency job.

## ☕ License

MIT License – free to use, modify, and share.
