# TC

Software to analyze TC reports.

## 1. Preparation

Execute the main script `run.ps1`. This script installs the dependencies, creates the analysis script cards.py

```powershell
.\run.ps1
```

## 2. Analysis Execution and Data Visualization

1. Run the script in the terminal:

```
python app.py
```
2. Load the pdf files inside PDFS folder.

## 3. Results

After the Analysis, the `Resultados/` folder will also contain the analysis results in Excel files, organized in the subfolders `MC_Resultados/` and `Visa_Resultados/`. The resulting structure will be similar to the following:     

```
TC/
├── PDFS/
│   ├── MC/
│   │   ├── 24.MCGAenero2025.pdf
│   │   ├── 25.MCGAfebrero2025.pdf
│   │   ├── 26.MCGAmarzo2025.pdf
│   │   └── 27.MCGAabril2025.pdf
│   ├── Visa/
│   │   ├── 34.VisaGAenero2025.pdf
│   │   ├── 35.VisaGAfebrero2025.pdf
│   │   ├── 36.VisaGAmarzo2025.pdf
│   │   └── 37.VisaGAabril2025.pdf
│   ├── categorias.xlsx
│   ├── cedulas.xlsx
│   └── TRM.xlsx
├── Resulatdos/
│   ├── MC_Resultados/
│   │   └── MC_20250730_0813.xlsx
│   └── Visa_Resultados/
│       └── Visa_20250730_0813.xlsx
├── cards.py
├── .gitignore
├── README.md
└── run.ps1