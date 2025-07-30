function TC {
    param (
        [string]$ExcelFilePath = $null
    )

    $YELLOW = [ConsoleColor]::Yellow

    Write-Host "🚀 Creating TC framework" -ForegroundColor $YELLOW

    # Create Python virtual environment
    python -m venv .venv
    .\.venv\scripts\activate

    # Install required Python packages
    python.exe -m pip install --upgrade pip
    python -m pip install PyMuPDF pandas pdfplumber openpyxl

    # Create templates directory structure
    $directories = @(
        "PDFS",
        "PDFS/MC",
        "PDFS/Visa",
        "Resultados",
        "Resultados/MC_Resultados",
        "Resultados/Visa_Resultados"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

# Create models.py with cedula as primary key
Set-Content -Path "cards.py" -Value @" 
import os
import re
import fitz  
import pdfplumber  
import pandas as pd
from datetime import datetime

# --- Configuration ---
input_base_folder = "PDFS"
output_base_folder = "Resultados" # Base folder for all results
trm_file = os.path.join(input_base_folder, "TRM.xlsx")
trm_sheet = "Datos"
categorias_file = os.path.join(input_base_folder, "categorias.xlsx")
cedulas_file = os.path.join(input_base_folder, "cedulas.xlsx") 
pdf_password = "" 

# Create output base folder if it doesn't exist
os.makedirs(output_base_folder, exist_ok=True)

# Create specific output subfolders
mc_output_folder = os.path.join(output_base_folder, "MC_Resultados")
visa_output_folder = os.path.join(output_base_folder, "Visa_Resultados")
os.makedirs(mc_output_folder, exist_ok=True)
os.makedirs(visa_output_folder, exist_ok=True)


# --- TRM Data Loading ---
trm_df = pd.DataFrame() # Initialize an empty DataFrame
trm_loaded = False

if os.path.exists(trm_file):
    try:
        trm_df = pd.read_excel(trm_file, sheet_name=trm_sheet)
        trm_df.columns = trm_df.columns.str.strip()
        trm_df["Fecha"] = pd.to_datetime(trm_df["Fecha"], errors='coerce')
        trm_df["TRM"] = trm_df["Tasa Representativa del Mercado (TRM)"].astype(str).str.replace(",", "").astype(float)
        trm_df = trm_df[["Fecha", "TRM"]]
        trm_loaded = True
        print(f"✅ TRM file '{trm_file}' loaded successfully.")
    except Exception as e:
        print(f"⚠️ Error loading TRM file '{trm_file}': {e}. MC currency conversion will not be available.")
else:
    print(f"⚠️ TRM file '{trm_file}' not found. MC currency conversion will not be available.")


def obtener_trm(fecha):
    if trm_loaded and pd.isna(fecha):
        return ""
    if trm_loaded:
        fila = trm_df[trm_df["Fecha"] == fecha]
        if not fila.empty:
            return fila["TRM"].values[0]
    return ""

def formato_excel(valor):
    try:
        if isinstance(valor, (int, float)):
            return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        numero = float(str(valor).replace(",", "").replace(" ", ""))
        return f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        return valor

# --- Categorias Data Loading ---
categorias_df = pd.DataFrame()
categorias_loaded = False

if os.path.exists(categorias_file):
    try:
        categorias_df = pd.read_excel(categorias_file)
        # Ensure 'Descripción' column exists and is stripped of whitespace
        if 'Descripción' in categorias_df.columns:
            categorias_df['Descripción'] = categorias_df['Descripción'].astype(str).str.strip()
            categorias_loaded = True
            print(f"✅ Categorias file '{categorias_file}' loaded successfully.")
        else:
            print(f"⚠️ Categorias file '{categorias_file}' loaded, but 'Descripción' column not found. Categorization will not be available.")
    except Exception as e:
        print(f"⚠️ Error loading Categorias file '{categorias_file}': {e}. Categorization will not be available.")
else:
    print(f"⚠️ Categorias file '{categorias_file}' not found. Categorization will not be available.")

# --- Cedulas Data Loading ---
cedulas_df = pd.DataFrame()
cedulas_loaded = False

if os.path.exists(cedulas_file):
    try:
        cedulas_df = pd.read_excel(cedulas_file)
        # Ensure 'Tarjetahabiente' column exists and convert to Title Case
        if 'Tarjetahabiente' in cedulas_df.columns:
            cedulas_df['Tarjetahabiente'] = cedulas_df['Tarjetahabiente'].astype(str).str.title().str.strip()
            cedulas_loaded = True
            print(f"✅ Cedulas file '{cedulas_file}' loaded successfully.")
        else:
            print(f"⚠️ Cedulas file '{cedulas_file}' loaded, but 'Tarjetahabiente' column not found. Cedula data will not be available.")
    except Exception as e:
        print(f"⚠️ Error loading Cedulas file '{cedulas_file}': {e}. Cedula data will not be available.")
else:
    print(f"⚠️ Cedulas file '{cedulas_file}' not found. Cedula data will not be available.")


# --- Regex for MC (from mc.py) ---
mc_transaccion_regex = re.compile(
    r"(\w{5,})\s+(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+)"
)
mc_nombre_regex = re.compile(r"SEÑOR \(A\):\s*(.*)")
mc_tarjeta_regex = re.compile(r"TARJETA:\s+\*{12}(\d{4})")
mc_moneda_regex = re.compile(r"ESTADO DE CUENTA EN:\s+(DOLARES|PESOS)")

# --- Regex for Visa (from visa.py) ---
visa_pattern_transaccion = re.compile(
    r"(\d{6})\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,.]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+|0\.00)"
)
visa_pattern_tarjeta = re.compile(r"TARJETA:\s+\*{12}(\d{4})")


resultados_mc = []
resultados_visa = []

# --- Process MC PDFs ---
mc_input_folder = os.path.join(input_base_folder, "MC")
if os.path.exists(mc_input_folder):
    for archivo in sorted(os.listdir(mc_input_folder)):
        if archivo.endswith(".pdf"):
            print(f"📄 Procesando MC: {archivo}")
            ruta_pdf = os.path.join(mc_input_folder, archivo)
            try:
                with fitz.open(ruta_pdf) as doc:
                    if doc.needs_pass:
                        doc.authenticate(pdf_password)

                    moneda_actual = ""
                    nombre = ""
                    ultimos_digitos = ""
                    tiene_transacciones_mc = False

                    for page_num, page in enumerate(doc, start=1):
                        texto = page.get_text()

                        moneda_match = mc_moneda_regex.search(texto)
                        if moneda_match:
                            moneda_actual = "USD" if moneda_match.group(1) == "DOLARES" else "COP"

                        if not nombre:
                            nombre_match = mc_nombre_regex.search(texto)
                            if nombre_match:
                                nombre = nombre_match.group(1).strip() # Get raw name

                        if not ultimos_digitos:
                            tarjeta_match = mc_tarjeta_regex.search(texto)
                            if tarjeta_match:
                                ultimos_digitos = tarjeta_match.group(1).strip()

                        for match in mc_transaccion_regex.finditer(texto):
                            autorizacion, fecha_str, descripcion, valor_original, tasa_pactada, tasa_ea, cargo, saldo, cuotas = match.groups()

                            if "ABONO DEBITO AUTOMATICO" in descripcion.upper():
                                continue

                            try:
                                fecha_transaccion = pd.to_datetime(fecha_str, dayfirst=True).date()
                            except:
                                fecha_transaccion = None

                            tipo_cambio = obtener_trm(pd.to_datetime(fecha_transaccion)) if moneda_actual == "USD" else ""
                            
                            resultados_mc.append({
                                "Archivo": archivo,
                                "Tarjetahabiente": nombre, # Keep raw name here, convert to title case later for merge
                                "Número de Tarjeta": ultimos_digitos,
                                "Moneda": moneda_actual,
                                "Tipo de Cambio": formato_excel(str(tipo_cambio)) if tipo_cambio else "",
                                "Número de Autorización": autorizacion,
                                "Fecha de Transacción": fecha_transaccion,
                                "Descripción": descripcion.strip(), # Ensure description is stripped for matching
                                "Valor Original": formato_excel(valor_original),
                                "Tasa Pactada": formato_excel(tasa_pactada),
                                "Tasa EA Facturada": formato_excel(tasa_ea),
                                "Cargos y Abonos": formato_excel(cargo),
                                "Saldo a Diferir": formato_excel(saldo),
                                "Cuotas": cuotas,
                                "Página": page_num,
                            })
                            tiene_transacciones_mc = True
                    
                    if not tiene_transacciones_mc and (nombre or ultimos_digitos): # Only add if we found a cardholder/card
                        resultados_mc.append({
                            "Archivo": archivo,
                            "Tarjetahabiente": nombre, # Keep raw name here, convert to title case later for merge
                            "Número de Tarjeta": ultimos_digitos,
                            "Moneda": "",
                            "Tipo de Cambio": "",
                            "Número de Autorización": "Sin transacciones",
                            "Fecha de Transacción": "",
                            "Descripción": "",
                            "Valor Original": "",
                            "Tasa Pactada": "",
                            "Tasa EA Facturada": "",
                            "Cargos y Abonos": "",
                            "Saldo a Diferir": "",
                            "Cuotas": "",
                            "Página": "",
                        })

            except Exception as e:
                print(f"⚠️ Error procesando MC '{archivo}': {e}")
else:
    print(f"⏩ Carpeta MC no encontrada en '{input_base_folder}'. Saltando procesamiento de MC.")


# --- Process Visa PDFs ---
visa_input_folder = os.path.join(input_base_folder, "Visa")
if os.path.exists(visa_input_folder):
    for pdf_file in sorted(os.listdir(visa_input_folder)):
        if pdf_file.lower().endswith(".pdf"):
            pdf_path = os.path.join(visa_input_folder, pdf_file)
            print(f"📄 Procesando Visa: {pdf_file}")

            try:
                with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
                    tarjetahabiente_visa = ""
                    tarjeta_visa = ""
                    tiene_transacciones_visa = False
                    last_page_number_visa = 1

                    for page_number, page in enumerate(pdf.pages, start=1):
                        text = page.extract_text()
                        if not text:
                            continue

                        last_page_number_visa = page_number
                        lines = text.split("\n")

                        for idx, line in enumerate(lines):
                            line = line.strip()

                            tarjeta_match_visa = visa_pattern_tarjeta.search(line)
                            if tarjeta_match_visa:
                                # Before updating card, if the previous card had no transactions, add a row
                                if tarjetahabiente_visa and tarjeta_visa and not tiene_transacciones_visa:
                                    resultados_visa.append({
                                        "Archivo": pdf_file,
                                        "Tarjetahabiente": tarjetahabiente_visa,
                                        "Número de Tarjeta": tarjeta_visa,
                                        "Moneda": "",
                                        "Tipo de Cambio": "",
                                        "Número de Autorización": "Sin transacciones",
                                        "Fecha de Transacción": "",
                                        "Descripción": "",
                                        "Valor Original": "",
                                        "Tasa Pactada": "",
                                        "Tasa EA Facturada": "",
                                        "Cargos y Abonos": "",
                                        "Saldo a Diferir": "",
                                        "Cuotas": "",
                                        "Página": last_page_number_visa,
                                    })
                                
                                tarjeta_visa = tarjeta_match_visa.group(1)
                                tiene_transacciones_visa = False # Reset for new card

                                if idx > 0:
                                    posible_nombre = lines[idx - 1].strip()
                                    posible_nombre = (
                                        posible_nombre
                                        .replace("SEÑOR (A):", "")
                                        .replace("Señor (A):", "")
                                        .replace("SEÑOR:", "")
                                        .replace("Señor:", "")
                                        .strip()
                                        #.title() # Not converting here, will convert df column later
                                    )
                                    if len(posible_nombre.split()) >= 2:
                                        tarjetahabiente_visa = posible_nombre
                                continue

                            match_visa = visa_pattern_transaccion.search(' '.join(line.split()))
                            if match_visa and tarjetahabiente_visa and tarjeta_visa:
                                autorizacion, fecha_str, descripcion, valor_original, tasa_pactada, tasa_ea, cargo, saldo, cuotas = match_visa.groups()

                                # Visa specific numeric formatting
                                valor_original_formatted = valor_original.replace(".", "").replace(",", ".")
                                cargo_formatted = cargo.replace(".", "").replace(",", ".")
                                saldo_formatted = saldo.replace(".", "").replace(",", ".")

                                resultados_visa.append({
                                    "Archivo": pdf_file,
                                    "Tarjetahabiente": tarjetahabiente_visa, # Keep raw name here, convert to title case later for merge
                                    "Número de Tarjeta": tarjeta_visa,
                                    "Moneda": "COP", # Assuming Visa are in COP as no currency explicit extraction
                                    "Tipo de Cambio": "", # Not applicable for COP
                                    "Número de Autorización": autorizacion,
                                    "Fecha de Transacción": pd.to_datetime(fecha_str, dayfirst=True).date() if fecha_str else None,
                                    "Descripción": descripcion.strip(), # Ensure description is stripped for matching
                                    "Valor Original": formato_excel(valor_original_formatted),
                                    "Tasa Pactada": formato_excel(tasa_pactada),
                                    "Tasa EA Facturada": formato_excel(tasa_ea),
                                    "Cargos y Abonos": formato_excel(cargo_formatted),
                                    "Saldo a Diferir": formato_excel(saldo_formatted),
                                    "Cuotas": cuotas,
                                    "Página": page_number,
                                })
                                tiene_transacciones_visa = True
                    
                    # After processing all pages for a Visa PDF, check if no transactions were found for the last card processed
                    if tarjetahabiente_visa and tarjeta_visa and not tiene_transacciones_visa:
                        resultados_visa.append({
                            "Archivo": pdf_file,
                            "Tarjetahabiente": tarjetahabiente_visa, # Keep raw name here, convert to title case later for merge
                            "Número de Tarjeta": tarjeta_visa,
                            "Moneda": "",
                            "Tipo de Cambio": "",
                            "Número de Autorización": "Sin transacciones",
                            "Fecha de Transacción": "",
                            "Descripción": "",
                            "Valor Original": "",
                            "Tasa Pactada": "",
                            "Tasa EA Facturada": "",
                            "Cargos y Abonos": "",
                            "Saldo a Diferir": "",
                            "Cuotas": "",
                            "Página": last_page_number_visa,
                        })

            except Exception as e:
                print(f"⚠️ Error al procesar Visa '{pdf_file}': {e}")
else:
    print(f"⏩ Carpeta Visa no encontrada en '{input_base_folder}'. Saltando procesamiento de Visa.")

# --- Save MC Results ---
if resultados_mc:
    df_resultado_mc = pd.DataFrame(resultados_mc)
    
    # Convert 'Tarjetahabiente' to Title Case for merging with cedulas_df
    df_resultado_mc['Tarjetahabiente'] = df_resultado_mc['Tarjetahabiente'].astype(str).str.title().str.strip()

    # Merge with categorias_df if loaded
    if categorias_loaded:
        print("Merging MC results with categorias.xlsx...")
        df_resultado_mc = pd.merge(df_resultado_mc, categorias_df[['Descripción', 'Categoría', 'Subcategoría', 'Zona']],
                                   on='Descripción', how='left')
    else:
        # Add empty columns if categorias.xlsx was not loaded
        df_resultado_mc['Categoría'] = ''
        df_resultado_mc['Subcategoría'] = ''
        df_resultado_mc['Zona'] = ''

    # Merge with cedulas_df if loaded
    if cedulas_loaded:
        print("Merging MC results with cedulas.xlsx...")
        df_resultado_mc = pd.merge(df_resultado_mc, cedulas_df[['Tarjetahabiente', 'Cédula', 'Tipo', 'Cargo']],
                                   on='Tarjetahabiente', how='left')
    else:
        # Add empty columns if cedulas.xlsx was not loaded
        df_resultado_mc['Cédula'] = ''
        df_resultado_mc['Tipo'] = ''
        df_resultado_mc['Cargo'] = ''

    fecha_hora_salida = datetime.now().strftime("%Y%m%d_%H%M")
    archivo_salida_mc = f"MC_{fecha_hora_salida}.xlsx"
    ruta_salida_mc = os.path.join(mc_output_folder, archivo_salida_mc)
    df_resultado_mc.to_excel(ruta_salida_mc, index=False)
    print(f"\n✅ Archivo MC generado correctamente en:\n{ruta_salida_mc}")
    print("\nPrimeras 5 filas del resultado MC:")
    print(df_resultado_mc.head())
else:
    print("\n⚠️ No se extrajo ningún dato de los archivos MC.")

# --- Save Visa Results ---
if resultados_visa:
    df_resultado_visa = pd.DataFrame(resultados_visa)

    # Convert 'Tarjetahabiente' to Title Case for merging with cedulas_df
    df_resultado_visa['Tarjetahabiente'] = df_resultado_visa['Tarjetahabiente'].astype(str).str.title().str.strip()
    
    # Merge with categorias_df if loaded
    if categorias_loaded:
        print("Merging Visa results with categorias.xlsx...")
        df_resultado_visa = pd.merge(df_resultado_visa, categorias_df[['Descripción', 'Categoría', 'Subcategoría', 'Zona']],
                                   on='Descripción', how='left')
    else:
        # Add empty columns if categorias.xlsx was not loaded
        df_resultado_visa['Categoría'] = ''
        df_resultado_visa['Subcategoría'] = ''
        df_resultado_visa['Zona'] = ''

    # Merge with cedulas_df if loaded
    if cedulas_loaded:
        print("Merging Visa results with cedulas.xlsx...")
        df_resultado_visa = pd.merge(df_resultado_visa, cedulas_df[['Tarjetahabiente', 'Cédula', 'Tipo', 'Cargo']],
                                   on='Tarjetahabiente', how='left')
    else:
        # Add empty columns if cedulas.xlsx was not loaded
        df_resultado_visa['Cédula'] = ''
        df_resultado_visa['Tipo'] = ''
        df_resultado_visa['Cargo'] = ''

    fecha_hora_salida = datetime.now().strftime("%Y%m%d_%H%M")
    archivo_salida_visa = f"VISA_{fecha_hora_salida}.xlsx"
    ruta_salida_visa = os.path.join(visa_output_folder, archivo_salida_visa)
    df_resultado_visa.to_excel(ruta_salida_visa, index=False)
    print(f"\n✅ Archivo VISA generado correctamente en:\n{ruta_salida_visa}")
    print("\nPrimeras 5 filas del resultado VISA:")
    print(df_resultado_visa.head())
else:
    print("\n⚠️ No se extrajo ningún dato de los archivos VISA.")
"@
}

TC