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
    python -m pip install --upgrade pip
    python -m pip install PyMuPDF pandas pdfplumber openpyxl

    # Create templates directory structure
    $directories = @(
        "PDFS",
        "Resultados"
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
input_base_folder = "src"
output_base_folder = "Resultados" # Base folder for all results
trm_file = os.path.join(input_base_folder, "TRM.xlsx")
trm_sheet = "Datos"
categorias_file = os.path.join(input_base_folder, "categorias.xlsx")
cedulas_file = os.path.join(input_base_folder, "cedulas.xlsx") 
pdf_password = "" 

# Create output base folder if it doesn't exist
os.makedirs(output_base_folder, exist_ok=True)

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


all_resultados = [] # Combined list for all results

# --- Process all PDFs in the single input folder ---
if os.path.exists(input_base_folder):
    for archivo in sorted(os.listdir(input_base_folder)):
        if archivo.endswith(".pdf"):
            ruta_pdf = os.path.join(input_base_folder, archivo)
            
            # Use file name to determine card type
            card_type_is_mc = "MC" in archivo.upper() or "MASTERCARD" in archivo.upper()
            card_type_is_visa = "VISA" in archivo.upper()
            
            if card_type_is_mc:
                print(f"📄 Procesando Mastercard: {archivo}")
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
                                
                                all_resultados.append({
                                    "Archivo": archivo,
                                    "Tipo de Tarjeta": "Mastercard", # New column
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
                            all_resultados.append({
                                "Archivo": archivo,
                                "Tipo de Tarjeta": "Mastercard", # New column
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
            
            elif card_type_is_visa:
                print(f"📄 Procesando Visa: {archivo}")
                try:
                    with pdfplumber.open(ruta_pdf, password=pdf_password) as pdf:
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
                                        all_resultados.append({
                                            "Archivo": archivo,
                                            "Tipo de Tarjeta": "Visa", # New column
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

                                    all_resultados.append({
                                        "Archivo": archivo,
                                        "Tipo de Tarjeta": "Visa", # New column
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
                            all_resultados.append({
                                "Archivo": archivo,
                                "Tipo de Tarjeta": "Visa", # New column
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
                    print(f"⚠️ Error al procesar Visa '{archivo}': {e}")
            else:
                print(f"⏩ Archivo '{archivo}' no reconocido como Mastercard o Visa. Saltando.")

else:
    print(f"⏩ Carpeta de origen '{input_base_folder}' no encontrada. No hay archivos para procesar.")


# --- Save All Results to a Single Excel File ---
if all_resultados:
    df_resultado_final = pd.DataFrame(all_resultados)
    
    # Convert 'Tarjetahabiente' to Title Case for merging with cedulas_df
    df_resultado_final['Tarjetahabiente'] = df_resultado_final['Tarjetahabiente'].astype(str).str.title().str.strip()

    # Convert 'Fecha de Transacción' to datetime objects to enable day name extraction
    df_resultado_final['Fecha de Transacción'] = pd.to_datetime(df_resultado_final['Fecha de Transacción'], errors='coerce')
    
    # Add the 'Día' column
    # Ensure it handles NaT values gracefully, perhaps by filling with empty string
    df_resultado_final['Día'] = df_resultado_final['Fecha de Transacción'].dt.day_name(locale='es_ES').fillna('') # Use 'es_ES' for Spanish day names
    
    # Add the new 'Cant. de Tarjetas' column
    # This groups by 'Tarjetahabiente' and counts the number of unique 'Número de Tarjeta' for each person.
    # The transform() function ensures the result is a Series with the same index as the original DataFrame,
    # so it can be directly assigned as a new column.
    df_resultado_final['Cant. de Tarjetas'] = df_resultado_final.groupby('Tarjetahabiente')['Número de Tarjeta'].transform('nunique')

    # Merge with categorias_df if loaded
    if categorias_loaded:
        print("Merging all results with categorias.xlsx...")
        df_resultado_final = pd.merge(df_resultado_final, categorias_df[['Descripción', 'Categoría', 'Subcategoría', 'Zona']],
                                   on='Descripción', how='left')
    else:
        # Add empty columns if categorias.xlsx was not loaded
        df_resultado_final['Categoría'] = ''
        df_resultado_final['Subcategoría'] = ''
        df_resultado_final['Zona'] = ''

    # Merge with cedulas_df if loaded
    if cedulas_loaded:
        print("Merging all results with cedulas.xlsx...")
        df_resultado_final = pd.merge(df_resultado_final, cedulas_df[['Tarjetahabiente', 'Cédula', 'Tipo', 'Cargo']],
                                   on='Tarjetahabiente', how='left')
    else:
        # Add empty columns if cedulas.xlsx was not loaded
        df_resultado_final['Cédula'] = ''
        df_resultado_final['Tipo'] = ''
        df_resultado_final['Cargo'] = ''

    # Define the desired column order, placing 'Día' after 'Fecha de Transacción' and 'Cant. de Tarjetas' after 'Número de Tarjeta'
    
    # Get the base columns that are always present or added by extraction
    base_columns = [
        "Archivo",
        "Tipo de Tarjeta",
        "Tarjetahabiente",
        "Número de Tarjeta",
        "Cant. de Tarjetas", # The new column
        "Moneda",
        "Tipo de Cambio",
        "Número de Autorización",
        "Fecha de Transacción",
        "Día",
        "Descripción",
        "Valor Original",
        "Tasa Pactada",
        "Tasa EA Facturada",
        "Cargos y Abonos",
        "Saldo a Diferir",
        "Cuotas",
        "Página"
    ]

    # Add columns from merges if they exist in the final DataFrame
    if 'Categoría' in df_resultado_final.columns:
        base_columns.extend(['Categoría', 'Subcategoría', 'Zona'])
    if 'Cédula' in df_resultado_final.columns:
        base_columns.extend(['Cédula', 'Tipo', 'Cargo'])

    # Filter to only include columns that actually exist in the DataFrame to avoid errors
    final_columns_order = [col for col in base_columns if col in df_resultado_final.columns]
    df_resultado_final = df_resultado_final[final_columns_order]


    fecha_hora_salida = datetime.now().strftime("%Y%m%d_%H%M")
    archivo_salida_unificado = f"TC_{fecha_hora_salida}.xlsx"
    ruta_salida_unificado = os.path.join(output_base_folder, archivo_salida_unificado)
    df_resultado_final.to_excel(ruta_salida_unificado, index=False)
    print(f"\n✅ Archivo unificado de extractos generado correctamente en:\n{ruta_salida_unificado}")
    print("\nPrimeras 5 filas del resultado unificado:")
    print(df_resultado_final.head())
else:
    print("\n⚠️ No se extrajo ningún dato de los archivos PDF (MC o VISA).")
"@
}

TC