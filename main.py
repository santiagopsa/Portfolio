import os
import logging
from tkinter.filedialog import Open
import pandas as pd
import dotenv
from openai import OpenAI
from fpdf import FPDF

# Enviromental variables
dotenv.load_dotenv()

# Logging configuration
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configure OpenAI client
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    raise ValueError("La variable de entorno OPENAI_API_KEY no está configurada")
client = OpenAI(api_key=OPENAI_API_KEY)

# Configuration constants
FILE_PATH = "statements/bancolombia.xlsx"
CLEAN_FILE_PATH = "statements/bancolombia_clean.xlsx"
DINAMICA_FILE_PATH = "statements/bancolombia_dinamica.xlsx"
GPT_MODEL = "gpt-4o-mini"  # Ajusta el modelo según tus necesidades

# Pdf output report generation
def generate_pdf_report(analysis_result: str, output_path: str):
    """
    Generates a PDF report from the analysis result.
    
    Args:
        analysis_result: The analysis result text to include in the PDF.
        output_path: The file path where the PDF will be saved.
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Add a title
    pdf.cell(200, 10, txt="Financial Analysis Report", ln=True, align='C')
    
    # Add the analysis result
    pdf.multi_cell(0, 10, txt=analysis_result)
    
    # Save the PDF to the specified path
    pdf.output(output_path)
    logger.info(f"PDF report saved to {output_path}.")

# Organize the data so it can be converted to a float
def convert_amount(x):
    """
    Converts an Excel cell value representing an amount to a float.
    
    This function handles:
      - Removal of currency symbols (adjust symbols as needed).
      - Removal of thousand separators (assumes comma as thousand separator).
      - Conversion of numbers with a trailing minus sign (e.g., "123.45-") 
        into the proper format (i.e., "-123.45").
    
    Args:
        x: The cell value to convert.
        
    Returns:
        The converted float value or None if conversion fails.
    """
    if pd.isnull(x):
        return None
    
    # Convert to string and strip whitespace
    s = str(x).strip()
    
    # Remove common currency symbols (adjust the list as needed)
    for symbol in ['$', '€', '£']:
        s = s.replace(symbol, '')
    
    # Remove thousand separators (assumes comma as thousand separator)
    s = s.replace(',', '')
    
    # Check for a trailing minus sign (e.g., "123.45-") and adjust it to standard notation ("-123.45")
    if s.endswith('-'):
        s = '-' + s[:-1].strip()
    
    # If, for any reason, the negative sign is wrapped in parentheses (optional, in case some values use that)
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1].strip()
    
    try:
        return float(s)
    except Exception as e:
        logger.error(f"Error converting value '{x}': {e}")
        return None
    
# This function reads an excel file and returns a dataframe
def convert_xlsx_to_df(file_path: str) -> pd.DataFrame:
    """
    Lee un archivo Excel y devuelve un DataFrame.
    
    Args:
        file_path: Ruta del archivo Excel.
    
    Returns:
        pd.DataFrame con el contenido del archivo.
    """
    try:
        df = pd.read_excel(file_path)
        logger.info(f"Archivo {file_path} cargado correctamente.")
        return df
    except Exception as e:
        logger.error(f"Error al cargar el archivo {file_path}: {e}")
        raise

# This function analyze a traditional excel file balance sheet and needs to be adjusted according the bank statement format
def extract_tables(raw_df: pd.DataFrame) -> list[pd.DataFrame]:
    """
    Extrae secciones o tablas del DataFrame basándose en marcas en los datos.
    Se identifica el inicio de una tabla cuando se encuentra la palabra 'movimientos'
    en la primera columna, y el final cuando en la sexta columna aparece un valor nulo.
    
    Args:
        raw_df: DataFrame original leído del archivo Excel.
    
    Returns:
        Lista de DataFrames correspondientes a cada tabla extraída.
    """
    tables = []
    table_start = None  # Índice donde inicia una tabla

    for index, row in raw_df.iterrows():
        first_cell = str(row.iloc[0]).lower() if pd.notnull(row.iloc[0]) else ""
        # Detectar inicio de una tabla
        if "movimientos" in first_cell:
            logger.info(f"Se encontró 'movimientos' en la fila {index}.")
            table_start = index + 1
        # Detectar fin de la tabla: se asume que la columna 5 es la clave
        elif table_start is not None and pd.isna(row.iloc[5]):
            logger.info(f"Fin de la tabla detectado en la fila {index}.")
            table = raw_df.iloc[table_start:index]
            tables.append(table)
            table_start = None  # Reiniciamos el índice de inicio para la siguiente tabla
    return tables

# This function clean the combined dataframe and prepare it for the analysis
def clean_combined_df(combined_df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia el DataFrame combinado:
    - Utiliza la primera fila como encabezado.
    - Elimina filas redundantes que contengan 'FECHA' en la primera columna (a partir de la segunda fila).
    - Convierte la columna de montos a valores numéricos para diferenciar ingresos (positivos)
      y egresos (negativos).
    
    Args:
        combined_df: DataFrame concatenado de todas las tablas.
    
    Returns:
        DataFrame limpio y listo para analizar.
    """
    # Asignar la primera fila como encabezado
     # Set the first row as the header and drop that row from the data
    combined_df.columns = combined_df.iloc[0].str.lower()
    combined_df = combined_df.drop(0).reset_index(drop=True)
    
    # Remove rows that contain redundant headers (assuming 'FECHA' appears in the first column)
    mask = ~combined_df.iloc[1:, 0].str.contains("FECHA", case=False, na=False)
    first_row = combined_df.iloc[:1]
    filtered_rows = combined_df.iloc[1:][mask]
    cleaned_df = pd.concat([first_row, filtered_rows], ignore_index=True)
    
    # Convert the amounts column using the custom converter.
    # Check if the column is named 'Valor'. If not, assume the amounts are in the 3rd column (index 2)
    if 'Valor' in cleaned_df.columns:
        cleaned_df['Valor'] = cleaned_df['Valor'].apply(convert_amount)
    else:
        cleaned_df.iloc[:, 4] = cleaned_df.iloc[:, 4].apply(convert_amount)
    
    return cleaned_df

# Function to analyze the data and generate a report with GPT, to use it data needs to be clean and specific prompt to provide valuabla data is given
def analyze_data(dataframe):
    """
    Combines market data from different timeframes and additional analysis,
    then prepares a prompt for GPT to analyze.
    """
    prompt = f"""
        El el siguiente cuadro hay un resumen de gastos de 3 meses, la primera columna es la descripcion del gasto, la segunda el monto y la tercera la frecuencia.
        los valores negativos son gastos
        En base a la recurrencia se puede inferir que los gastos que se repiten 3 veces son fijos y el resto pueden o no ser variables, dependera de la descripcion del gasto.
        Quiero que des una recomendacion de finanzas personales, teniendo en cuenta que los gastos fijos son dificiles de reducir y entraría en un plan agresivo de reduccion y los gastos variables si se pueden bajar tomando medidas para ahorrar.
        enfoquemonos en los 10 gastos mayores
        Los ingresos mensuales son de 15000000 pesos colombianos
        Ademas quiero que hagamos un plan estructurado a 2 meses para empezar a reducir gastos variables y como se vería ese ahorro proyectado en el tiempo, ademas una proyeccion de como se vería el ahorro si se pone en un fonde de inversion de alta rentabilidad en un periodo de 5 años
        el output de tu respuesta debe ser un para word listo para convertir a pdf con un formato bonito y legible
        {dataframe}
    """
    response = client.chat.completions.create(
        model=GPT_MODEL,
        messages=[
            {"role": "developer", "content": "Eres un experto en finanzas personales de hogares y familias."},
            {"role": "user", "content": prompt}
        ],
        temperature=0
    )
    return response.choices[0].message.content.strip()

# Create a new dataframe calle "dinamica" to group expenses by description and frecuency to provided better data to analyze function
def summarize_and_sort_dinamica(dinamica_df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrupa el DataFrame por la columna de descripción, suma los valores, cuenta la recurrencia y ordena de mayor a menor.
    
    Args:
        dinamica_df: DataFrame que contiene la tabla dinámica.
    
    Returns:
        DataFrame agrupado, con columna de recurrencia y ordenado.
    """
    # Debugging: Print column names and first few rows before groupby
    print("Before groupby - Column names in dinamica_df:", dinamica_df.columns)
    print("Before groupby - First few rows of dinamica_df:\n", dinamica_df.head())

    # Asegúrate de que la columna de descripción y la de valores están correctamente identificadas
    dinamica_df.columns = dinamica_df.columns.str.strip()

    grouped_df = dinamica_df.groupby('descripción').agg(
        valor_sum=('valor', 'sum'),
        recurrencia=('valor', 'size')
    ).reset_index()
    
    sorted_df = grouped_df.sort_values(by='valor_sum', ascending=False)
    return sorted_df

# Main function to run the program
def main():
    # Cargar datos del archivo Excel
    raw_df = convert_xlsx_to_df(FILE_PATH)
    
    # Extraer las tablas del DataFrame original
    tables = extract_tables(raw_df)
    if not tables:
        logger.warning("No se encontraron tablas en el archivo.")
        return
    
    # Combinar todas las tablas en un solo DataFrame
    combined_df = pd.concat(tables, ignore_index=True)
    cleaned_df = clean_combined_df(combined_df)
    dinamica_df=summarize_and_sort_dinamica(cleaned_df)
    # Guardar el DataFrame limpio en un archivo Excel
    cleaned_df.to_excel(CLEAN_FILE_PATH, index=False)
    dinamica_df.to_excel(DINAMICA_FILE_PATH, index=False)
    logger.info(f"Archivo limpio guardado en {CLEAN_FILE_PATH}.")
    
    
    # Analizar los datos a través de la API de OpenAI
    analysis_result = analyze_data(dinamica_df)

    print("Resultado del análisis:\n", analysis_result)
    pdf_output_path = "financial_analysis_report.pdf"
    generate_pdf_report(analysis_result, pdf_output_path)

# Make sure the program runs only when the file is executed directly
if __name__ == "__main__":
    main()