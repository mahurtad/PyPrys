import pandas as pd

# Load the specific sheet and specify columns of interest
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'
columns_of_interest = ['N° GL', 'N° CP', 'N° D', 'Fecha Replanificada CP', 'BPO', 'PM Asignado', 'Periférico', 'Estado CP', 'Hora de Inicio', 'Hora de Fin']
data = pd.read_excel(file_path, sheet_name='Plan Pruebas Integrales', usecols=columns_of_interest)

# Ensure date columns are datetime type
data['Fecha Replanificada CP'] = pd.to_datetime(data['Fecha Replanificada CP'], errors='coerce')

# Define the filter function with optional multiple "N° GL" and "Estado CP" values
def filter_data(df, start_date, end_date, gl_numbers=None, estados_cp=None):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    filtered_df = df[(df['Fecha Replanificada CP'] >= start_date) & 
                     (df['Fecha Replanificada CP'] <= end_date)]
    
    # Apply optional filter for multiple "N° GL" values if provided
    if gl_numbers:
        filtered_df = filtered_df[filtered_df['N° GL'].isin(gl_numbers)]
    
    # Apply optional filter for multiple "Estado CP" values if provided
    if estados_cp:
        filtered_df = filtered_df[filtered_df['Estado CP'].isin(estados_cp)]
    
    return filtered_df

# Example usage
start_date = '2024-10-14'
end_date = '2024-10-18'
gl_numbers = [1, 3]  # Optional: List of "N° GL" values to filter by, e.g., [1, 3]
estados_cp = ['Pendiente retest', 'Planificado']  # Optional: List of states to filter by
filtered_data = filter_data(data, start_date, end_date, gl_numbers, estados_cp)

# Display the filtered data
print(filtered_data)

# Display the count of unique "N° CP" values in the filtered data
unique_cp_count = filtered_data['N° CP'].nunique()
print(f"\nTotal unique 'N° CP' records: {unique_cp_count}")
