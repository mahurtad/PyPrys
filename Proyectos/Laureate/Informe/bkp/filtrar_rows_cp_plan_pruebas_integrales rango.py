import pandas as pd

# Load the specific sheet
file_path = r'C:\Users\manue\OneDrive - EduCorpPERU\UPC - Calidad de Software\Proyectos\Pruebas Integrales\Planificación de Pruebas Integrales GL1 y GL3 v3.xlsx'
data = pd.read_excel(file_path, sheet_name='Plan Pruebas Integrales')

# Ensure date columns are datetime type
data['Fecha Replanificada CP'] = pd.to_datetime(data['Fecha Replanificada CP'], errors='coerce')

# Define the filter function
def filter_data(df, start_date, end_date, gl_number):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    filtered_df = df[(df['Fecha Replanificada CP'] >= start_date) & 
                     (df['Fecha Replanificada CP'] <= end_date) & 
                     (df['N° GL'] == gl_number)]
    return filtered_df

# Example usage
start_date = '2024-05-22'
end_date = '2024-06-10'
gl_number = 1
filtered_data = filter_data(data, start_date, end_date, gl_number)

# Display or save the filtered data
print(filtered_data)
