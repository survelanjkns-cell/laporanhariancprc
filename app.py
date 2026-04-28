# Define the fixed order and list of PKDs as per your reference image
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

def process_excel_data(excel_path):
    df = pd.read_excel(excel_path)
    
    # 1. Filter: Exclude "Abai Notifikasi"
    df = df[df['Notifikasi Status'] != 'Abai Notifikasi']
    
    # 2. Filter: Only include PKDs that exist in your template
    # This ignores any "New PKD" or data outside the 9 districts
    df = df[df['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
    
    # 3. Pivot: Matrix Table
    matrix_df = pd.crosstab(df['Diagnosis'], df['Pejabat Kesihatan'], rownames=['PENYAKIT'])
    
    # 4. Reindex columns to match the template order exactly
    # This ensures even if a PKD has 0 cases, it still appears in the table
    matrix_df = matrix_df.reindex(columns=TEMPLATE_PKDS, fill_value=0)
    
    # 5. Calculate Grand Totals
    matrix_df['Grand Total'] = matrix_df.sum(axis=1)
    col_sums = matrix_df.sum(axis=0)
    
    return matrix_df, col_sums
