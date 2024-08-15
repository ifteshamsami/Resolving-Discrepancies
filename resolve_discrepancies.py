import pandas as pd
file_path = r"C:\Users\mkd\Desktop\Programming\Projects\Resolving Discrepancies\Interest Rates Map (K-Drive).xlsm"
df_ms = pd.read_excel(file_path, sheet_name='MS', header=2)  # Update with correct sheet name
df_pm = pd.read_excel(file_path, sheet_name='PM', header=2)  # Update with correct sheet name
required_columns_ms = {'ID', 'Name'}
required_columns_pm = {'ID', 'Name', 'Row Index'}
if not required_columns_ms.issubset(df_ms.columns):
    raise ValueError(f"The 'MS' sheet is missing one or more required columns: {required_columns_ms}")

if not required_columns_pm.issubset(df_pm.columns):
    raise ValueError(f"'PM' sheet is missing one or more required columns: {required_columns_pm}")


if 'Row Index' not in df_pm.columns:
    df_pm['Row Index'] = range(1, len(df_pm) + 1)
    print("Added 'Row Index' column to 'PM' sheet.")


ms_names = df_ms.groupby('ID')['Name'].apply(lambda x: list(x.unique())).reset_index()
pm_names = df_pm.groupby('ID')['Name'].apply(lambda x: list(x.unique())).reset_index()


merged_names = pd.merge(ms_names, pm_names, on='ID', how='outer', suffixes=('_MS', '_PM'))


merged_names['Name_MS'] = merged_names['Name_MS'].apply(lambda x: x if isinstance(x, list) else [])
merged_names['Name_PM'] = merged_names['Name_PM'].apply(lambda x: x if isinstance(x, list) else [])


def compare_name_arrays(row):
    names_ms = set(row['Name_MS'])
    names_pm = set(row['Name_PM'])
    return names_ms != names_pm


discrepancies = merged_names[merged_names.apply(compare_name_arrays, axis=1)]

# Filter out rows where either 'Name_MS' or 'Name_PM' is empty
discrepancies = discrepancies[
    (discrepancies['Name_MS'].apply(lambda x: len(x) > 0)) &
    (discrepancies['Name_PM'].apply(lambda x: len(x) > 0))
]


discrepancies['Master'] = discrepancies['Name_MS'].apply(lambda x: ', '.join(x) if x else 'No Name')
discrepancies['People Moves'] = discrepancies['Name_PM'].apply(lambda x: ', '.join(x) if x else 'No Name')


discrepancies = discrepancies[['ID', 'Master', 'People Moves']]


print("Discrepancies found:")
print(discrepancies)


output_path = r'C:\Users\mkd\Desktop\Programming\Projects\Resolving Discrepancies\discrepancies_output.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    discrepancies.to_excel(writer, index=False, sheet_name='Discrepancies')
    
    
    workbook = writer.book
    worksheet = writer.sheets['Discrepancies']
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"Discrepancies exported to {output_path}")
