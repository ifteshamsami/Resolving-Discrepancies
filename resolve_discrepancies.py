import pandas as pd
file_path = r"C:\Users\mkd\Desktop\Programming\Projects\Resolving Discrepancies\Interest Rates Map (K-Drive).xlsm"  # Update this with the path to your file


df_ms = pd.read_excel(file_path, sheet_name='MS', header=2)  
df_pm = pd.read_excel(file_path, sheet_name='PM', header=2)  


required_columns_ms = {'ID', 'Name'}
required_columns_pm = {'ID', 'Name', 'Row Index'}


if not required_columns_ms.issubset(df_ms.columns):
    print("The 'MS' sheet might have its header row at a different position. Here are the current columns:")
    print(df_ms.columns)
    raise ValueError("The 'MS' sheet is missing one or more required columns: {}".format(required_columns_ms))

if not required_columns_pm.issubset(df_pm.columns):
    print("The 'PM' sheet might have its header row at a different position. Here are the current columns:")
    print(df_pm.columns)
    raise ValueError("The 'PM' sheet is missing one or more required columns: {}".format(required_columns_pm))


if 'Row Index' not in df_pm.columns:
    df_pm['Row Index'] = range(1, len(df_pm) + 1)
    print("Added 'Row Index' column to 'PM' sheet.")

ms_names = df_ms.groupby('ID')['Name'].apply(lambda x: list(x.unique())).reset_index()
pm_names = df_pm.groupby('ID')['Name'].apply(lambda x: list(x.unique())).reset_index()


merged_names = pd.merge(ms_names, pm_names, on='ID', how='outer', suffixes=('_MS', '_PM'))


merged_names['Name_MS'] = merged_names['Name_MS'].apply(lambda x: x if isinstance(x, list) else [])
merged_names['Name_PM'] = merged_names['Name_PM'].apply(lambda x: x if isinstance(x, list) else [])


def compare_name_arrays(row):
    return set(row['Name_MS']) != set(row['Name_PM'])

discrepancies = merged_names[merged_names.apply(compare_name_arrays, axis=1)]


print(discrepancies)


ids_with_discrepancies = discrepancies['ID'].tolist()

print("IDs with discrepancies:")
for id_ in ids_with_discrepancies:
    print(id_)

discrepancies.to_excel('discrepancies_output.xlsx', index=False)
