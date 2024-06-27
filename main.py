import pandas as pd

# Leer el archivo Excel y cargarlo en un DataFrame
df = pd.read_excel('original.xlsx')
df_original = df

# Convertir las columnas a tipos adecuados
df['Tracking Number (HAWB)'] = df['Tracking Number (HAWB)'].astype(str)
df['TOTAL QTY OF ITEMS IN PARCEL'] = df['TOTAL QTY OF ITEMS IN PARCEL'].astype(float)
df['TOTAL DECLARED VALUE'] = df['TOTAL DECLARED VALUE'].astype(float)

# Agregar la columna IVA
df['IVA'] = df['TOTAL DECLARED VALUE'].apply(lambda x: 0.0 if x < 50.01 else (0.17 if 50.01 <= x <= 117.01 else 0.19))

# Mapear los shippers
ships = {"IMEX - Mattel One Shop": "JUGUETE", "FragranceNet.com": "PERFUME"}
df['SHORT DESCRIPTION'] = df['SHIPPER'].map(lambda x: ships.get(x, ""))

# Definir condiciones de filtrado
condiciones_filtrado = (df['Tracking Number (HAWB)'].str.len() == 22) | (df['TOTAL QTY OF ITEMS IN PARCEL'] > 10) | (df['TOTAL DECLARED VALUE'] >= 500) | (df['PRODUCT DESCRIPTION'].str.contains(r'\bother\b', case=False))

# Separar registros especiales y normales
df_especiales = df[condiciones_filtrado]
df_normal = df[~condiciones_filtrado].sort_values(by='TOTAL DECLARED VALUE')

# Separar registros normales en menores y mayores
df_menores = df_normal[df_normal['TOTAL DECLARED VALUE'] < 50.01].reset_index(drop=True)
df_mayores = df_normal[df_normal['TOTAL DECLARED VALUE'] >= 50.01].reset_index(drop=True)

# Crear bloques de registros mayores
limite = 5000
bloques = []
sumatoria = 0.0
inicio_b = 0
for index, log in df_mayores.iterrows():
    actual_price = float(log['TOTAL DECLARED VALUE'])
    if (sumatoria + actual_price) <= limite:
        sumatoria += actual_price
    else:
        bloques.append((inicio_b, index))
        inicio_b = index
        sumatoria = actual_price
bloques.append((inicio_b, len(df_mayores)))

# Crear DataFrames de secciones
agrupacion_dfs = {"MENORES": df_menores}
for i, bloque in enumerate(bloques):
    identificador = f"MAYORES {i + 1}"
    agrupacion_dfs[identificador] = df_mayores.iloc[bloque[0]:bloque[1]]
agrupacion_dfs['ESPECIALES'] = df_especiales

# Inicializar df_final
df_final = pd.DataFrame()

# Agrupamos los bloques obtenidos
def agrupar_bloque(titulo, bloque_df):
  df_title = pd.DataFrame({'GRUPO': [titulo]})
  return pd.concat([df_title, bloque_df], ignore_index=True)


# Concatenar DataFrames en df_final
for key, value in agrupacion_dfs.items():
    df_final = pd.concat([df_final, agrupar_bloque(key, value)], ignore_index=True)

# Escribir los DataFrames en un archivo Excel
nombre_archivo = 'datos_filtrados.xlsx'

# Crear un escritor de Excel
writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')

# Escribir cada DataFrame en una hoja diferente del archivo de Excel
df_final.to_excel(writer, sheet_name='Separados', index=False)
df_especiales.to_excel(writer, sheet_name='Especiales', index=False)
df_normal.to_excel(writer, sheet_name='Normales', index=False)
df_mayores.to_excel(writer, sheet_name='Mayores', index=False)
df_menores.to_excel(writer, sheet_name='Menores', index=False)
df_original.to_excel(writer, sheet_name='Originales', index=False)


# Estilos para las hojas (backgrounds)
worksheet = writer.sheets['Separados']
worksheet.set_tab_color('orange')
worksheet = writer.sheets['Especiales']
worksheet.set_tab_color('purple')
worksheet = writer.sheets['Normales']
worksheet.set_tab_color('red')
worksheet = writer.sheets['Mayores']
worksheet.set_tab_color('blue')
worksheet = writer.sheets['Menores']
worksheet.set_tab_color('yellow')

# Encontrar los índices de los valores distintos de NaN en la columna 'GRUPOS'
titulos = df_final[df_final['GRUPO'].notna()].index

workbook  = writer.book
worksheet = writer.sheets['Separados']

# Definir un formato para la fila
formato_fila = workbook.add_format({'bg_color': 'pink', 'bold': True, 'font_size': 26})

# Ajustar el ancho de la fila y aplicar el formato a cada fila con titulo
for i in titulos:
  worksheet.set_row(i + 1, 35, formato_fila)

# Cerrar el escritor
writer.close()
