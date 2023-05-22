import pandas as pd
import pyperclip
import io
import re
import os

# Leer los datos del portapapeles
clipboard_data = pyperclip.paste()

# Convertir los datos en una tabla de pandas
df = pd.read_csv(io.StringIO(clipboard_data), sep='\t')

# Buscar la columna que contenga las referencias
columna_referencias = None
regex = r'^[a-zA-Z]{2}\d{6}[a-zA-Z]\d{3}$'
for columna in df.columns:
    if df[columna].apply(lambda x: bool(re.match(regex, x))).all():
        columna_referencias = columna
        break

if columna_referencias is None:
    print('No se encontraron referencias con el formato requerido.')
else:
    # Extraer las referencias que cumplan el formato requerido
    referencias = df[columna_referencias][df[columna_referencias].apply(lambda x: bool(re.match(regex, x)))].tolist()

    # Crear un nuevo DataFrame con las referencias
    df_referencias = pd.DataFrame({"Modelo": referencias})

    # Guardar el DataFrame en un archivo Excel en la carpeta Temp
    ruta_archivo = ''
    df_referencias.to_excel(ruta_archivo + "modelos.xlsx", index=False)

    # Abrir la carpeta donde se ha guardado el archivo
    os.startfile(os.path.dirname(ruta_archivo))
