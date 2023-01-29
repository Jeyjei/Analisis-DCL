import sys
import pandas as pd
import matplotlib.pyplot as plt
import os
import xlrd
import pandas as pd
import openpyxl
import numpy as np


# =================================== FUNCIONES ======================================

# Carga los archivos xls. y .xlsx
def load_Excel_books(ruta_archivo):

    # Obtenemos la extensión del archivo
    extension = os.path.splitext(ruta_archivo)[1]

    # Comparamos la extensión con las cadenas de caracteres 'xls' y 'xlsx'
    if extension == '.xls':
        # Es un archivo de Excel en el formato .xls
        
        # Abrimos el archivo de Excel
        workbook = xlrd.open_workbook(ruta_archivo)

        # Obtenemos el nombre de las hojas del archivo de Excel
        sheet_names = workbook.sheet_names()

        # Leemos cada hoja del archivo de Excel y guardamos los datos en una lista de data frames
        sheets = []
        name_sheets = []

        for sheet_name in sheet_names:
            if not sheet_name.startswith(("Hoja3", "PORTADA", "INTRODUCCIÓN")):
                
                # Guardamos el nombre de la hoja
                name_sheets.append(sheet_name)

                # Abrimos la hoja del archivo de Excel
                sheet = workbook.sheet_by_name(sheet_name)

                # Obtenemos los datos de la hoja en una lista de listas
                data = []
                for row_index in range(sheet.nrows):
                    row = sheet.row_values(row_index)
                    data.append(row)

                # Creamos un data frame a partir de los datos de la hoja
                df = pd.DataFrame(data)

                # Añadimos el data frame a la lista de data frames
                sheets.append(df)


    elif extension == '.xlsx':
        # Es un archivo de Excel en el formato .xlsx
        
        # Abrimos el archivo de Excel
        workbook = openpyxl.load_workbook(ruta_archivo)

        # Obtenemos el nombre de las hojas del archivo de Excel
        sheet_names = workbook.sheetnames

        # Leemos cada hoja del archivo de Excel y guardamos los datos en una lista de data frames
        sheets = []
        name_sheets = []

        for sheet_name in sheet_names:
            if not sheet_name.startswith(("Hoja3", "PORTADA", "INTRODUCCIÓN")):

                # Guardamos el nombre de la hoja
                name_sheets.append(sheet_name)

                # Abrimos la hoja del archivo de Excel
                sheet = workbook[sheet_name]

                # Obtenemos los datos de la hoja en una lista de listas
                data = []
                for row in sheet.rows:
                    data.append([cell.value for cell in row])

                # Creamos un data frame a partir de los datos de la hoja
                df = pd.DataFrame(data)

                # Añadimos el data frame a la lista de data frames
                sheets.append(df)

    else:
        # No es un archivo de Excel
        print("El archivo no es un archivo de Excel")


    return(sheets, name_sheets)


# Genera un dataset con las variables interés del archivo
def Get_Data(fichero, names_sheet_fichero, var_interes):

    # Creamos un dataframe para cada año/fichero
    df = pd.DataFrame()

    for ind_sheet in range(len(fichero)):

        i0_sheet = []
        # Listas para guardar las variables
        indicadores = []
        ciudad = []
        distrito = []
        i = 0

        # Recorremos todas las variables del dataset
        for i_0 in range(len(fichero[ind_sheet][0])): 

            if fichero[ind_sheet][0][i_0] is not None:
                # Si dichas variables empiezan por las variables de interés que tenemos en "var_interes"
                if fichero[ind_sheet][0][i_0].strip().startswith(var_interes): #.strip() elimina los espacios en blanco al principio y final de una cadena

                    # Valores de cada distrito
                    if i < 2:
                        if isinstance(fichero[ind_sheet].iloc[i_0, 3], str):
                            distrito.append( float(fichero[ind_sheet].iloc[i_0, 3].replace(".","").replace(",", ".")) )
                        else:
                            distrito.append( fichero[ind_sheet].iloc[i_0, 3] )
                    else:
                        if isinstance(fichero[ind_sheet].iloc[i_0, 4], str):
                            distrito.append( float(fichero[ind_sheet].iloc[i_0, 4].replace(".","").replace(",", ".")) )
                        else:
                            distrito.append( fichero[ind_sheet].iloc[i_0, 4] )
                    i = i + 1

                    # Solo necesitamos guardar estos valores una única vez
                    if ind_sheet == 0:
                        # Guardamos los valores de las variables
                        indicadores.append(fichero[ind_sheet][0][i_0].strip())
                        if fichero[ind_sheet].iloc[i_0, 2] is not None:
                            if isinstance(fichero[ind_sheet].iloc[i_0, 2], str):
                                ciudad.append( float(fichero[ind_sheet].iloc[i_0, 2].replace(".","").replace(",", ".")) )
                            else:
                                ciudad.append( fichero[ind_sheet].iloc[i_0, 2] )
                        else:
                            if isinstance(fichero[ind_sheet].iloc[i_0, 1], str):
                                ciudad.append( float(fichero[ind_sheet].iloc[i_0, 1].replace(".","").replace(",", ".")) )
                            else:
                                ciudad.append( fichero[ind_sheet].iloc[i_0, 1] )

                            

                    # Para la comprobación de errores
                    i0_sheet.append(True)
                else:
                    # Para la comprobación de errores
                    i0_sheet.append(False)

        if ind_sheet == 0:
            # Para la comprobación de errores
            # print(fichero[ind_sheet].iloc[i0_sheet, [0, 2, 3, 4]])

            # Guardamos los datos
            df["Indicadores"] = indicadores
            df["Cuidad"] = ciudad
            df[names_sheet_fichero[ind_sheet]] = distrito

            
        else:
            # Para la comprobación de errores
            # print(fichero[ind_sheet].iloc[i0_sheet, [0, 2, 3, 4]]) 
        
            # Guardamos los datos
            df[names_sheet_fichero[ind_sheet]] = distrito 

    return(df)



# ==========================================================================================
# ===================================== MAIN ===============================================

def main(file1, file2, file3, file4, file5):
    
    # Cargamos los archivos
    try:
        # ------------- Datos del Ayuntamiento de Madrid -------------
        # Año 2016 
        (xls_2016, name_sheet_2016) = load_Excel_books(file1)

        # Año 2017
        (xls_2017, name_sheet_2017) = load_Excel_books(file2)

        # Año 2018
        (xls_2018, name_sheet_2018) = load_Excel_books(file3)

        # Año 2019
        (xls_2019, name_sheet_2019) = load_Excel_books(file4)

        # ------------- Datos INE -------------
        datos_INE = pd.read_csv(file5, 
                        sep = ';', 
                        index_col = 0, 
                        decimal = ',', 
                        thousands = '.')
        # Obtenemos los nombres únicos de los distritos, indicadores y períodos
        name_distritos = datos_INE['Distritos'].unique()
        name_indicadores = datos_INE['Indicadores de renta media y mediana'].unique()
        name_periodo = datos_INE['Periodo'].unique()

    except:
        print('Ha habido un problema a la hora de cargar los archivos')


    # ------------- Variables de interés del Dataset del Ayuntamiento de Madrid -------------
    var_AYU_interes = ('Superficie (Ha.)', 
                    'Densidad (hab./Ha.)',
                    'No sabe leer ni escribir', 
                    'Bachiller Elemental', 
                    'Formación profesional', 
                    'Titula',
                    'Estudios superiores',
                    'Nivel de estudios',
                    'Escuelas Infantiles Municipales',
                    'Escuelas Infantiles Públicas CAM',
                    'Escuelas Infantiles Privadas',
                    'Colegios Públicos Infantil y Primaria',
                    'Institutos Públicos de Educación Secundaria',
                    'Colegios Privados Inf. o Pri. o Inf. y Pri.',
                    'Colegios Privados Inf. o Pri. o Inf. y Pri.')

    # Guardamos los datos por año en una DataFrame
    data_2016 = Get_Data(xls_2016, name_sheet_2016, var_AYU_interes)
    data_2017 = Get_Data(xls_2017, name_sheet_2017, var_AYU_interes)
    data_2018 = Get_Data(xls_2018, name_sheet_2018, var_AYU_interes)
    data_2019 = Get_Data(xls_2019, name_sheet_2019, var_AYU_interes)


    # ---------------------------- Modificamos el Dataset del INE ----------------------------
    for i in range(0, len(name_sheet_2016)):
        datos_INE.loc[datos_INE['Distritos'] == name_distritos[i+1], 'Distritos'] = name_sheet_2016[i]

    # Indicador de interés en el dataset del INE
    var_INE_interes = name_indicadores[0]
    # Nos quedamos únicamente con el indicador de interés
    renta_neta = datos_INE[(datos_INE['Indicadores de renta media y mediana'] == var_INE_interes)]

    # Elimino la columna 'Secciones' que son NaN
    renta_neta = renta_neta.drop(columns=["Secciones"])

    # Divido la columan 'Periodo'
    df_pivoted = renta_neta.pivot_table(index=["Distritos", "Indicadores de renta media y mediana"], 
                                            columns="Periodo", values="Total", aggfunc="first")
    # Reajustamos los índices
    df_pivoted.columns = list(df_pivoted.columns)
    df_pivoted = df_pivoted.reset_index()
    df_pivoted.columns

    # Año 2016
    renta_neta_2016 = datos_INE[(datos_INE['Periodo'] == 2016) & 
                                (datos_INE['Indicadores de renta media y mediana'] == var_INE_interes)]

    # Año 2017
    renta_neta_2017 = datos_INE[(datos_INE['Periodo'] == 2017) & 
                                (datos_INE['Indicadores de renta media y mediana'] == var_INE_interes)]

    # Año 2018
    renta_neta_2018 = datos_INE[(datos_INE['Periodo'] == 2018) & 
                                (datos_INE['Indicadores de renta media y mediana'] == var_INE_interes)]

    # Año 2019
    renta_neta_2019 = datos_INE[(datos_INE['Periodo'] == 2019) & 
                                (datos_INE['Indicadores de renta media y mediana'] == var_INE_interes)]


    # ---------------------------- Unificamos los datos ----------------------------
    # Quitamos los espacios en blanco porque queremos porteriormente ordendar el DataFrame segúne este indicador
    var_INE_interes = var_INE_interes.strip()

    # Año 2016
    values_renta_16 = renta_neta_2016['Total'].tolist()
    values_renta_16.insert(0, var_INE_interes)      # Añadimos el nombre del indicador
    last_index = data_2016.index.max() + 1          # Obtenemos el último índice del DataFrame
    data_2016.loc[last_index] = values_renta_16

    # Año 2017
    values_renta_17 = renta_neta_2017['Total'].tolist()
    values_renta_17.insert(0, var_INE_interes)      # Añadimos el nombre del indicador
    last_index = data_2017.index.max() + 1          # Obtenemos el último índice del DataFrame
    data_2017.loc[last_index] = values_renta_17

    # Año 2018
    values_renta_18 = renta_neta_2018['Total'].tolist()
    values_renta_18.insert(0, var_INE_interes)      # Añadimos el nombre del indicador
    last_index = data_2018.index.max() + 1          # Obtenemos el último índice del DataFrame
    data_2018.loc[last_index] = values_renta_18

    # Año 2019
    values_renta_19 = renta_neta_2019['Total'].tolist()
    values_renta_19.insert(0, var_INE_interes)      # Añadimos el nombre del indicador
    last_index = data_2019.index.max() + 1          # Obtenemos el último índice del DataFrame
    data_2019.loc[last_index] = values_renta_19

    # Crear un diccionario con los dataframes y los nombres de las hojas
    dataframes = {'Datos_2016': data_2016, 
                'Datos_2017': data_2017,
                'Datos_2018': data_2018, 
                'Datos_2019': data_2019}

    try:
        # Crear un archivo excel con varias hojas y escribir los dataframes en ellas
        with pd.ExcelWriter('Data_final.xlsx') as writer:
            for sheet_name, df in dataframes.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except:
        print('Hubo un error a la hora de guardar el archivo final')


if __name__ == "__main__":
    print('Number of arguments: %i arguments' % len(sys.argv))
    print('Argument List:' + str(sys.argv))

    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
