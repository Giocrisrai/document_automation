import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

######################## CONFIGURACIÓN DEL USUARIO ########################

# path de salida
OUTPUT_PATH = "Outputs"

# path fichero Excel
EXCEL_PATH = "data/Input_de_Terreno_Biodiversidad_Aves_y_Mamíferos.xlsx"

# path de la plantilla
WORD_TEMPLATE_PATH = "data/template/Biodiversidad_Aves_y_Mamiferos.docx"

IMAGEN_PATH = "data/images/imagen.png"

######################## CONFIGURACIÓN DEL USUARIO ########################

# Rutina para eliminar y volver a crear la carpeta de salida


def EliminarCrearCarpetaSalida(path):
    if os.path.exists(path):
        shutil.rmtree(path)

    # Crear la carpeta de salida
    os.mkdir(path)

# Rutina para leer los datos del fichero Excel


def LeerDatosInforme(path, worksheet):
    # Leer el fichero Excel
    excel_df = pd.read_excel(path, sheet_name=worksheet)

    return excel_df

# Rutina para crear el informe Word


def CrearWordInforme(df_informe):
    # Cargar la plantilla
    docx_tpl = DocxTemplate(WORD_TEMPLATE_PATH)

    # Iterar sobre las filas del dataframe
    for index, r_val in df_informe.iterrows():
        # Añadir imagen gráfico circular y de barras
        # img_path = IMAGEN_PATH + '/' + r_val['Imagen']
        # img = InlineImage(docx_tpl, img_path, width=Mm(15))

        # Crear el diccionario con los datos
        context = r_val.to_dict()

        # Renderizar la plantilla usando el contexto
        docx_tpl.render(context)

        # Guardar el documento
        if (pd.notna(r_val['Documento'])):
            nombre_doc = 'Documento ' + \
                str(r_val['Documento']).upper() + '.docx'
        else:
            nombre_doc = 'Documento ' + str(index) + '.docx'

        docx_tpl.save(OUTPUT_PATH + '/' + nombre_doc)


def main():
    # Eliminar y volver a crear la carpeta Outputs
    EliminarCrearCarpetaSalida(OUTPUT_PATH)

    # Leer el fichero Excel
    df_informe = LeerDatosInforme(EXCEL_PATH, 'Data')

    # Crear informe word
    CrearWordInforme(df_informe)


if __name__ == '__main__':
    main()
