import os
import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt

# FOLDER_PATH = r"path_to_save_folder"  # e.g. r"C:\Users\Name\Downloads"
FOLDER_PATH = r"temp"  # e.g. r"C:\Users\Name\Downloads"

# Import CSV file using one of the two methods below
# df = pd.read_csv(r"path_to_csv\fruit_and_veg_sales.csv")
df = pd.read_csv(r"fruit_and_veg_sales.csv")

wb = xw.Book()
sht = wb.sheets["Sheet1"]
sht.name = "fruit_and_veg_sales"
sht.range("A1").options(index=False).value = df
wb.sheets.add('Dashboard')
sht_dashboard = wb.sheets('Dashboard')

# ===========================# MANIPULACION DE DATA # ==========================

# Verifique los nombres de las columnas
print(df.columns)

# Girar por artículo y mostrar el beneficio total
pv_total_profit = pd.pivot_table(df, index='Item', values='Total Profit ($)', aggfunc='sum')

# Pivote los datos por artículo y muestre la cantidad vendida
pv_quantity_sold = pd.pivot_table(df, index='Item', values='Quantity Sold', aggfunc='sum')

# Tipo de datos de fecha correcta de venta
print(df.dtypes)
df["Date Sold"] = pd.to_datetime(df["Date Sold"], format='%d/%m/%Y')

# Agrupar por fecha de venta en meses
gb_date_sold = df.groupby(df["Date Sold"].dt.to_period('m')).sum()[["Quantity Sold", 'Total Revenue ($)',
                                                                    'Total Cost ($)', "Total Profit ($)"]]
gb_date_sold.index = gb_date_sold.index.to_series().astype(str)

# Agrupar por fecha de venta, ordenar por ingresos totales, mostrar las 8 filas principales
gb_top_revenue = (df.groupby(df["Date Sold"])
                  .sum()
                  .sort_values('Total Revenue ($)', ascending=False)
                  .head(8)
                  )[["Quantity Sold", 'Total Revenue ($)',
                     'Total Cost ($)', "Total Profit ($)"]]

#  ============================ # Dar formato al panel # =========================

# Background
sht_dashboard.range('A1:Z1000').color = (198, 224, 180)

# A:B Ancho de columna
sht_dashboard.range('A:B').column_width = 2.22

# Titulo
sht_dashboard.range('B2').value = 'Sales Dashboard'
sht_dashboard.range('B2').api.font.Name = 'Arial'
sht_dashboard.range('B2').api.font.Size = 48
sht_dashboard.range('B2').api.font.Bold = True
sht_dashboard.range('B2').api.font.Color = 0x000000
sht_dashboard.range('B2').row_height = 61.2

# Título subrayado
data_range = sht_dashboard.range('B2:W2')
sht_dashboard.range('B2:W2').api.borders(9).Weight = 4
sht_dashboard.range('B2:W2').api.borders(9).Color = 0x00B050

# Sub-titulo
sht_dashboard.range('M2').value = 'Total Profit Per Item Chart'
sht_dashboard.range('M2').api.font.Name = 'Arial'
sht_dashboard.range('M2').api.font.Size = 20
sht_dashboard.range('M2').api.font.Bold = True
sht_dashboard.range('M2').api.font.Color = 0x000000

# Linea divisora entre titulo y sub-titulo
sht_dashboard.range('L2').api.borders(7).Weight = 3
sht_dashboard.range('L2').api.borders(7).Color = 0x00B050
sht_dashboard.range('L2').api.borders(7).LineStyle = -4115


# Función que formatea un marco de datos.
def create_formatted_summary(header_cell, title, df_summary, color):
    """
    Parámetros
    ----------
    header_cell: Str
        Proporcione la ubicación de la celda superior izquierda donde desea colocar un marco de datos. p.ej. 'B2'

    tittle: Str
        Especifica qué título quieres que tenga este bloque. p.ej. 'Título dinámico'

    df_summary: DataFrame
        Proporcione el DataFrame que desea colocar en Excel.

    color: Str
        Proporcione el nombre de un color. p.ej. 'azul', etc.
        Consulte la función de diccionario de colores (colores).
        Se pueden agregar más colores a este diccionario,
        simplemente agregue tuplas RGB de un tono más claro y más oscuro del mismo color

    returns
    -------
    Ninguno. Esta función solo formatea Excel.
    """

    # Diccionario de colores, [(color más oscuro), (color más claro)]
    colors = {"purple": [(112, 48, 160), (161, 98, 208)],
              "blue": [(0, 112, 192), (155, 194, 230)],
              "green": [(0, 176, 80), (169, 208, 142)],
              "yellow": [(255, 192, 0), (255, 217, 102)]}

    # Establecer el ancho de columna de la primera columna de resumen
    sht_dashboard.range(header_cell).column_width = 1.5

    # Asignar ubicación de fila y columna a variables
    row, col = sht_dashboard.range(header_cell).row, sht_dashboard.range(header_cell).column

    # Título del formato del resumen de DataFrame
    summary_title_range = sht_dashboard.range(header_cell, header_cell)
    summary_title_range.value = title
    summary_title_range.api.font.Size = 14
    summary_title_range.row_height = 32.5
    summary_title_range.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    summary_title_range.api.font.Color = 0xFFFFFF
    summary_title_range.api.font.Bold = True
    sht_dashboard.range((row, col),
                        (row, col + len(df_summary.columns) + 1)).color = colors[color][0]  # Darker color

    # Dar formato a los encabezados del resumen de DataFrame
    summary_header_range = sht_dashboard.cells(row + 1, col + 1)
    summary_header_range.value = df_summary
    summary_header_range = summary_header_range.expand('right')
    summary_header_range.api.font.Size = 11
    summary_header_range.api.font.Bold = True
    sht_dashboard.range((row + 1, col),
                        (row + 1, col + len(df_summary.columns) + 1)).color = colors[color][1]  # Darker color
    sht_dashboard.range((row + 1, col + 1),
                        (row + len(df_summary), col + len(df_summary.columns) + 1)).autofit()

    for num in range(1, len(df_summary) + 2, 2):
        sht_dashboard.range((row + num, col),
                            (row + num, col + len(df_summary.columns) + 1)).color = colors[color][1]

    # Encuentra la última fila del resumen de DataFrame
    last_row = sht_dashboard.cells(row + 1, col + 1).expand('down').last_cell.row
    side_border_range = sht_dashboard.range((row + 1, col), (last_row, col))

    # Agregue un borde punteado con color a la izquierda del resumen de DataFrame
    sht_dashboard.range(side_border_range).api.borders(7).Weight = 3
    sht_dashboard.range(side_border_range).api.borders(7).Color = xw.utils.rgb_to_int(colors[color][1])
    sht_dashboard.range(side_border_range).api.borders(7).LineStyle = -4115


# Ejecuta la función y crea cada sección de nuestra hoja de resumen.
create_formatted_summary('B5', 'Total Profit per Item', pv_total_profit, 'green')
create_formatted_summary('B17', 'Total Items Sold', pv_quantity_sold, 'purple')
create_formatted_summary('F17', 'Sales by Month', gb_date_sold, 'blue')
create_formatted_summary('F5', 'Top 5 Days by Revenue ', gb_top_revenue, 'yellow')

# Hace un gráfico usando Matplotlib
fig, ax = plt.subplots(figsize=(6, 3))
pv_total_profit.plot(color='g', kind='bar', ax=ax)

# Agregar gráfico a la hoja del tablero
sht_dashboard.pictures.add(fig, name='ItemsChart',
                           left=sht_dashboard.range("M5").left,
                           top=sht_dashboard.range("M5").top,
                           update=True)

# =====================# BONUS - AÑADIR LOGO A SU TABLERO #===================

image_url = r"pie_logo"

# image_path = rf"{FOLDER_PATH}\pie_logo.png"
current_path = os.getcwd()
image_path = f'{current_path}/pie_logo.png'

# Agrega una imagen al tablero de Excel
logo = sht_dashboard.pictures.add(image=image_path,
                                  name='PC_3',
                                  left=sht_dashboard.range("J2").left,
                                  top=sht_dashboard.range("J2").top + 5,
                                  update=True)

# Cambia el tamaño de la imagen
logo.width = 54
logo.height = 54

# ==============================================================================

# Guarde su archivo de Excel
wb.save(f"{current_path}/dashboard.xlsx")
