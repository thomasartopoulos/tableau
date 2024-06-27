### construido con documentación oficial de tableau, desarrollado para tableau cloud
### thomas artopoulos, 2024

import tableauserverclient as TSC
import os
import time
import pandas as pd

# reemplaza 'Nomenclador.xlsx' con la ruta a tu archivo de excel
## la idea es armar un diccionario para trabajar con dos filtros a la vez
file_path = 'Nomenclador.xlsx'
df = pd.read_excel(file_path, engine='openpyxl')

df = df[['NOMBRE', 'ORGANIZACIÓN ']]

# corrigiendo el posible espacio al final en el nombre de la columna 'ORGANIZACIÓN '
df.columns = df.columns.str.strip()

# construyendo el diccionario con 'NOMBRE' como claves y 'ORGANIZACIÓN' como valores
tiendas = df.set_index('NOMBRE')['ORGANIZACIÓN'].to_dict()

# url del servidor de tableau y credenciales corregidas
SERVER_URL = "url_servidor"
USERNAME = "email"
PERSONAL_ACCESS_TOKEN_NAME = "your_token_name"
PERSONAL_ACCESS_TOKEN_SECRET = os.getenv("TABLEAU_PERSONAL_ACCESS_TOKEN")  # usar variables de entorno
SITE = "sitio"  # establece tu id de sitio/content url aquí

# inicializar objeto servidor
server = TSC.Server(SERVER_URL, use_server_version=True)

# usar token de acceso personal para autenticación
auth = TSC.PersonalAccessTokenAuth(PERSONAL_ACCESS_TOKEN_NAME, PERSONAL_ACCESS_TOKEN_SECRET, SITE)
# información del libro de trabajo y vista
WORKBOOK_NAME = "nombre_workbook"
VIEW_NAME = "nombre_vista"

# configuración de exportación de pdf
EXPORT_FOLDER_LOCATION = os.getcwd() + "/Exported_Views/"
EXPORT_FILE_EXTENSION = ".pdf"

# iniciar sesión en el servidor
with server.auth.sign_in(auth):
    print('conectado al servidor de tableau exitosamente.')

    # encontrar el libro de trabajo por nombre
    req_option = TSC.RequestOptions()
    req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name,
                                     TSC.RequestOptions.Operator.Equals, WORKBOOK_NAME))
    all_workbooks, _ = server.workbooks.get(req_option)

    if not all_workbooks:
        print(f'no se encontró ningún libro de trabajo llamado {WORKBOOK_NAME}.')
        exit()

    workbook = all_workbooks[0]
    server.workbooks.populate_views(workbook)

    # asegurar que el directorio de exportación exista
    if not os.path.exists(EXPORT_FOLDER_LOCATION):
        os.makedirs(EXPORT_FOLDER_LOCATION)
        print("el directorio de exportación fue creado.")
    
    # recorrer cada tienda en el diccionario 'tiendas'
    for tienda, formato_vistas in tiendas.items():
        # encontrar la vista específica en el libro de trabajo
        all_views, pagination_item = server.views.get(req_option)
        view_item = next((view for view in all_views if view.name == VIEW_NAME), None)
        if view_item is None:
            print(f"vista llamada '{VIEW_NAME}' no encontrada.")
            continue
    
        # especificar los parámetros de exportación, incluyendo el filtro
        pdf_req_option = TSC.PDFRequestOptions(page_type=TSC.PDFRequestOptions.PageType.A4,
                                               orientation=TSC.PDFRequestOptions.Orientation.Landscape,
                                               maxage=1)
    
        # aplicar valores 'formato vistas' y 'tienda 2' del diccionario
        pdf_req_option.vf('Formato Vistas', formato_vistas)  
        pdf_req_option.vf('Tienda', tienda)  # ajustar la aplicación del filtro según sea necesario
        #pdf_req_option.vf('Periodo', '2024-04')  
    
        # exportar la vista como pdf
        pdf_path = EXPORT_FOLDER_LOCATION + tienda.replace(" - ", "_") + EXPORT_FILE_EXTENSION
        server.views.populate_pdf(view_item, pdf_req_option)

        time.sleep(10)

        with open(pdf_path, 'wb') as file:
            file.write(view_item.pdf)
        print(f"vista de {tienda} exportada a pdf exitosamente.")
