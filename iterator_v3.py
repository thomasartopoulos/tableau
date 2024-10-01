import os
import pandas as pd
import tableauserverclient as TSC
import re

file_path = 'NOMENCLADOR PATH'

# Load the Excel file into a DataFrame
df = pd.read_csv(file_path)

# Selecting necessary columns and correcting column names
df = df[['NOMBRE', 'ORGANIZACIÓN ']]
df.columns = df.columns.str.strip()

# Building the dictionary with 'NOMBRE' as keys and 'ORGANIZACIÓN' as values
tiendas = df.set_index('NOMBRE')['ORGANIZACIÓN'].to_dict()

# Tableau Server credentials and settings
SERVER_URL = ""
USERNAME = ""
PERSONAL_ACCESS_TOKEN_NAME = ""
PERSONAL_ACCESS_TOKEN_SECRET = ""
SITE = ""  # Set your site ID/content URL here

# Initialize Server object
server = TSC.Server(SERVER_URL, use_server_version=True)

# Use Personal Access Token for authentication
auth = TSC.PersonalAccessTokenAuth(PERSONAL_ACCESS_TOKEN_NAME, PERSONAL_ACCESS_TOKEN_SECRET, SITE)
WORKBOOK_NAME = "Ficha Tienda"
VIEW_NAME = "Ficha Tienda"

# PDF Export Settings
EXPORT_FOLDER_LOCATION = os.path.join(os.getcwd(), "Exported_Views")
EXPORT_FILE_EXTENSION = ".pdf"
LAST_PROCESSED_FILE = 'last_processed_store.txt'

# Ensure the export directory exists
os.makedirs(EXPORT_FOLDER_LOCATION, exist_ok=True)

# Function to read the last processed store from a file
def read_last_processed():
    if os.path.exists(LAST_PROCESSED_FILE):
        with open(LAST_PROCESSED_FILE, 'r') as file:
            return file.read().strip()
    return None

# Function to write the last processed store to a file
def write_last_processed(store):
    with open(LAST_PROCESSED_FILE, 'w') as file:
        file.write(store)

# Ensure the path length is within Windows limits
def ensure_path_length(folder, filename, max_length=255):
    full_path = os.path.join(folder, filename)
    if len(full_path) > max_length:
        raise ValueError(f"Path length exceeds {max_length} characters: {full_path}")
    return full_path

# Function to sanitize file names
def sanitize_filename(filename):
    # Remove or replace invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return filename

# Get the last processed store
last_processed = read_last_processed()
start_processing = last_processed is None

# Sign in to the server
with server.auth.sign_in(auth):
    print('Logged in to the Tableau Server successfully.')

    # Find the workbook by name
    req_option = TSC.RequestOptions()
    req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name, TSC.RequestOptions.Operator.Equals, WORKBOOK_NAME))
    all_workbooks, _ = server.workbooks.get(req_option)

    if not all_workbooks:
        print(f'No workbook named {WORKBOOK_NAME} found.')
        exit()

    workbook = all_workbooks[0]
    server.workbooks.populate_views(workbook)

    # Loop through each store in 'tiendas' dictionary
    for tienda, organizacion in tiendas.items():
        if start_processing or tienda == last_processed:
            start_processing = True
        else:
            continue

        # Find the specific view in the workbook
        all_views, _ = server.views.get(req_option)
        view_item = next((view for view in all_views if view.name == VIEW_NAME), None)
        if view_item is None:
            print(f"View named '{VIEW_NAME}' not found.")
            continue

        # Specify the export parameters, including the filter
        pdf_req_option = TSC.PDFRequestOptions(page_type=TSC.PDFRequestOptions.PageType.A4,
                                               orientation=TSC.PDFRequestOptions.Orientation.Landscape,
                                               maxage=1)

        # Apply 'Formato Vistas' and 'Tienda 2' values from the dictionary
        pdf_req_option.vf('Formato Vistas', organizacion)
        pdf_req_option.vf('Tienda 2', tienda)

        # Create a folder for the ORGANIZACIÓN if it doesn't exist
        organizacion_folder = os.path.join(EXPORT_FOLDER_LOCATION, organizacion)
        os.makedirs(organizacion_folder, exist_ok=True)

        # Sanitize the file name
        sanitized_tienda = sanitize_filename(f"{tienda} (07-2024){EXPORT_FILE_EXTENSION}")

        # Ensure the path length is within limits
        pdf_path = ensure_path_length(organizacion_folder, sanitized_tienda)

        # Export the view as PDF
        server.views.populate_pdf(view_item, pdf_req_option)

        with open(pdf_path, 'wb') as file:
            file.write(view_item.pdf)
        print(f"Exported {tienda} view to PDF successfully in folder: {organizacion_folder}")

        # Update the last processed store
        write_last_processed(tienda)
