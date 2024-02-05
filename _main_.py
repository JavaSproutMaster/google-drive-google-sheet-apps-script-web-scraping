from googleapiclient.discovery import build
import csv
import io
import requests
import os
import app_script
import screenshot
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google_sheet_utils import set_column_width, set_wrap_text, write_and_highlight_values
import drive_utils
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.http import MediaIoBaseDownload


SCOPES = [
    "https://www.googleapis.com/auth/script.projects",
    'https://www.googleapis.com/auth/script.deployments',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive.scripts',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/script.triggers.readonly',
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/spreadsheets'
  ]
token_url = "token.json"
credential_url = 'app_script_credentials.json'

creds = None
if os.path.exists(token_url):
    creds = Credentials.from_authorized_user_file(token_url, SCOPES)
  # If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            credential_url, SCOPES
        )

        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(token_url, "w") as token:
        token.write(creds.to_json())

csv_files = []

# Create a Google Drive API service
service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)
apps_script_service = build('script', 'v1', credentials=creds)

def find_folder_id(service, folder_name):
    results = service.files().list(q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'").execute()
    items = results.get('files', [])
    if items:
        return items[0]['id']
    else:
        return None

def find_file_id(service, file_name, mimeType):
    results = service.files().list(q=f"name='{file_name}' and mimeType='{mimeType}'").execute()
    items = results.get('files', [])
    if items:
        return items[0]['id']
    else:
        return None
    
def find_file_by_id(file_id):
    try:
        # Get file metadata by file ID
        file_metadata = service.files().get(fileId=file_id).execute()
        return file_metadata
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
    
def get_files_in_a_folder(service, folder_id):
    results = service.files().list(q=f"'{folder_id}' in parents", pageSize=100, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found in the folder.')
    else:
        print('Files in the folder:')
        for item in items:
            print(f"{item['name']} ({item['id']})")
    return items

def get_folders_in_folder(folder_id):
    results = service.files().list(
        q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)").execute()
    
    folders = results.get('files', [])

    return folders

def find_sub_folderId_in_folder(parent_folder_id, folder_name):

    sub_folder_id = None
    folders = get_folders_in_folder(parent_folder_id)
    for folder in folders:
        if folder['name'] == folder_name:
            sub_folder_id = folder['id']
            break

    return sub_folder_id

def create_folder(parent_id, folder_name):
    exist_folder_id = find_folder_id(service, folder_name)
    folder_metadata = {
        'name': folder_name,
        'parents': [parent_id],
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = service.files().create(body=folder_metadata, fields='id').execute()
    return folder['id']

def move_file(service, file_id, destination_folder_id):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))

    # Remove the file from the source folder
    service.files().update(fileId=file_id,
                           addParents=destination_folder_id,
                           removeParents=previous_parents,
                           fields='id, parents').execute()

    print(f"Moved file ID {file_id} to folder ID {destination_folder_id}")

    return file_id
def check_and_delete_empty_folder(service, folder_id):
    try:
        # List files in the folder
        results = service.files().list(q=f"'{folder_id}' in parents", fields="files(id)").execute()
        files = results.get('files', [])

        if not files:
            # Folder is empty, delete it
            service.files().delete(fileId=folder_id).execute()
            print("Folder deleted successfully!")
        else:
            print("Folder is not empty.")
    except Exception as e:
        print(f"An error occurred: {e}")

def move_files_between_folders(service, src_folder_id, to_folder_id):
    response = service.files().list(q=f"'{src_folder_id}' in parents").execute()

    files = response.get('files', [])
    try:
        for file in files:
            # If the file is a folder, list its contents recursively
            if file['mimeType'] != 'application/vnd.google-apps.folder':
                move_file(service, file['id'], to_folder_id)
    
    except Exception as e:
        return False

    check_and_delete_empty_folder(service, src_folder_id)
    return True

def get_file_name(file_id):
    try:
        file_metadata = service.files().get(fileId=file_id).execute()
        return file_metadata['name']
    except Exception as e:
        print(f"Error getting file name: {e}")
        return None
    
def create_google_sheet_in_folder(parent_folder_id, sheet_title):
    file_metadata = {
        'name': sheet_title,
        'parents': [parent_folder_id],
        'mimeType': 'application/vnd.google-apps.spreadsheet',
    }
    res = service.files().create(body=file_metadata).execute()
    return res.get('id')


def rename_first_sheet(spreadsheet_id, new_sheet_name):
    try:
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet['sheets']

        if sheets:
            first_sheet_id = sheets[0]['properties']['sheetId']

            request_body = {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": first_sheet_id,
                                "title": new_sheet_name
                            },
                            "fields": "title"
                        }
                    }
                ]
            }

            response = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            ).execute()

            print(f"First sheet renamed to '{new_sheet_name}' successfully.")
        else:
            print("No sheets found in the spreadsheet.")
    except Exception as e:
        print(f"Error renaming sheet: {e}")

def rename_file(file_id, new_name):
    try:
        file_metadata = {'name': new_name}
        service.files().update(fileId=file_id, body=file_metadata).execute()
    except Exception as e:
        print(f"Error renaming file: {e}")

def downloadFromDrive(file_name, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = open(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    fh.close()

    print("Download complete.")

def create_sheet_and_import_csv(sheets_service, new_folder_id, csv_data, csv_title):
    sheet_id = create_google_sheet_in_folder(new_folder_id, csv_title)
     # Clear existing data in the sheet
    test_data=['test;this is test', 'test1;this is test2']
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range='Sheet1',  # Update with your sheet name or range
        body={}
    ).execute()

    # Update the sheet with the CSV data
    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range='Sheet1',  # Update with your sheet name or range
        valueInputOption='RAW',
        body={'values': csv_data}
    ).execute()
    
    rename_first_sheet(sheet_id, 'Data')

    return sheet_id
def list_csv_files(folder_id):
    csv_files = []

    response = service.files().list(q=f"'{folder_id}' in parents", pageSize=50, fields="nextPageToken, files(id, name, mimeType, parents)").execute()

    files = response.get('files', [])
    
    for file in files:
        if file['mimeType'] == 'text/csv':
            file_info = {
                'file_name': file['name'],
                'file_id': file['id'],
                'parents': file['parents']
            }
            csv_files.append(file_info)

        # If the file is a folder, list its contents recursively
        if file['mimeType'] == 'application/vnd.google-apps.folder':
            csv_files.extend(list_csv_files(file['id']))

    return csv_files

def list_sheet_files(folder_id):
    sheet_files = []

    response = service.files().list(q=f"'{folder_id}' in parents", pageSize=50, fields="nextPageToken, files(id, name, mimeType)").execute()
    
    files = response.get('files', [])
    
    for file in files:
        if file['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            file_info = {
                'file_name': file['name'],
                'file_id': file['id']
            }
            sheet_files.append(file_info)

        # If the file is a folder, list its contents recursively
        if file['mimeType'] == 'application/vnd.google-apps.folder':
            sheet_files.extend(list_sheet_files(file['id']))

    return sheet_files

def get_custom_val_from_csv_reader(reader, field_name, content):
    val = None
    try:
        for row in reader:
            if row.get('sep=;') and field_name == row.get('sep=;').split(';')[0]:
                val = row.get('sep=;').split(';')[1]
                return val
            if row.get('sep=,') and field_name == row.get('sep=,').split(';')[0]:
                val = row.get(None)[0]
                return val
    except Exception as e:
        csv_reader = csv.reader(content.splitlines())
        for row in csv_reader:
            if field_name in row:
                # Assuming the value is in the next row
                value_row = next(csv_reader)
                return value_row[0] if value_row else None
        return None

    # if val == None:
    #     csv_reader = csv.reader(content.splitlines())
    #     for row in csv_reader:
    #         if field_name in row:
    #             # Assuming the value is in the next row
    #             value_row = next(csv_reader)
    #             return value_row[0] if value_row else None
    return None

def get_csv_reader(csv_file_id):
    request = service.files().get_media(fileId=csv_file_id)
    content = request.execute()
    # Assuming the CSV has a header row, you may need to adjust the logic accordingly
    reader = csv.DictReader(io.StringIO(content.decode('utf-8')))
    
    return reader

def get_csv_content(csv_file_id):
    request = service.files().get_media(fileId=csv_file_id)
    content = request.execute()

    return content

def get_candidate_datas(file):
    reader = get_csv_reader(file.get('file_id'))
    content = get_csv_content(file.get('file_id'))
    agency_val = get_custom_val_from_csv_reader(reader, 'Full Name of Hiring Agency:', content)

    depart_name = ''
    if agency_val == None:
        return None

    item_id = 0
    for a_val in agency_val.split(' '):
        if a_val != '':
            if item_id == 0:
                depart_name = depart_name + a_val    
            else:
                depart_name = depart_name + a_val[0]
        item_id = item_id + 1

    can_name = get_custom_val_from_csv_reader(reader, 'Full Name',content)

    id_name = 0
    fname = ''
    can_name_list = can_name.split(' ')
    can_name_list.reverse()
    for name_snippet in can_name_list:
        if len(name_snippet) > 0:
            if id_name == 0:
                fname = fname + name_snippet
            else:
                fname = fname + name_snippet[0]
        id_name = id_name + 1

    position_val = get_custom_val_from_csv_reader(reader, 'Applying for what position?', content)
    
    return [str(fname) + ' ' + str(depart_name) + ' ' + str(position_val), can_name]

def signature_on_ROI(file):
    reader = get_csv_reader(file.get('file_id'))
    content = get_csv_content(file.get('file_id'))

    signatures = []

    signature = get_custom_val_from_csv_reader(reader, 'Signature of Client', content)
    if signature and signature != '':
        signatures.append(signature)

    if len(signatures) > 0 and signatures[len(signatures) - 1] == "Yes":
        return True
    else:
        return False
    
def is_vet_and_disability(file):
    reader = get_csv_reader(file.get('file_id'))
    content = get_csv_content(file.get('file_id'))

    is_vet = False
    has_disability = False
    
    if get_custom_val_from_csv_reader(reader, 'Did you serve in the military?', content) == 'Yes':
        is_vet = True
    if get_custom_val_from_csv_reader(reader, 'Do you have a Military Disability Rating?', content) == 'Yes':
        has_disability = True
    
    if is_vet and has_disability:
        return True
    
    return False

    
def get_script_id(script_name):
    try:
        # Search for the Apps Script by name
        query = f"name = '{script_name}' and mimeType = 'application/vnd.google-apps.script'"
        results = service.files().list(q=query).execute()

        # Get the script ID from the first result
        if 'files' in results and results['files']:
            script_id = results['files'][0]['id']
            print(f"Script ID for '{script_name}': {script_id}")
            return script_id
        else:
            print(f"No Apps Script found with the name '{script_name}'.")
            return None
    except Exception as e:
        print(f"Error getting script ID: {e}")
        return None

def run_apps_script(spreadsheet_id, script_property_key):
    try:
        # Retrieve the Script property value
        response = service.files().get(fileId=spreadsheet_id).execute()
        script_property_value = response.get('appProperties', {}).get(script_property_key)

        if script_property_value:
            print(f"Running Apps Script with property: {script_property_value}")
            # Implement the logic to run the script based on the property value
        else:
            print("No script property found.")
    except Exception as e:
        print(f"Error running script: {e}")

if __name__ == "__main__":

    agency_name = "mhoisingtonlmft"
    # Get folder Id
    background_folder_id = find_folder_id(service, "Backgrounds")

    agency_folder_id = find_sub_folderId_in_folder(background_folder_id, agency_name)

    mary_folder_id = find_folder_id(service, "Mary Reports New")
    mhoisingtonlmft_folder_id = find_folder_id(service, agency_name)
    invoice_folder_id = find_folder_id(service, "Invoices NEW - Mary Access Only")
    
    csv_files_list = list_csv_files(agency_folder_id)
    
    print("this is csv files_list")
    for csv_file in csv_files_list:
        file_id = csv_file.get('file_id')
        file_name = csv_file.get('file_name')
        downloadFromDrive(file_name, file_id)

    for csv_file in csv_files_list:
        file_id = csv_file.get('file_id')
        file_name = csv_file.get('file_name')
        downloadFromDrive(file_name, file_id)
        candidate_datas = get_candidate_datas(csv_file)
        
        new_can_folder_name = candidate_datas[0]
        
        candidate_name = candidate_datas[1]

        if new_can_folder_name == None:
            continue
        new_candidate_folder_id = create_folder(mary_folder_id, new_can_folder_name)

        signature_of_client = signature_on_ROI(csv_file)
        
        is_vet_and_have_disability_rating = is_vet_and_disability(csv_file)
        
        formdr_driver = screenshot.login()
        
        # upload signature of the client
        if signature_of_client:
            screenshot_media_body = screenshot.screenshot_signature(formdr_driver, candidate_name)
            if screenshot_media_body != None:
                file_metadata = {
                    'name': 'signature_screenshot.png',
                    'parents': [new_candidate_folder_id],
                }
                media = MediaIoBaseUpload(screenshot_media_body, mimetype='image/png')

                service = build('drive', 'v3', credentials=creds)
                service.files().create(body=file_metadata, media_body=media).execute()


                print("Screenshot uploaded successfully!")
        
        # upload va_background
        if is_vet_and_have_disability_rating:
            va_url = screenshot.get_va_download_url(formdr_driver, candidate_name)
            if va_url != None:
                va_pdf_response = requests.get(va_url)
                va_pdf_content = va_pdf_response.content
                service = build('drive', 'v3', credentials=creds)
                drive_utils.upload_pdf_content_into_drive(service, va_pdf_content, "VA_Background.pdf", new_candidate_folder_id)
        
        formdr_driver.quit()
        service = build('drive', 'v3', credentials=creds)

        move_file(service, file_id, new_candidate_folder_id)
        if len(csv_file['parents']) > 0:
            service = build('drive', 'v3', credentials=creds)

            move_files_between_folders(service, csv_file['parents'][0], new_candidate_folder_id)
        
        # bg_file_id = drive_utils.find_bg_file_in_backgrounds_folder(service, candidate_name)
        # if bg_file_id:
        #     move_file(service, bg_file_id, new_candidate_folder_id)

        csv_content = service.files().get_media(fileId=file_id).execute()
        reader = csv.reader(csv_content.decode('utf-8').splitlines())
        csv_data = []
    
        for d in csv_content.decode('utf-8').splitlines():
            csv_data.append(d.split(';'))
        file_name = get_file_name(file_id)
        if len(file_name.split('-')) > 2:
            csv_title = get_file_name(file_id).split('-')[1].split('.')[0]
        else:
            csv_title = file_name.split('.')[0]

        sheetId = create_sheet_and_import_csv(sheets_service, new_candidate_folder_id, csv_data , csv_title)
        set_column_width(sheets_service, sheetId, 0, 2, 220)
        write_and_highlight_values(sheets_service, sheetId, [['When stressed or upset'], ['Hobbies'], ['Supports']])
        set_wrap_text(sheets_service, sheetId, 0, 3)


        try:
            app_script.run(sheetId, new_candidate_folder_id, invoice_folder_id, new_can_folder_name)
        except Exception as e:
            print(e)
        
        print(new_can_folder_name)