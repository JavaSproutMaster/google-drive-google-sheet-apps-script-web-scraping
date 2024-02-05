from googleapiclient.http import MediaIoBaseUpload
import io
import _main_
from googleapiclient.http import MediaIoBaseDownload
import PyPDF2
import re

def list_files_in_a_folder_with_mimeType(drive_service, folder_id, mimetype):
    sheet_files = []

    response = drive_service.files().list(q=f"'{folder_id}' in parents", pageSize=50, fields="nextPageToken, files(id, name, mimeType)").execute()
    
    files = response.get('files', [])
    
    for file in files:
        if file['mimeType'] == mimetype:
            file_info = {
                'file_name': file['name'],
                'file_id': file['id']
            }
            sheet_files.append(file_info)

        # If the file is a folder, list its contents recursively
        if file['mimeType'] == 'application/vnd.google-apps.folder':
            sheet_files.extend(list_files_in_a_folder_with_mimeType(drive_service, file['id'], mimetype))

    return sheet_files

def find_pdf_by_name(service, folder_id, start_name):
    # Call the Drive v3 API to search for PDF files in the specified folder
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and name contains '{start_name}'"
    results = service.files().list(q=query, pageSize=10, fields="files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print(f'No PDF files found in the folder with the name starting with "{start_name}".')
    else:
        print(f'PDF files in the folder with the name starting with "{start_name}":')
        for item in items:
            print(f"{item['name']} ({item['id']})")

    return items

def upload_pdf_content_into_drive(drive_service, pdf_content, pdf_name, folder_id):
    # Create metadata for the file
    file_metadata = {
        'name': pdf_name,  # You can specify the desired file name
        'parents': [folder_id],
    }
    media = MediaIoBaseUpload(io.BytesIO(pdf_content), mimetype='application/pdf')
    drive_service.files().create(body=file_metadata, media_body=media).execute()

def extract_name_from_bg_file(drive_service, file_id):
    test_content = download_pdf_from_drive(drive_service, file_id)

    full_name = extract_name_from_bg_pdf(test_content)

    names = full_name.split()

    cand_name = ''
    if len(names) > 2:
        first_name = names[0]
        last_name = names[-1]
        cand_name = first_name + ' ' + last_name
        print("First Name:", first_name)
        print("Last Name:", last_name)
        print(cand_name)
    if len(names) == 2:
        cand_name = full_name
    
    return cand_name

def find_bg_file_in_backgrounds_folder(drive_service, candidate_name):
    backgrounds_folder_id = _main_.find_folder_id(drive_service, "Backgrounds")
    pdf_files_in_backgrounds = list_files_in_a_folder_with_mimeType(drive_service, backgrounds_folder_id, "application/pdf")
    file_id = None
    for pdf_file in pdf_files_in_backgrounds:
        file_name = pdf_file.get('file_name')
        if file_name == None:
            continue

        if file_name.endswith('BG-report.pdf'):
            pdf_file_id = pdf_file.get('file_id')
            if candidate_name == extract_name_from_bg_file(drive_service, pdf_file_id):
                file_id = pdf_file_id
                break
    
    return file_id
    
# def extract_text_from_pdf(drive_service, file_id):
#     request = drive_service.files().export_media(fileId=file_id, mimeType='text/plain')

#     text_content = io.BytesIO(request.execute())

#     text_content.seek(0)
#     return text_content.read().decode('utf-8')

def download_pdf_from_drive(drive_service, file_id):
    request = drive_service.files().get_media(fileId=file_id)
    pdf_content = io.BytesIO()
    downloader = MediaIoBaseDownload(pdf_content, request)
    
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    
    pdf_content.seek(0)
    return pdf_content

def extract_name_from_bg_pdf(pdf_content):
    pdf_reader = PyPDF2.PdfReader(pdf_content)
    text = ''
    page_number = 0
    if len(pdf_reader.pages) > 0:
        text += pdf_reader.pages[page_number].extract_text()
    
    match = re.search(r'APPLICANT\n(.*?)\nINVESTIGATION COMPLETED BY', text, re.DOTALL)

    full_name = ''
    if match:
        full_name = match.group(1).strip()
        print(full_name)
        return full_name
    else:
        print("Pattern not found in the text.")

        return ''