from pydrive.auth import GoogleAuth 
from pydrive.drive import GoogleDrive

gauth = GoogleAuth() 
drive = GoogleDrive(gauth)

def upload2drive(file_path, file_name, folder_id):
    gfile = drive.CreateFile({'parents': [{'id': folder_id}]})
    
    gfile['title'] = file_name
    gfile.SetContentFile(file_path)
    
    gfile.Upload()
