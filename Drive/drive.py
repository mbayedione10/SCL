#NioulBoy 08/2020
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import glob, os


def fileToDrive():
    gauth = GoogleAuth()
    # Try to load saved client credentials
    gauth.LoadCredentialsFile("Drive\mycreds.txt")
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("Drive\mycreds.txt")
    drive = GoogleDrive(gauth)


    list_of_files = glob.iglob("C:\\Users\\Mandiaye\\Documents\\Backup\\*.xlsx")  # * means all if need specific format then *.csv
    latest_file = sorted(list_of_files, key=os.path.getmtime, reverse=True)[:1]

    for file in latest_file:
        print(file)
        file_metadata = {'title': os.path.basename(file)}
        file_drive = drive.CreateFile(metadata=file_metadata)
        file_drive.SetContentFile(file)
        file_drive.Upload()

        print("The file: " + file + " has been uploaded")

    print("All files have been uploaded")


fileToDrive()