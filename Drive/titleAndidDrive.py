#NioulBoy
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

g_login = GoogleAuth()
#g_login.LoadClientConfigFile("client_secrets.json")
# Try to load saved client credentials
g_login.LoadCredentialsFile("Drive\mycreds.txt")
if g_login.credentials is None:
  # Authenticate if they're not there
  g_login.LocalWebserverAuth()
elif g_login.access_token_expired:
  # Refresh them if expired
  g_login.Refresh()
else:
  # Initialize the saved creds
  g_login.Authorize()
# Save the current credentials to a file
g_login.SaveCredentialsFile("Drive\mycreds.txt")
drive = GoogleDrive(g_login)

# Create GoogleDrive instance with authenticated GoogleAuth instance.
drive = GoogleDrive(g_login)

fileList = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()
for file in fileList:
  print('Title: %s, ID: %s' % (file['title'], file['id']))



#Title: SCL Backup, ID: 1wet6NecLJduJ9FqiyKPd3PuIyefsNzqA