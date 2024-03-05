# from google.oauth2.credentials import Credentials
# from google_auth_oauthlib.flow import InstalledAppFlow
# from google.auth.transport.requests import Request
# from googleapiclient.discovery import build
# from Converter import Txt_to_xlsx


# class GoogleDriveHelper:
#     def __init__(self, creds_file):
#         self.creds_file = creds_file
#         self.SCOPES = ['https://www.googleapis.com/auth/drive']
#         self.service = self.authenticate()

#     def authenticate(self):
#         creds = None
#         if self.creds_file.exists():
#             creds = Credentials.from_authorized_user_file(self.creds_file)
#         if not creds or not creds.valid:
#             if creds and creds.expired and creds.refresh_token:
#                 creds.refresh(Request())
#             else:
#                 flow = InstalledAppFlow.from_client_secrets_file(
#                     'credentials.json', self.SCOPES)
#                 creds = flow.run_local_server(port=0)
#             with open(self.creds_file, 'w') as token:
#                 token.write(creds.to_json())
#         return build('drive', 'v3', credentials=creds)

#     def get_file_id(self, file_name):
#         query = f"name='{file_name}'"
#         results = self.service.files().list(q=query,
#                                             fields="files(id)").execute()
#         items = results.get('files', [])
#         if items:
#             return items[0]['id']
#         else:
#             print(f'File "{file_name}" not found.')
#             return None


# # Example usage:
# creds_file = '/Users/jameshoang/Downloads/client_secret_827849408803-aohcmoa5lbtjn86gbfogjmpl68qgcdrv.apps.googleusercontent.com.json'
# drive_helper = GoogleDriveHelper(creds_file)
# file_name = 'Địa-bàii-16-17-18.docx'  # Replace with the name of your file
# file_id = drive_helper.get_file_id(file_name)
# if file_id:
#     txt_to_xlxs_instance = Txt_to_xlsx(file_id)

for i in range(500):
    print("Hello")
