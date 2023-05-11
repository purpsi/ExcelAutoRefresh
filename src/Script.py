import win32com.client
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

def download_file_from_sharepoint(site_url, user_name, password, file_url, file_name):
    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(user_name, password):
        ctx = ClientContext(site_url, ctx_auth)
        with open(file_name, "wb") as local_file:
            File.open_binary(ctx, file_url, local_file)

def refresh_excel(file_paths):
    # Create an instance of Excel Application
    xlapp = win32com.client.DispatchEx("Excel.Application")
    
    for file_path in file_paths:
        # Open the workbook in Excel
        wb = xlapp.Workbooks.Open(file_path)
    
        # Refresh all data connections.
        wb.RefreshAll()

        # Save and close
        wb.Save()
        wb.Close()

    # Quit the application
    xlapp.Quit()

# Download files from SharePoint
site_url = "https://your_sharepoint_site_url"
user_name = "your_username"
password = "your_password"
file_urls = ["file_url_1", "file_url_2", "file_url_3"]  # URLs of the files on SharePoint
file_names = ["file_name_1.xlsx", "file_name_2.xlsx", "file_name_3.xlsx"]  # Names to save the files as locally

for file_url, file_name in zip(file_urls, file_names):
    download_file_from_sharepoint(site_url, user_name, password, file_url, file_name)

# Refresh downloaded files
refresh_excel(file_names)
