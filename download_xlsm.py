from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import os
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger('GoogleDriveLogger')

# Path to your downloaded JSON credentials
SERVICE_ACCOUNT_FILE = '/workspace/DailyPlan/credintials/service-account-file.json'  # <-- Update this path

# Define the scope
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

def authenticate_drive():
    """
    Authenticate with Google Drive API using service account credentials.
    """
    try:
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('drive', 'v3', credentials=credentials)
        logger.debug("Authenticated successfully with Google Drive API.")
        return service
    except Exception as e:
        logger.exception(f"Authentication failed: {e}")
        raise

def get_file_id_from_url(file_url):
    """
    Extract the file ID from a Google Drive file URL.
    """
    try:
        # Example URL formats:
        # https://drive.google.com/file/d/FILE_ID/view?usp=sharing
        # https://docs.google.com/spreadsheets/d/FILE_ID/edit#gid=0
        import re
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', file_url)
        if match:
            file_id = match.group(1)
            logger.debug(f"Extracted File ID: {file_id}")
            return file_id
        else:
            raise ValueError("Could not extract file ID from the URL.")
    except Exception as e:
        logger.exception(f"Failed to extract file ID: {e}")
        raise

def download_xlsm_file(service, file_id, download_dir='downloads'):
    """
    Download the XLSM file from Google Drive.
    
    Args:
        service: Authenticated Google Drive service instance.
        file_id: ID of the file to download.
        download_dir: Directory to save the downloaded file.
    """
    try:
        # Get file metadata to obtain the name
        file = service.files().get(fileId=file_id, fields='name, mimeType').execute()
        file_name = file.get('name')
        mime_type = file.get('mimeType')
        logger.debug(f"File Name: {file_name}, MIME Type: {mime_type}")

        # Check if the file is indeed an XLSM
        if mime_type != 'application/vnd.ms-excel.sheet.macroEnabled.12':
            logger.warning(f"The file is not an XLSM. Detected MIME Type: {mime_type}")

        # Prepare the download request
        request = service.files().get_media(fileId=file_id)
        fh = io.FileIO(os.path.join(download_dir, file_name), 'wb')
        downloader = MediaIoBaseDownload(fh, request)

        # Ensure the download directory exists
        os.makedirs(download_dir, exist_ok=True)

        done = False
        while not done:
            status, done = downloader.next_chunk()
            logger.debug(f"Download {int(status.progress() * 100)}%.")

        logger.info(f"File downloaded successfully to {os.path.join(download_dir, file_name)}")

    except Exception as e:
        logger.exception(f"An error occurred while downloading the file: {e}")
        raise

def main():
    # Google Drive File URL
    file_url = "https://docs.google.com/spreadsheets/d/10yGaM_ZNzhFmdvqsPB1jWfiVNUllvMHv/edit?usp=sharing&ouid=115990900741931691948&rtpof=true&sd=true"
    
    # Extract File ID
    file_id = get_file_id_from_url(file_url)
    
    # Authenticate and build the Drive service
    service = authenticate_drive()
    
    # Define download directory and path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    download_dir = "downloads"
    
    # Download the file
    download_xlsm_file(service, file_id, download_dir)

if __name__ == "__main__":
    main()
