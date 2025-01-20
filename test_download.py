import os
import logging
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# Setup logging
logger = logging.getLogger('TestLogger')
logger.setLevel(logging.DEBUG)  # Set to DEBUG for detailed logs
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

# Load environment variables
load_dotenv()

SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')
SHAREPOINT_FILE_URL = os.getenv('SHAREPOINT_FILE_URL')
SHAREPOINT_USERNAME = os.getenv('SHAREPOINT_USERNAME')
SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')

def download_sharepoint_file():
    """Download the Excel file from SharePoint"""
    try:
        # Validate environment variables
        if not all([SHAREPOINT_SITE_URL, SHAREPOINT_FILE_URL, SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD]):
            logger.error("One or more environment variables are missing.")
            return None

        logger.debug(f"SharePoint Site URL: {SHAREPOINT_SITE_URL}")
        logger.debug(f"SharePoint File URL: {SHAREPOINT_FILE_URL}")
        logger.debug(f"SharePoint Username: {SHAREPOINT_USERNAME}")
        # Do not log the password for security reasons

        # Set up SharePoint authentication using UserCredential
        credentials = UserCredential(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credentials)
        
        # Define the download path
        download_dir = 'downloads'
        os.makedirs(download_dir, exist_ok=True)
        download_path = os.path.join(download_dir, 'Plan.xlsm')
        
        # Attempt to download the file
        logger.debug(f"Attempting to download the file from {SHAREPOINT_FILE_URL} to {download_path}")
        response = ctx.web.get_file_by_server_relative_url(SHAREPOINT_FILE_URL).download(download_path)
        ctx.execute_query()
        
        logger.info("File downloaded successfully.")
        return download_path
    except Exception as e:
        logger.error(f"Failed to download SharePoint file: {e}")
        return None

if __name__ == "__main__":
    downloaded_file = download_sharepoint_file()
    if downloaded_file and os.path.exists(downloaded_file):
        logger.info(f"Downloaded file path: {downloaded_file}")
    else:
        logger.error("Download failed or file does not exist.")
