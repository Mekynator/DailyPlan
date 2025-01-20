import os
from dotenv import load_dotenv
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import tempfile
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('TestLogger')

# Load environment variables
load_dotenv()

# SharePoint credentials and file paths
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')
SHAREPOINT_FILE_URL = os.getenv('SHAREPOINT_FILE_URL')
SHAREPOINT_USERNAME = os.getenv('SHAREPOINT_USERNAME')
SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')

def download_sharepoint_file():
    """Download the Excel file from SharePoint"""
    try:
        # Verify that all environment variables are set
        if not all([SHAREPOINT_SITE_URL, SHAREPOINT_FILE_URL, SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD]):
            logger.error("One or more SharePoint environment variables are not set.")
            raise ValueError("SharePoint configuration incomplete.")
        
        # Set up SharePoint authentication
        credentials = UserCredential(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credentials)
        
        # Fetch the file from SharePoint
        file = ctx.web.get_file_by_server_relative_url(SHAREPOINT_FILE_URL)
        ctx.load(file)
        ctx.execute_query()
        
        # Use tempfile to create a temporary file for the downloaded Excel file
        with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as temp_file:
            file.download(temp_file)
            ctx.execute_query()
            temp_file_path = temp_file.name
        
        logger.info(f"File downloaded successfully to {temp_file_path}")
        return temp_file_path
    except Exception as e:
        logger.error(f"Failed to download SharePoint file: {e}")
        return None

if __name__ == '__main__':
    download_sharepoint_file()
