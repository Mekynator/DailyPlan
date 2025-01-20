from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import logging
import os

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger('TestLogger')

# SharePoint credentials and URLs
site_url = "https://dscloud-my.sharepoint.com/personal/mark_szeibert_sallinggroup_com"
file_url = "/personal/mark_szeibert_sallinggroup_com/Documents/Plan.xlsm"
download_path = "downloads/Plan.xlsm"
username = os.getenv("mark.szeibert@sallinggroup.com")  # Use environment variables
password = os.getenv("SHAREPOINT_PASSWORD")  # Use environment variables

def download_file():
    try:
        # Ensure download directory exists
        os.makedirs(os.path.dirname(download_path), exist_ok=True)
        
        # Initialize ClientContext with credentials
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        logger.debug(f"Connecting to SharePoint site: {site_url}")
        
        # Attempt to retrieve the file
        file = ctx.web.get_file_by_server_relative_url(file_url)
        ctx.load(file)
        ctx.execute_query()
        
        # Download the file
        with open(download_path, "wb") as local_file:
            file.download(local_file)
            ctx.execute_query()
        
        logger.info(f"File downloaded successfully to {download_path}")
    
    except IndexError as ie:
        logger.error(f"List index out of range error: {ie}")
    
    except Exception as e:
        # Log the exception with traceback for deeper insights
        logger.exception(f"Failed to download SharePoint file: {e}")

if __name__ == "__main__":
    download_file()
