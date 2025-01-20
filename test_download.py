from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger('TestLogger')

# SharePoint credentials and URLs
site_url = "https://dscloud-my.sharepoint.com/personal/mark_szeibert_sallinggroup_com"
file_url = "/personal/mark_szeibert_sallinggroup_com/Documents/Plan.xlsm"
download_path = "downloads/Plan.xlsm"
username = "mark.szeibert@sallinggroup.com"
password = ""  # Consider using environment variables or secure storage

try:
    ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
    logger.debug(f"Connecting to SharePoint site: {site_url}")
    
    # Retrieve the file
    response = ctx.web.get_file_by_server_relative_url(file_url).download(download_path)
    ctx.execute_query()
    
    logger.info(f"File downloaded successfully to {download_path}")
except IndexError as ie:
    logger.error(f"List index out of range error: {ie}")
except Exception as e:
    logger.error(f"Failed to download SharePoint file: {e}")
