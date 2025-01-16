import os
import logging
import tempfile
from flask import Flask, redirect, url_for, send_file, jsonify, render_template_string
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from openpyxl import load_workbook
import excel2img  # Ensure this is in requirements.txt

app = Flask(__name__)

# Load environment variables
load_dotenv()

# SharePoint credentials and file paths
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')
SHAREPOINT_FILE_URL = os.getenv('SHAREPOINT_FILE_URL')
SHAREPOINT_USERNAME = os.getenv('SHAREPOINT_USERNAME')
SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')

# Logger setup
logger = logging.getLogger('FlaskAppLogger')
logger.setLevel(logging.INFO)

# Console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)

logger.addHandler(console_handler)

logger.info("Logger initialized.")


def download_sharepoint_file():
    """Download the Excel file from SharePoint"""
    try:
        # Set up SharePoint authentication
        credentials = UserCredential(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credentials)
        
        # Fetch the file from SharePoint
        file = ctx.web.get_file_by_server_relative_url(SHAREPOINT_FILE_URL)
        ctx.load(file)
        ctx.execute_query()

        # Download the file content
        with open('downloaded_file.xlsm', 'wb') as local_file:
            file_content = file.read()
            local_file.write(file_content)
        
        logger.info("File downloaded successfully.")
        return 'downloaded_file.xlsm'
    except Exception as e:
        logger.error(f"Failed to download SharePoint file: {e}")
        return None


def load_workbook_simple(file_path):
    """Load an unprotected Excel workbook using openpyxl"""
    try:
        workbook = load_workbook(file_path, data_only=True)
        logger.info(f"Successfully loaded workbook: {file_path}")
        return workbook
    except Exception as e:
        logger.error(f"Failed to load workbook '{file_path}': {e}")
        return None


def generate_image(sheet_name, cell_range, image_filename):
    """Generate an image from an Excel sheet"""
    try:
        # Download the file from SharePoint
        excel_file_path = download_sharepoint_file()
        if not excel_file_path:
            raise Exception("Unable to download Excel file from SharePoint.")
        
        # Load workbook
        workbook = load_workbook_simple(excel_file_path)
        if not workbook:
            raise Exception("Unable to load workbook for image export.")
        
        # Check if the sheet exists
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
        
        # Hide gridlines
        wb_sheet = workbook[sheet_name]
        wb_sheet.sheet_view.showGridLines = False

        # Save the workbook to a temporary file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            temp_decrypted_path = tmp.name
            workbook.save(temp_decrypted_path)

        # Define image path in the static/images directory
        image_path = os.path.join('static', 'images', image_filename)
        
        # Ensure the images directory exists
        os.makedirs(os.path.dirname(image_path), exist_ok=True)
        
        # Generate image using excel2img
        excel2img.export_img(temp_decrypted_path, image_path, sheet_name, cell_range)
        logger.info(f"Generated image for '{sheet_name}' -> {image_path}")

        return image_path

    except Exception as e:
        logger.error(f"Error generating image for '{sheet_name}': {e}")
        raise e


@app.route('/')
def home():
    """Main route to render the web page"""
    active_pages = ['Morning', 'Evening', 'Night', 'Friday']

    page_ranges = {
        'Morning': ('Morning', 'A1:H33'),
        'Evening': ('Evening', 'A1:H33'),
        'Night': ('Night', 'A1:H33'),
        'Friday': ('Friday', 'A1:H33')
    }

    pages = []
    for page in active_pages:
        if page not in page_ranges:
            logger.error(f"No cell range defined for page '{page}'. Skipping.")
            continue

        sheet_name, cell_range = page_ranges[page]
        image_filename = f"{page}.png"
        
        try:
            image_path = generate_image(sheet_name, cell_range, image_filename)
            image_url = url_for('static', filename=f'images/{image_filename}')
            pages.append({'url': image_url, 'header': f"{sheet_name} Shift"})
        except Exception as e:
            logger.error(f"Failed to generate image for page '{page}': {e}")

    if not pages:
        return "<h1>No images available for the active pages.</h1>"

    # HTML content for the page
    html_content = f"""
    <html>
    <head>
      <style>
        html, body {{
          margin: 0; padding: 0;
          width: 100%; height: 100%;
          background-color: white;
          display: flex; flex-direction: column; align-items: center;
        }}
        #header {{
          position: fixed;
          top: 10px;
          left: 0; right: 0;
          text-align: center;
          font-size: 24px;
          color: black;
          padding: 10px 0;
        }}
        #image-container {{
          width: 90vw;
          height: 90vh;
          margin-top: 60px;
          display: flex; justify-content: center; align-items: center;
        }}
        img {{
          max-width: 100%;
          max-height: 100%;
          object-fit: contain;
          display: none;
        }}
        img.active {{
          display: block;
        }}
      </style>
    </head>
    <body>
      <div id="header">Loading...</div>
      <div id="image-container"></div>
      <script>
        const pages = {pages};
        let currentIndex = 0;
        
        window.onload = function() {{
          const header = document.getElementById('header');
          const imageContainer = document.getElementById('image-container');

          const imgElements = pages.map(page => {{
            const img = document.createElement('img');
            img.src = page.url;
            imageContainer.appendChild(img);
            return img;
          }}); 

          function showPage(idx) {{
            imgElements.forEach(img => img.classList.remove('active'));
            imgElements[idx].classList.add('active');
            header.innerText = pages[idx].header;
          }}

          setInterval(() => {{
            currentIndex = (currentIndex + 1) % pages.length;
            showPage(currentIndex);
          }}, 30000);

          showPage(currentIndex);
        }};</script>
    </body>
    </html>
    """

    return render_template_string(html_content, pages=pages)


if __name__ == '__main__':
    app.run(debug=True)
