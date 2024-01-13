import os
import logging
import shutil
import tarfile
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import requests
import threading
from concurrent.futures import ThreadPoolExecutor
import time

# Configure global logging
base_local_log_directory = "/root/data/logs/" ## Enter the path in which logs needs to be saved. 

# Get the current time as a string for file naming
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

# Ensure the log directory exists
if not os.path.exists(base_local_log_directory):
    os.makedirs(base_local_log_directory)

def execute_with_retry(func, *args, **kwargs):
    retry_count = kwargs.pop('retry_count', 5)
    delay = kwargs.pop('delay', 5)
    for attempt in range(1, retry_count + 1):
        try:
            return func(*args, **kwargs)
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 503:
                safe_log(logging, f"Attempt {attempt} failed with HTTP 503 error. Retrying after {delay} seconds...")
                time.sleep(delay)
            else:
                raise
        except Exception as e:
            if attempt == retry_count:
                safe_log(logging, f"Operation failed after {retry_count} attempts: {e}")
                raise
            safe_log(logging, f"Attempt {attempt} failed with an exception. Retrying after {delay} seconds...")
            time.sleep(delay)

def setup_logger(site_name):
    log_file = f"{base_local_log_directory}sharepoint_downloads_{site_name}_{current_time}.log"
    logger = logging.getLogger(site_name)
    logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger

def safe_log(logger, message, level='info'):
    """Log messages safely with UTF-8 encoding."""
    encoded_message = message.encode('utf-8', errors='replace').decode('utf-8')
    if level == 'info':
        logger.info(encoded_message)
    elif level == 'error':
        logger.error(encoded_message)



# List of SharePoint sites
sharepoint_sites = [
    {
        "site_url": "https://<sharepoint-name>.sharepoint.com/sites/first-site",
        "site_base_url": "/sites/first-site/Shared Documents",
        "client_id": "<Application Client ID>",
        "client_secret": "<Client Secrect Value>"
    },
     {   
        "site_url": "https://<sharepoint-name>.sharepoint.com/sites/second-site",
        "site_base_url": "/sites/second-site/Shared Documents",
        "client_id": "<Application Client ID>",
        "client_secret": "<Client Secrect Value>"
     },
        # Add more SharePoint sites as needed
]

# Local path for downloads
base_local_download_path = "/root/data/" # Path to which backups needs to be taken

# Ensure the directory exists
if not os.path.exists(base_local_download_path):
    os.makedirs(base_local_download_path)

# Functions (download_file, list_and_download_files_and_folders, make_tarfile) remain the same...
# Function to download a file from SharePoint
def threaded_download_file(ctx, file_url, local_path, logger):
    try:
        response = execute_with_retry(File.open_binary, ctx, file_url)
        with open(local_path, "wb") as local_file:
            local_file.write(response.content)
        safe_log(logger, f"Successfully downloaded {file_url}".replace('\u200b', ''))
    except Exception as e:
        safe_log(logger, f"Error downloading {file_url}: {e}".replace('\u200b', ''))



# Function to list and download all files and folders from a given folder
def list_and_download_files_and_folders(url, folder_url, local_folder_path, client_id, client_secret, logger):
    context_auth = AuthenticationContext(url)
    if not context_auth.acquire_token_for_app(client_id, client_secret):
        safe_log(logger, f"Authentication failed for {url}", 'error')
        return
    ctx = ClientContext(url, context_auth)
    web = ctx.web
    folder = web.get_folder_by_server_relative_url(folder_url)
    ctx.load(folder)
    execute_with_retry(ctx.execute_query)
    # Check if 'ServerRelativeUrl' is loaded
    if 'ServerRelativeUrl' in folder.properties:
        safe_log(logger, f"Accessing Folder: {folder.properties['ServerRelativeUrl']}", 'info')
    else:
        safe_log(logger, f"ServerRelativeUrl not found for the folder: {folder_url}", 'error')
        return  # or handle as appropriate
        
    safe_log(logger, f"Accessing Folder: {folder.properties['ServerRelativeUrl']}")
    
    if not os.path.exists(local_folder_path):
        os.makedirs(local_folder_path)

    files = folder.files
    ctx.load(files)
    execute_with_retry(ctx.execute_query)
    with ThreadPoolExecutor(max_workers=10) as executor:
        for file in files:
            ctx.load(file)
            execute_with_retry(ctx.execute_query)
            file_name = file.properties['Name']
            file_url = file.properties['ServerRelativeUrl']
            safe_log(logger, f"Queueing download for {file_name}")
            executor.submit(threaded_download_file, ctx, file_url, os.path.join(local_folder_path, file_name), logger)

        folders = folder.folders
        ctx.load(folders)
        execute_with_retry(ctx.execute_query)
        for subfolder in folders:
            ctx.load(subfolder)
            execute_with_retry(ctx.execute_query)
            folder_name = subfolder.properties['Name']
            folder_url = subfolder.properties['ServerRelativeUrl']
            subfolder_path = os.path.join(local_folder_path, folder_name)
            list_and_download_files_and_folders(url, folder_url, subfolder_path, client_id, client_secret, logger)
 
# Function to tar a directory
def make_tarfile(output_filename, source_dir):
    with tarfile.open(output_filename, "w:gz") as tar:
        tar.add(source_dir, arcname=os.path.basename(source_dir))
    safe_log(logging, f"Created tar archive {output_filename}")
    
def process_site(site_details):
    site_url = site_details["site_url"]
    site_base_url = site_details["site_base_url"]
    client_id = site_details["client_id"]
    client_secret = site_details["client_secret"]
    # Extract the site name from the URL
    site_name = site_url.split('/')[-1]

    # Local path for this site's downloads
    local_download_path = f"{base_local_download_path}{site_name}_{current_time}"
    
    logger = setup_logger(site_name) 
    
    # Log the start of the process for this site
    safe_log(logger, f"Starting download script for {site_name}", 'info')

    # Download files and folders
    list_and_download_files_and_folders(site_url, site_base_url, local_download_path, client_id, client_secret, logger)
    # Tar the downloaded directory
    tar_filename = f"{local_download_path}.tar.gz"
    make_tarfile(tar_filename, local_download_path)

    # Log the end of the process for this site
    safe_log(logger, f"Download script finished for {site_name}")

    # Remove the directory after making the tar file
    try:
        shutil.rmtree(local_download_path)
        safe_log(logger, f"Successfully removed directory: {local_download_path}")
    except Exception as e:
        safe_log(logger, f"Error removing directory {local_download_path}: {e}")
    
    # Close the file handler to properly close the log file
    file_handler.close()

# Process each site
for site in sharepoint_sites:
    process_site(site)


