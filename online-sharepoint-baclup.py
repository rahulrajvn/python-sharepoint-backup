import os
import logging
import shutil
import tarfile
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Configure global logging
base_local_log_directory = "E:/Download/logs/" ## Enter the path in which logs needs to be saved. 

# Ensure the log directory exists
if not os.path.exists(base_local_log_directory):
    os.makedirs(base_local_log_directory)

def setup_logger(site_name):
    log_file = f"{base_local_log_directory}sharepoint_downloads_{site_name}_{current_time}.log"
    logger = logging.getLogger(site_name)
    logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(log_file)
    formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger, file_handler

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
base_local_download_path = "E:/Download/data/" # Path to which backups needs to be taken

# Ensure the directory exists
if not os.path.exists(base_local_download_path):
    os.makedirs(base_local_download_path)

# Functions (download_file, list_and_download_files_and_folders, make_tarfile) remain the same...
# Function to download a file from SharePoint
def download_file(ctx, file_url, local_path):
    response = File.open_binary(ctx, file_url)
    with open(local_path, "wb") as local_file:
        local_file.write(response.content)

# Function to list and download all files and folders from a given folder
def list_and_download_files_and_folders(url, folder_url, local_folder_path,client_id, client_secret):
    # Extract the site name from the URL
    site_name = url.split('/')[-1]
    
    context_auth = AuthenticationContext(url)
    if not context_auth.acquire_token_for_app(client_id, client_secret):
        logger.error(f"Authentication failed for {site_name}")
        return
    ctx = ClientContext(url, context_auth)
    web = ctx.web
    folder = web.get_folder_by_server_relative_url(folder_url)
    ctx.load(folder)
    ctx.execute_query()
    print(f"Accessing Folder: {folder.properties['ServerRelativeUrl']}")
    
    # Ensure local folder exists
    if not os.path.exists(local_folder_path):
        os.makedirs(local_folder_path)
    # List and download files in the folder
    files = folder.files
    ctx.load(files)
    ctx.execute_query()
    for file in files:
        file_name = file.properties['Name']
        file_url = file.properties['ServerRelativeUrl']
        print(f"Downloading File: {file_name}")
        download_file(ctx, file_url, os.path.join(local_folder_path, file_name))

    # List folders in the folder and recursively list and download
    folders = folder.folders
    ctx.load(folders)
    ctx.execute_query()
    for folder in folders:
        folder_name = folder.properties['Name']
        folder_url = folder.properties['ServerRelativeUrl']
        print(f"Accessing Folder: {folder_name}")
        list_and_download_files_and_folders(url, folder_url, os.path.join(local_folder_path, folder_name),client_id, client_secret)


# Function to tar a directory
def make_tarfile(output_filename, source_dir):
    with tarfile.open(output_filename, "w:gz") as tar:
        tar.add(source_dir, arcname=os.path.basename(source_dir))
    logging.info(f"Created tar archive {output_filename}")
    
def process_site(site_details):
    site_url = site_details["site_url"]
    site_base_url = site_details["site_base_url"]
    client_id = site_details["client_id"]
    client_secret = site_details["client_secret"]
    # Extract the site name from the URL
    site_name = site_url.split('/')[-1]

    # Local path for this site's downloads
    local_download_path = f"{base_local_download_path}{site_name}_{current_time}"
    
    logger, file_handler = setup_logger(site_name)
    
    # Log the start of the process for this site
    logger.info(f"Starting download script for {site_name}")

    # Download files and folders
    list_and_download_files_and_folders(site_url, site_base_url, local_download_path, client_id, client_secret)

    # Tar the downloaded directory
    tar_filename = f"{local_download_path}.tar.gz"
    make_tarfile(tar_filename, local_download_path)

    # Log the end of the process for this site
    logger.info(f"Download script finished for {site_name}")

    # Remove the directory after making the tar file
    try:
        shutil.rmtree(local_download_path)
        logger.info(f"Successfully removed directory: {local_download_path}")
    except Exception as e:
        logger.error(f"Error removing directory {local_download_path}: {e}")

    
    # Close the file handler to properly close the log file
    file_handler.close()

# Process each site
for site in sharepoint_sites:
    process_site(site)
