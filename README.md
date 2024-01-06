
# SharePoint Downloader

This script (`online-sharepoint-backup.py`) is designed to automate the downloading of files and folders from specified SharePoint sites. It logs the process and saves the files locally.

## Features

- Authentication with SharePoint
- Download files and folders from specified sites
- Log activities in a dedicated log file
- Archive downloaded content using tar

## Prerequisites

- Python 3.x
- Office365-REST-Python-Client library


## Configuration
Update the sharepoint_sites list in the script with your SharePoint site details:

python
Copy code
sharepoint_sites = [

    {
        "site_url": "YOUR_SITE_URL",
        "site_base_url": "YOUR_SITE_BASE_URL",
        "client_id": "YOUR_CLIENT_ID",
        "client_secret": "YOUR_CLIENT_SECRET"
    },
    
    # Add more SharePoint sites as needed
]

## Usage
Run the script with Python:

python sharepoint_downloader.py

Check the E:/Download/logs/ directory for log files to monitor the download process.
