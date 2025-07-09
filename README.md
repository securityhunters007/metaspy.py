# 🕵️ MetaSpy: Metadata Extraction & Privacy Tool

MetaSpy is a powerful command-line tool built with Python to analyze and manage file metadata. It can extract hidden information from various file types for OSINT investigations and privacy audits. It also includes a feature to strip metadata from files to protect your privacy.

## Features

* **Multi-Format Support**: Extracts metadata from PDFs, Microsoft Office documents (`.docx`, `.pptx`, `.xlsx`), and various image formats (`.jpg`, `.png`, etc.).
* **Advanced Image Analysis**: Uses `ExifTool` to pull detailed EXIF data from images, including GPS coordinates, camera information, and software details.
* **Geolocation Links**: Automatically generates Google Maps links from GPS data found in images.
* **Metadata Stripping**: Removes all metadata from files to protect user privacy, creating a "clean" version of the file.
* **Flexible Reporting**: Outputs extracted data directly to the console or saves it in `.txt`, `.json`, or `.csv` formats.
* **Command-Line Interface**: Simple and efficient CLI for easy use in scripts and workflows.

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/securityhunters007/metaspy.py
    cd metaspy.py
    ```
2.  **Create and activate a virtual environment:**
    ```bash
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```
3.  **Install the required libraries:**
    ```bash
    pip install -r requirements.txt
    ```
4.  **Install ExifTool:**
    This tool requires a system-level installation of ExifTool. Follow the instructions at [exiftool.org](https://exiftool.org/).

## Usage

Make sure your virtual environment is activated before running the commands.

### Extracting Metadata

* **Print output to the console:**
    ```bash
    python metaspy.py file1.pdf image.jpg
    ```
* **Save output to a JSON file:**
    ```bash
    python metaspy.py my_document.docx -o json
    ```
* **Save a report for multiple files as a TXT file:**
    ```bash
    python metaspy.py document.pptx photo.png data.xlsx -o txt
    ```
### Stripping Metadata (Privacy)

* **Remove all metadata from a single file:**
    ```bash
    python metaspy.py sensitive_photo.jpg --strip
    ```
    This will create a clean file `sensitive_photo.jpg` and keep the original with its metadata in `sensitive_photo.jpg_original`.

### Getting Help

* **View all available commands and options:**
    ```bash
    python metaspy.py -h
    ```
