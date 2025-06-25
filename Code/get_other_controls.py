import pandas as pd
import os
import time
import random
import requests
from urllib.parse import urlparse
import sys

def get_file_extension_from_content_type(content_type):
    """Map common content types to file extensions (focused on Excel formats)"""
    content_type_map = {
        'application/vnd.ms-excel': '.xls',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        'application/octet-stream': '.xlsx',  # Sometimes Excel files are served as octet-stream
        'text/csv': '.csv',
        'application/json': '.json',
        'text/plain': '.txt',
        'application/xml': '.xml',
        'text/xml': '.xml'
    }
    
    # Clean content type (remove charset and other parameters)
    clean_content_type = content_type.split(';')[0].strip().lower()
    return content_type_map.get(clean_content_type, '.xlsx')  # Default to .xlsx

def detect_file_type_from_content(content):
    """Detect file type from the first few bytes (focused on Excel formats)"""
    if not content:
        return '.xlsx'  # Default to .xlsx
    
    # Check magic numbers for Excel file types
    magic_numbers = {
        b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1': '.xls',  # MS Office (old format)
        b'\x50\x4B\x03\x04': '.xlsx',          # XLSX (ZIP-based)
        b'\x50\x4B\x05\x06': '.xlsx',          # XLSX (empty ZIP)
        b'\x50\x4B\x07\x08': '.xlsx',          # XLSX (spanned ZIP)
    }
    
    # Check against magic numbers
    for magic, ext in magic_numbers.items():
        if content.startswith(magic):
            # For ZIP-based formats, check if it's really an Excel file
            if magic.startswith(b'\x50\x4B'):
                if b'xl/' in content[:1024] or b'[Content_Types].xml' in content[:1024]:
                    return '.xlsx'
            return ext
    
    # If no magic number matches, default to .xlsx
    return '.xlsx'

def sanitize_filename(filename):
    """Sanitize filename by removing or replacing invalid characters"""
    # Characters that are not allowed in filenames on Windows/Linux
    invalid_chars = '<>:"/\\|?*'
    
    # Replace invalid characters with underscores
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Remove leading/trailing whitespace and dots
    filename = filename.strip('. ')
    
    # Limit length to avoid filesystem issues
    if len(filename) > 200:
        filename = filename[:200]
    
    return filename

def download_file(url, base_filename):
    """Download file from URL and save with proper extension"""
    try:
        # Add headers to mimic a browser request
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        print(f"Downloading from: {url}")
        response = requests.get(url, headers=headers, stream=True, timeout=30)
        response.raise_for_status()
        
        # First, try to get extension from Content-Type header
        content_type = response.headers.get('content-type', '')
        extension_from_header = get_file_extension_from_content_type(content_type)
        
        # Read the first chunk to detect file type from content
        content_chunks = []
        first_chunk = None
        
        for chunk in response.iter_content(chunk_size=8192):
            if first_chunk is None:
                first_chunk = chunk
                # Detect file type from first chunk
                extension_from_content = detect_file_type_from_content(chunk)
            content_chunks.append(chunk)
        
        # Decide on the best extension (prioritize content detection for Excel files)
        if extension_from_content in ['.xls', '.xlsx']:
            final_extension = extension_from_content
            detection_method = "content analysis"
        elif extension_from_header in ['.xls', '.xlsx']:
            final_extension = extension_from_header
            detection_method = "HTTP header"
        else:
            final_extension = '.xlsx'  # Default to .xlsx for controls data
            detection_method = "default"
        
        # Create final filename
        final_filename = base_filename + final_extension
        
        # Save the file
        with open(final_filename, 'wb') as file:
            for chunk in content_chunks:
                file.write(chunk)
        
        print(f"Successfully downloaded: {final_filename}")
        print(f"File type detected via {detection_method}: {final_extension}")
        print(f"Content-Type header: {content_type}")
        
        return True, final_filename
        
    except requests.exceptions.RequestException as e:
        print(f"Error downloading {url}: {str(e)}")
        return False, None
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return False, None

def main():
    # Store the original directory
    original_dir = os.getcwd()
    print(f"Starting directory: {original_dir}")
    
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Script directory: {script_dir}")
    
    # Determine project directory based on script location
    # If script is in Code subfolder, go up one level
    # If script is run from project root, use current directory
    if os.path.basename(script_dir) == "Code":
        project_dir = os.path.dirname(script_dir)
    else:
        # Script might be run from project root or moved elsewhere
        # Try to find the project directory by looking for Data folder
        if os.path.exists(os.path.join(original_dir, "Data")):
            project_dir = original_dir
        elif os.path.exists(os.path.join(os.path.dirname(original_dir), "Data")):
            project_dir = os.path.dirname(original_dir)
        else:
            print("ERROR: Cannot locate project directory with Data folder")
            print("Please ensure you're running from the project root or Code subfolder")
            return
    
    print(f"Project directory: {project_dir}")
    
    # Path to the Excel file in the Data folder
    excel_file = os.path.join(project_dir, "Data", "Other_controls_sources.xlsx")
    print(f"Excel file path: {excel_file}")
    
    # Verify the Excel file exists before proceeding
    if not os.path.exists(excel_file):
        print(f"ERROR: Excel file not found at: {excel_file}")
        print(f"Current working directory: {original_dir}")
        print(f"Script directory: {script_dir}")
        print(f"Project directory detected as: {project_dir}")
        return
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_file)
        print(f"Loaded {len(df)} rows from Excel file")
        
        # Verify required columns exist
        required_columns = ['Oblast', 'Type', 'Link']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"ERROR: Missing required columns: {missing_columns}")
            print(f"Available columns: {list(df.columns)}")
            return
            
        print(f"Required columns found: {required_columns}")
        
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        print(f"Make sure the file exists at: {excel_file}")
        return
    
    # Create "Controls" folder inside the existing Data folder
    data_dir = os.path.join(project_dir, "Data")
    controls_dir = os.path.join(data_dir, "Controls")
    
    if not os.path.exists(controls_dir):
        os.makedirs(controls_dir)
        print(f"Created Controls folder: {controls_dir}")
    else:
        print(f"Controls folder already exists: {controls_dir}")
    
    # Change to Controls directory
    os.chdir(controls_dir)
    print(f"Changed to Controls directory: {os.getcwd()}")
    
    # Keep track of download statistics
    successful_downloads = 0
    failed_downloads = 0
    failed_downloads_list = []  # Track specific failures
    
    # Iterate through data rows
    for index, row in df.iterrows():
        # Sleep with random interval (1-10 seconds)
        sleep_time = random.randint(1, 10)
        print(f"\nSleeping for {sleep_time} seconds...")
        time.sleep(sleep_time)
        
        oblast = row['Oblast']
        control_type = row['Type']
        link = row['Link']
        
        print(f"\n{'='*60}")
        print(f"Processing row {index + 1}/{len(df)}: {oblast} - {control_type}")
        print(f"{'='*60}")
        
        # Generate base filename (without extension) - Oblast_Type format
        base_filename = sanitize_filename(f"{oblast}_{control_type}")
        
        # Download the file
        success, final_filename = download_file(link, base_filename)
        
        if success:
            print(f"✓ File saved as: {final_filename}")
            successful_downloads += 1
        else:
            print(f"✗ Failed to download file for {oblast} - {control_type}")
            failed_downloads += 1
            failed_downloads_list.append(f"{oblast} - {control_type}")  # Add to failures list
    
    # Return to original directory when done
    os.chdir(original_dir)
    
    # Create summary report file in the original directory
    from datetime import datetime
    
    # Generate timestamp for the report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"Controls_download_report_{timestamp}.txt"
    
    # Write comprehensive report to file
    with open(report_filename, 'w', encoding='utf-8') as report_file:
        report_file.write("="*60 + "\n")
        report_file.write("CONTROLS DATA DOWNLOAD REPORT\n")
        report_file.write("="*60 + "\n")
        report_file.write(f"Report generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        report_file.write(f"Controls directory: {controls_dir}\n")
        report_file.write(f"Excel source file: {excel_file}\n\n")
        
        report_file.write("SUMMARY:\n")
        report_file.write("-"*30 + "\n")
        report_file.write(f"Successful downloads: {successful_downloads}\n")
        report_file.write(f"Failed downloads: {failed_downloads}\n")
        report_file.write(f"Total files processed: {successful_downloads + failed_downloads}\n\n")
        
        if failed_downloads > 0:
            report_file.write("FAILED DOWNLOADS:\n")
            report_file.write("-"*30 + "\n")
            report_file.write(f"Total failures: {failed_downloads}\n\n")
            report_file.write("Failed Oblast - Type combinations:\n")
            for i, failure in enumerate(failed_downloads_list, 1):
                report_file.write(f"{i:2d}. {failure}\n")
        else:
            report_file.write("SUCCESS:\n")
            report_file.write("-"*30 + "\n")
            report_file.write("🎉 All downloads completed successfully!\n")
            report_file.write("No failures to report.\n")
        
        report_file.write("\n" + "="*60 + "\n")
        report_file.write("END OF REPORT\n")
        report_file.write("="*60 + "\n")
    
    # Print brief console summary
    print(f"\n{'='*60}")
    print("CONTROLS DATA DOWNLOAD PROCESS COMPLETED!")
    print(f"{'='*60}")
    print(f"✓ Report saved to: {report_filename}")
    print(f"✓ Controls directory: {controls_dir}")
    print(f"✓ Files processed: {successful_downloads + failed_downloads}")
    if failed_downloads > 0:
        print(f"⚠ Failures: {failed_downloads} (see report for details)")
    else:
        print("🎉 All downloads successful!")

if __name__ == "__main__":
    main()