import pandas as pd
import os
import time
import random
import requests
from urllib.parse import urlparse
import sys

def get_file_extension_from_content_type(content_type):
    """Map common content types to file extensions"""
    content_type_map = {
        'application/pdf': '.pdf',
        'application/vnd.ms-excel': '.xls',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        'application/zip': '.zip',
        'application/x-rar-compressed': '.rar',
        'application/x-7z-compressed': '.7z',
        'application/octet-stream': '.bin',
        'text/csv': '.csv',
        'application/json': '.json',
        'text/plain': '.txt',
        'application/xml': '.xml',
        'text/xml': '.xml'
    }
    
    # Clean content type (remove charset and other parameters)
    clean_content_type = content_type.split(';')[0].strip().lower()
    return content_type_map.get(clean_content_type, '.file')

def detect_file_type_from_content(content):
    """Detect file type from the first few bytes (magic numbers)"""
    if not content:
        return '.file'
    
    # Check magic numbers for common file types
    magic_numbers = {
        b'\x50\x4B\x03\x04': '.zip',           # ZIP
        b'\x50\x4B\x05\x06': '.zip',           # ZIP (empty)
        b'\x50\x4B\x07\x08': '.zip',           # ZIP (spanned)
        b'\x52\x61\x72\x21': '.rar',           # RAR
        b'\x37\x7A\xBC\xAF\x27\x1C': '.7z',   # 7-Zip
        b'\x25\x50\x44\x46': '.pdf',           # PDF
        b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1': '.xls',  # MS Office (old format)
        b'\x50\x4B\x03\x04': '.xlsx',          # XLSX (also ZIP-based)
        b'\x1F\x8B': '.gz',                    # GZIP
        b'\x42\x5A\x68': '.bz2',               # BZIP2
    }
    
    # Check against magic numbers
    for magic, ext in magic_numbers.items():
        if content.startswith(magic):
            # Special handling for ZIP-based formats
            if magic == b'\x50\x4B\x03\x04':
                # Could be ZIP, XLSX, DOCX, etc. - check content
                if b'xl/' in content[:1024] or b'[Content_Types].xml' in content[:1024]:
                    return '.xlsx'
                elif b'word/' in content[:1024]:
                    return '.docx'
                else:
                    return '.zip'
            return ext
    
    # If no magic number matches, try to detect based on content patterns
    content_str = content[:1024].decode('utf-8', errors='ignore').lower()
    
    if content_str.startswith('<?xml'):
        return '.xml'
    elif content_str.startswith('{') or content_str.startswith('['):
        return '.json'
    elif ',' in content_str and '\n' in content_str:  # Simple CSV detection
        return '.csv'
    
    return '.file'

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
        
        # Decide on the best extension
        if extension_from_content != '.file':
            final_extension = extension_from_content
            detection_method = "content analysis"
        elif extension_from_header != '.file':
            final_extension = extension_from_header
            detection_method = "HTTP header"
        else:
            final_extension = '.file'
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
    excel_file = os.path.join(project_dir, "Data", "Investments_sources.xlsx")
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
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        print(f"Make sure the file exists at: {excel_file}")
        return
    
    # Create "Investments" folder inside the existing Data folder
    data_dir = os.path.join(project_dir, "Data")
    investments_dir = os.path.join(data_dir, "Investments")
    
    if not os.path.exists(investments_dir):
        os.makedirs(investments_dir)
        print(f"Created Investments folder: {investments_dir}")
    else:
        print(f"Investments folder already exists: {investments_dir}")
    
    # Change to Investments directory
    os.chdir(investments_dir)
    print(f"Changed to Investments directory: {os.getcwd()}")
    
    previous_oblast = None
    
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
        year = row['Year']
        link = row['Link']
        
        print(f"\n{'='*60}")
        print(f"Processing row {index + 1}/{len(df)}: {oblast} - {year}")
        print(f"{'='*60}")
        
        # Check if we need to create a new folder (first row or oblast changed)
        if previous_oblast is None or oblast != previous_oblast:
            # If not the first iteration, go back to Investments folder
            if previous_oblast is not None:
                os.chdir(investments_dir)
                print(f"Returned to Investments directory: {os.getcwd()}")
            
            # Create folder for the oblast if it doesn't exist
            if not os.path.exists(oblast):
                os.makedirs(oblast)
                print(f"Created folder: {oblast}")
            
            # Change to the oblast folder
            os.chdir(oblast)
            print(f"Changed directory to: {os.getcwd()}")
        
        # Generate base filename (without extension) - changed format
        base_filename = f"{year}_Investments"
        
        # Download the file
        success, final_filename = download_file(link, base_filename)
        
        if success:
            print(f"✓ File saved as: {final_filename}")
            successful_downloads += 1
        else:
            print(f"✗ Failed to download file for {oblast} {year}")
            failed_downloads += 1
            failed_downloads_list.append(f"{oblast}_{year}")  # Add to failures list
        
        # Update previous_oblast for next iteration
        previous_oblast = oblast
    
    # Return to original directory when done
    os.chdir(original_dir)
    
    # Create summary report file in the original directory (Code folder) 
    from datetime import datetime
    
    # Generate timestamp for the report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"download_report_{timestamp}.txt"
    
    # Write comprehensive report to file
    with open(report_filename, 'w', encoding='utf-8') as report_file:
        report_file.write("="*60 + "\n")
        report_file.write("KAZAKHSTAN STATISTICS DOWNLOAD REPORT\n")
        report_file.write("="*60 + "\n")
        report_file.write(f"Report generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        report_file.write(f"Investments directory: {investments_dir}\n")
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
            report_file.write("Failed items (Oblast_Year format):\n")
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
    print("DOWNLOAD PROCESS COMPLETED!")
    print(f"{'='*60}")
    print(f"✓ Report saved to: {report_filename}")
    print(f"✓ Investments directory: {investments_dir}")
    print(f"✓ Files processed: {successful_downloads + failed_downloads}")
    if failed_downloads > 0:
        print(f"⚠ Failures: {failed_downloads} (see report for details)")
    else:
        print("🎉 All downloads successful!")

if __name__ == "__main__":
    main()