import os
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import app_functions as f
import time

# Base URL of the repository
base_url = "https://hurricanes.ral.ucar.edu/repository/data/adecks_open/"
print(" ####### RUNNING VERSION 3 #######")

# Years to loop through (starting from 2003 to the current year)
start_year = datetime.now().year
current_year = datetime.now().year

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Use absolute paths
download_folder = os.path.join(script_dir, "forecast_data")
csv_path = os.path.join(script_dir, 'storm_adeck_directory.csv')

# Create the folder if it doesn't exist
os.makedirs(download_folder, exist_ok=True)


def extract_storm_info(dat_file_path):
    """
    Extract storm name, start date, and end date from the .dat file's content.
    """
    storm_name = None
    storm_start_date = None
    storm_end_date = None
    dates = []
    storm_name_dict = {}

    # List of known basin codes
    basin_codes = ["AL", "EP", "CP", "WP", "IO", "SH"]

    # Set of number words up to 20
    number_words_set = {
        'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE', 'TEN',
        'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN',
        'EIGHTEEN', 'NINETEEN', 'TWENTY'
    }

    with open(dat_file_path, 'r') as file:
        for line in file:
            columns = line.strip().split()

            if len(columns) >= 28:
                potential_name = columns[27].strip().rstrip(',').upper()

                # Add potential_name to the dictionary
                if potential_name:
                    storm_name_dict[potential_name] = storm_name_dict.get(potential_name, 0) + 1

            # Parse dates as before
            if len(columns) >= 3:
                date_str = columns[2].strip()
                date_str = re.sub(r'[^0-9]', '', date_str)
                if len(date_str) == 10:
                    try:
                        date = datetime.strptime(date_str, "%Y%m%d%H")
                        dates.append(date)
                    except ValueError:
                        print(f"Error parsing date: {date_str} in file {dat_file_path}")

    # Step 1: Remove basin codes from the dictionary keys
    storm_name_dict = {k: v for k, v in storm_name_dict.items() if k not in basin_codes}

    # Step 2: Remove entries that contain digits
    storm_name_dict = {k: v for k, v in storm_name_dict.items() if not any(char.isdigit() for char in k)}

    # Step 3: Remove entries that are in number_words
    storm_name_dict = {k: v for k, v in storm_name_dict.items() if k not in number_words_set}

    # Get list of storm names
    storm_names = list(storm_name_dict.keys())

    # Determine the storm name based on the cleaned list
    if len(storm_names) == 0:
        storm_name = 'DISTURBANCE'
    elif storm_names == ['INVEST']:
        storm_name = 'INVEST'
    else:
        # Remove 'INVEST' if other names are present
        if 'INVEST' in storm_names:
            storm_names.remove('INVEST')

        if len(storm_names) == 1:
            storm_name = storm_names[0]
        else:
            # Multiple names remain
            # Attempt to strip basin codes from ends
            cleaned_names = set()
            for name in storm_names:
                cleaned_name = name
                for basin_code in basin_codes:
                    if name.endswith(basin_code):
                        name_without_basin = name[:-len(basin_code)].strip()
                        # Ensure the name is non-empty and alphabetical
                        if name_without_basin.isalpha():
                            cleaned_name = name_without_basin
                            break  # Stop checking after removing basin code
                cleaned_names.add(cleaned_name)

            if len(cleaned_names) == 1:
                storm_name = cleaned_names.pop()
            else:
                # Names don't match, take the last one
                storm_name = storm_names[-1]

  # Deduce the start and end dates and extract the year
    if dates:
        storm_start_datetime = min(dates)
        storm_end_datetime = max(dates)
        storm_start_date = storm_start_datetime.strftime("%Y-%m-%d %H:%M:%S")
        storm_end_date = storm_end_datetime.strftime("%Y-%m-%d %H:%M:%S")
        storm_year = storm_end_datetime.year  # Extract the year
    else:
        storm_start_date = None
        storm_end_date = None
        storm_year = None

    return storm_name.title(), storm_start_date, storm_end_date, storm_year

def download_and_update_storm_data(file_url, save_path, csv_path):
    """
    Download the .dat file from the given URL, extract storm information,
    and update the CSV file with the new or updated storm data.
    """
    # Download the data
    download_file(file_url, save_path)
    
    # Process the .dat file to extract storm info
    storm_name, storm_start_date, storm_end_date, storm_year = extract_storm_info(save_path)
    
    # Extract Basin and Storm Number from the file name
    file_name = os.path.basename(save_path)
    basin = file_name[1:3]  # Adjust based on actual file naming convention
    storm_number = file_name[3:5]  # Adjust based on actual file naming convention

    # Read existing CSV or create a new DataFrame
    if os.path.exists(csv_path):
        df = pd.read_csv(csv_path)
    else:
        df = pd.DataFrame(columns=['Basin', 'Storm_Number', 'Storm_Name', 'Storm_Start_Date', 'Storm_End_Date','Storm_Year','adeck_path'])
    
    # Ensure storm_year is not None
    if storm_year is None:
        # If storm_year is not available, you may choose to skip this storm or handle it accordingly
        print(f"Storm year not found for {file_name}. Skipping update.")
        return
    
    # Convert data types for consistency
    storm_number = int(storm_number)
    basin = str(basin)
    storm_year = int(storm_year)

    # Check if the storm already exists in the CSV using Basin, Storm Number, and Year
    storm_exists = ((df['Basin'] == basin) & (df['Storm_Number'] == storm_number) & (df['Storm_Year'] == storm_year)).any()
    
    if storm_exists:
        # Update the existing storm information
        df.loc[(df['Basin'] == basin) & (df['Storm_Number'] == storm_number) & (df['Storm_Year'] == storm_year),
               ['Storm_Name', 'Storm_Start_Date', 'Storm_End_Date']] = [storm_name, storm_start_date, storm_end_date]
        print(f"Updated existing storm data for {storm_name} ({basin}{storm_number}, {storm_year}) in {csv_path}.")
    else:
        # Append the new storm information
        new_row = {
            'Basin': basin,
            'Storm_Number': storm_number,
            'Storm_Name': storm_name,
            'Storm_Start_Date': storm_start_date,
            'Storm_End_Date': storm_end_date,
            'Storm_Year': storm_year,
            'adeck_path':file_name,
            'adeck_path_parquet': './forecast_data_parquet/' + file_name.replace('.dat','.parquet')
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        print(f"Added new storm data for {storm_name} ({basin}{storm_number}, {storm_year}) to {csv_path}.")

        print('./forecast_data/' + file_name)
    
    # Save the updated CSV
    df.to_csv(csv_path, index=False)
    print(f"Storm data has been updated in {csv_path}.")

def get_remote_last_modified(url):
    """Get the last modified date of the remote file using a HEAD request."""
    try:

        response = requests.head(url)
        if response.status_code == 200 and 'Last-Modified' in response.headers:
            remote_last_modified = parsedate_to_datetime(response.headers['Last-Modified'])
            #If its not in UTC time, make it so. 
            if remote_last_modified.tzinfo is not None:
                remote_last_modified = remote_last_modified.astimezone(timezone.utc)
            return remote_last_modified
        else:
            return None
        
    except requests.exceptions.RequestException as e:
        print(f"Failed to get the last modified date for {url}: {e}")
        return None

def get_local_last_modified(file_path):
    """Get the last modified date of a local file."""
    if os.path.exists(file_path):
        last_modified_timestamp = os.path.getmtime(file_path)
        # Convert timestamp to datetime in UTC
        last_modified_timestamp = datetime.fromtimestamp(last_modified_timestamp,tz=timezone.utc) 
        last_modified_timestamp = last_modified_timestamp.replace(tzinfo=timezone.utc)

        return last_modified_timestamp
    else:
        return None

def download_file(file_url, save_path):
    """Download a file from a URL and save it to a local path."""
    try:
        response = requests.get(file_url)
        response.raise_for_status()  # Check if the request was successful
        
        try:
            with open(save_path, 'wb') as f:
                f.write(response.content)
                f.flush()  # Flush internal buffer
                os.fsync(f.fileno())  # Ensure file is fully written to disk
        except:
            time.sleep(3)
            with open(save_path, 'wb') as f:
                f.write(response.content)
                f.flush()  # Flush internal buffer
                os.fsync(f.fileno())  # Ensure file is fully written to disk

        print(f"Downloaded: {save_path}")
    except requests.exceptions.RequestException as e:
        print(f"Failed to download {file_url}: {e}")

def check_and_download_file(file_url, save_path):
    """Check if a file should be downloaded (re-download if updated) and download it."""
    remote_last_modified = get_remote_last_modified(file_url)
    local_last_modified = get_local_last_modified(save_path)

    # Log both timestamps for verification
    print(f"Checking file: {save_path}")
    print(f"Remote last modified: {remote_last_modified}")
    print(f"Local last modified: {local_last_modified}")

    # Download if the remote file is newer or if the file doesn't exist locally
    if remote_last_modified and (local_last_modified is None or remote_last_modified > local_last_modified):
        print(f"Updating file {save_path}. Remote file is newer.")
        download_file(file_url, save_path)
        # Update storm_adeck_directory.csv
        download_and_update_storm_data(file_url, save_path, csv_path)
        
        # Trigger conversion to .parquet after download
        convert_dat_to_parquet(save_path)
    else:
        print(f"File {save_path} is up-to-date. Skipping download.")

        # Check if the .parquet file is up-to-date with the .dat file
        base_name = os.path.basename(save_path).replace('.dat', '')
        parquet_file_path = os.path.join(script_dir, "forecast_data_parquet", f'{base_name}.parquet')
        dat_last_modified = get_local_last_modified(save_path)
        parquet_last_modified = get_local_last_modified(parquet_file_path)

        # If .parquet is missing or outdated, update it
        if parquet_last_modified is None or dat_last_modified > parquet_last_modified:
            print(f"Converting .dat to .parquet as .parquet file is outdated or missing for {base_name}.")
            convert_dat_to_parquet(save_path)

def convert_dat_to_parquet(dat_file_path):
    """Convert a .dat file to a .parquet file with retry and error logging."""
    try:
        # Extract the base name (without .dat) for the parquet file
        base_name = os.path.basename(dat_file_path).replace('.dat', '')
        parquet_file_path = os.path.join("./forecast_data_parquet", f'{base_name}.parquet')

        # Convert the .dat file
        gdf = f.read_adeck_dat_file_to_gdf(dat_file_path)
        #gdf= f.format_adeck_parquet(gdf)
        gdf.to_parquet(parquet_file_path)
        print(f"Converted {dat_file_path} to {parquet_file_path}")
        
    except Exception as e:
        # Log any conversion errors for troubleshooting
        print(f"Failed to convert {dat_file_path} to parquet: {e}")
        with open(os.path.join(script_dir, 'conversion_errors.log'), 'a') as log_file:
            log_file.write(f"{datetime.now()}: Failed to convert {dat_file_path}. Error: {e}\n")


def get_dat_files(url):
    """Find and download all .dat files from the given URL."""
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Failed to access {url}: {e}")
        return

    # Parse the page
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find all links on the page
    for link in soup.find_all('a'):
        href = link.get('href')

        # Skip any parent directory links
        if href == '../' or href == '/':
            continue

        # Build full URL for each href
        full_url = urljoin(url, href)

        # If the href is a .dat file, check if it needs to be downloaded or updated
        if href.endswith('.dat'):
            # Save all files directly into the forecast_data folder, without subfolders
            file_name = os.path.basename(href)  # Extract the file name from the URL
            file_path = os.path.join(download_folder, file_name)
            check_and_download_file(full_url, file_path)


# Loop through years and download .dat files for each year
for year in range(start_year, current_year+1):
    
    year_url = f"{base_url}{year}/"
    
    if year==current_year:
        year_url = f"{base_url}/"
    
    print(f"Checking and downloading .dat files for year: {year}")
    get_dat_files(year_url)

"""
Need to now convert all the files that are in the .dat folder to parquet, for app speed
"""

#Get all the file names in the dat folder
dat_files=os.listdir(os.path.join(script_dir, "forecast_data"))
dat_names= [f.replace('.dat', '') for f in dat_files]

#now the same for the parquet files
parquet_files=os.listdir(os.path.join(script_dir,"forecast_data_parquet"))
parquet_names= [f.replace('.parquet', '') for f in parquet_files]

# Find .dat files that don't have corresponding .parquet files
files_to_process = [f for f in dat_names if f not in parquet_names]

for Storm_Name in files_to_process:
    gdf=f.read_adeck_dat_file_to_gdf(os.path.join(script_dir, f'forecast_data/{Storm_Name}.dat'))
    print(gdf.columns)
    try:
        gdf.to_parquet(os.path.join(script_dir,f'./forecast_data_parquet/{Storm_Name}.parquet'))
    except:
        print(f'Skipping storm. Storm Dat type {type(gdf)}')

