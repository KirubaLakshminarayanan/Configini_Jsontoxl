import os
import json
import configparser
from openpyxl import Workbook
from datetime import datetime
import logging
import pytz
import time  # Import the time module for adding delay
import tqdm  # Import tqdm for displaying progress bar

# Create a logs directory if it doesn't exist
logs_dir = 'logs'
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)


# Function to configure logging with a custom log file name
def configure_logging(log_filename):
    logging.basicConfig(filename=log_filename, level=logging.DEBUG,
                        format='%(asctime)s - %(levelname)s - %(message)s')


# Function to flatten JSON data
def flatten_json(data, parent_key='', sep='_'):
    items = {}
    if isinstance(data, dict):
        for k, v in data.items():
            new_key = parent_key + sep + k if parent_key else k
            items.update(flatten_json(v, new_key, sep=sep))
    elif isinstance(data, list):
        for i, v in enumerate(data):
            new_key = parent_key + sep + str(i)
            items.update(flatten_json(v, new_key, sep=sep))
    elif data is None:
        items[parent_key] = None
    elif isinstance(data, (str, int, float, bool)) and parent_key:
        items[parent_key] = data
    return items


# Function to convert JSON data to Excel
def convert_to_excel(file_paths, output_dir, sheet_name):
    # Logging Configuration
    log_filename = os.path.join(logs_dir, f"errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    configure_logging(log_filename)  # Configure logging with a new log file name
    # Initialization
    total_files = len(file_paths)
    success_count = 0  # Track the number of successful conversions

    # Get the Indian Standard Timezone
    ist_timezone = pytz.timezone('Asia/Kolkata')
# Processing each file
    for file_path in tqdm.tqdm(file_paths, desc="Converting JSON to Excel"):
        # Add a delay to simulate processing time
        time.sleep(0.1)  # Adjust the delay time as needed
# File Format and Size Validation
        if not file_path.endswith('.json'):
            logging.error(f"Skipping {file_path}: Invalid file format. Please provide a JSON file.")
            continue

        if os.path.getsize(file_path) == 0:
            logging.error(f"Skipping {file_path}: File is empty.")
            continue
# Reading Json File
        try:
            with open(file_path, 'r') as json_file:
                try:
                    data = json.load(json_file)
                except ValueError as e:
                    error_message = f"Invalid JSON format in {file_path}: {str(e)}"
                    logging.error(error_message)
                    continue
# Checking Json Format. Check Json data is dict or list.If it's dict convert to list and if it's list use as it is.
                if isinstance(data, dict):
                    records = [data]
                elif isinstance(data, list):
                    records = data
                else:
                    logging.error(f"Skipping {file_path}: Invalid JSON data format.")
                    continue
# Checking For empty records
                if not records:
                    logging.error(f"No records found in {file_path}.")
                    continue
# Flattening Json Records
                flattened_records = [flatten_json(record) for record in records]

                # Get the headers in the order of keys from the first record
                headers = list(flattened_records[0].keys())
# Creating O/P File name and Excel workbook
                output_file = os.path.join(output_dir,
                                           os.path.basename(file_path)[:-5] + '_' + datetime.now(ist_timezone).strftime(
                                               "%Y%m%d%H%M%S") + '.xlsx')  # Convert datetime to Indian Standard Time
                wb = Workbook()
                ws = wb.active
                ws.title = sheet_name
# Writing headers and data to Excel and Save the Excel File.
                ws.append(headers)

                for record in flattened_records:
                    row_data = []
                    for header in headers:
                        row_data.append(
                            record.get(header, ''))  # Get value for each header or empty string if not present
                    ws.append(row_data)

                wb.save(output_file)
                logging.info(f"{file_path} conversion successful. Converted to {output_file}.")
                logging.info(f"Size of converted file: {os.path.getsize(output_file)} bytes")
                success_count += 1  # Increment the success count
        except Exception as e:
            logging.error(f"{file_path} conversion failed! Error: {str(e)}")
# Logging Summary
    failed_files_count = total_files - success_count
    if failed_files_count == 0:
        logging.info(f"Completed {total_files} files. Conversion of JSON to Excel file is successful.")
    else:
        logging.info(
            f"Completed {total_files} files. Conversion of JSON to Excel file failed for {failed_files_count} file(s).")


# Main function to read config and process files
def main():
    config = configparser.ConfigParser()
    config.read('config.ini')

    input_dir = config['Files']['input_dir']
    output_dir = config['Files']['output_dir']
    chunk_size: int = int(config['Files']['chunk'])
    sheet_name = config['Excel']['sheet_name']

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    json_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.endswith('.json')]

    if json_files:
        # Process files in chunks
        for i in range(0, len(json_files), chunk_size):
            chunk_files = json_files[i:i + chunk_size]
            convert_to_excel(chunk_files, output_dir, sheet_name)
    else:
        logging.error("No JSON files found in the input directory.")

# Entry Point
if __name__ == "__main__":
    main()
