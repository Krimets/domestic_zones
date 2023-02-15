import os
import requests
from bs4 import BeautifulSoup
import pandas as pd


# Function to download zone files for UPS domestic zip ranges
def download_zone_files(url, zip_ranges_file):
    # Read in zip ranges file
    zip_ranges = pd.read_excel(zip_ranges_file, sheet_name='UPS zip ranges')

    # Loop through each zip range and download corresponding zone file
    for _, row in zip_ranges.iterrows():
        start_zip = row['zip from']
        end_zip = row['zip to']

        # Find zone file download
        r = requests.get(url)
        soup = BeautifulSoup(r.content, 'html.parser')
        form = soup.find('form', {'action': 'https://www.ups.com/zonecharts/', 'method': 'post'})
        r = requests.post(form['action'], data={'zipcode': start_zip})

        filename = f'{start_zip}-{end_zip}.xls'
        with open(filename, 'wb') as f:
            f.write(r.content)

        # Check that the downloaded zone file matches the expected zip range
        if not os.path.exists('./wrong_zone_range'):  # create wrong_zone_range directory if it doesn't exist
            os.mkdir('./wrong_zone_range')

        try:
            df = pd.read_excel(filename, engine='openpyxl', header=None)
            zone_range = df.iat[4, 0]
            print('Try to upload:\n', zone_range)
            new_start_zip = start_zip
            if start_zip > 9999:
                new_start_zip = str(start_zip)[:-2]
                new_end_zip = str(end_zip)[:-2] + '-99'
            elif start_zip > 999:
                new_start_zip = str(start_zip)[:-1]
                new_end_zip = str(end_zip)[:-2] + '0-99'
            else:
                new_end_zip = (str(end_zip)[:-2]) + '00-99'
            expected_range = f'{new_start_zip}-01 to {new_end_zip}'
            print(expected_range)

            if expected_range not in zone_range:
                os.rename(filename, os.path.join('wrong_zone_range', os.path.basename(filename)))
                print(f'File {os.path.basename(filename)} moved to the "wrong_zone_range" directory')
        except BaseException as e:
            print(e)

        convert()


# Convert all xls files to xlsx files
def convert():
    input_dir = os.getcwd()  # current directory
    output_dir = os.path.join(os.getcwd(), 'xlsx')
    xls_files = [f for f in os.listdir(input_dir) if f.endswith('.xls')]

    if len(xls_files) > 10:
        if not os.path.exists(output_dir):  # create .xlsx directory if it doesn't exist
            os.makedirs(output_dir)
        if not os.path.exists('./bad_files'):  # create bad_files directory if it doesn't exist
            os.mkdir('./bad_files')
        for file in os.listdir(input_dir):
            if file.endswith('.xls'):
                try:
                    input_file_path = os.path.join(input_dir, file)
                    output_file_path = os.path.join(output_dir, os.path.splitext(file)[0] + '.xlsx')
                    df = pd.read_excel(input_file_path, engine='openpyxl')  # read file using openpyxl engine
                    df.to_excel(output_file_path, index=False)  # save the converted file
                    os.remove(input_file_path)  # remove file
                    print(file, 'converted to .xlsx')
                except BaseException as e:
                    os.rename(file, os.path.join('bad_files', os.path.basename(file)))
                    print(f'File {os.path.basename(file)} moved to the "bad_files" directory')


# Download domestic zone files for UPS
url = 'https://www.ups.com/us/en/support/shipping-support/shipping-costs-rates/retail-rates.page'
zip_ranges_file = 'Carriers zone ranges.xlsx'
download_zone_files(url, zip_ranges_file)
