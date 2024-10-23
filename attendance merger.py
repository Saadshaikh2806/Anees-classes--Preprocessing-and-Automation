import csv
from collections import defaultdict
import os
import sys
import re
import openpyxl

def get_file_path(prompt):
    print(prompt)
    print("Please drag and drop the file into this window and press Enter.")
    
    file_path = input().strip()
    
    # Remove quotes if present
    file_path = file_path.strip("\"'")
    
    # Handle Windows-style paths with spaces
    if sys.platform.startswith('win'):
        # Remove any leading "& " that Windows might add
        file_path = re.sub(r'^& ', '', file_path)
        
        # Handle cases where the path is duplicated
        match = re.match(r"([A-Za-z]:\\.*?)'([A-Za-z]:\\.*)", file_path)
        if match:
            file_path = match.group(2)
        
        # Remove any remaining single quotes
        file_path = file_path.replace("'", "")
        
        # If the path starts with a drive letter followed by a colon, it's likely a full path
        if re.match(r'^[a-zA-Z]:', file_path):
            return file_path
        else:
            # If it's not a full path, prepend the current working directory
            return os.path.join(os.getcwd(), file_path)
    else:
        return os.path.abspath(file_path)

def read_attendance_data(file_path):
    attendance_data = {}
    if file_path.lower().endswith('.xlsx'):
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        for _ in range(4):  # Skip the first 4 rows
            next(rows)
        headers = next(rows)
        date_columns = headers[2:-4]  # Get all date columns
        for row in rows:
            if row and row[1]:  # Check if the row is not empty and has a name
                name = row[1]
                attendance_data[name] = {
                    'Roll No.': row[0],
                    'Attendance': {date: status for date, status in zip(date_columns, row[2:-4])},
                    'Total Lectures': row[-4],
                    'Presents': row[-3],
                    'Absents': row[-2],
                    'Percent': row[-1]
                }
    else:  # Assume CSV if not XLSX
        with open(file_path, 'r') as file:
            csv_reader = csv.reader(file)
            for _ in range(4):  # Skip the first 4 rows
                next(csv_reader)
            headers = next(csv_reader)
            date_columns = headers[2:-4]  # Get all date columns
            for row in csv_reader:
                if row and row[1]:  # Check if the row is not empty and has a name
                    name = row[1]
                    attendance_data[name] = {
                        'Roll No.': row[0],  # Ensure we're capturing the roll number
                        'Attendance': {date: status for date, status in zip(date_columns, row[2:-4])},
                        'Total Lectures': row[-4],
                        'Presents': row[-3],
                        'Absents': row[-2],
                        'Percent': row[-1]
                    }
    return attendance_data

def read_contact_info(file_path):
    contact_info = {}
    if file_path.lower().endswith('.xlsx'):
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            name = row[headers.index('Name')]
            contact_info[name] = {
                'Mother/Guardian Contact No.': row[headers.index('Mother/Guardian Contact No.')],
                'Father/Guardian Contact No.': row[headers.index('Father/Guardian Contact No.')],
                'Student Contact No.': row[headers.index('Student Contact No.')]
            }
    else:  # Assume CSV if not XLSX
        with open(file_path, 'r') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                name = row['Name']
                contact_info[name] = {
                    'Mother/Guardian Contact No.': row['Mother/Guardian Contact No.'],
                    'Father/Guardian Contact No.': row['Father/Guardian Contact No.'],
                    'Student Contact No.': row['Student Contact No.']
                }
    return contact_info

def merge_data(attendance_data, contact_info):
    merged_data = defaultdict(dict)
    for name, data in attendance_data.items():
        merged_data[name].update(data)
        if name in contact_info:
            merged_data[name].update(contact_info[name])
    return merged_data

def write_merged_data(merged_data, output_file):
    if output_file.lower().endswith('.xlsx'):
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Filter out None values and convert to string
        all_dates = sorted(set(str(date) for data in merged_data.values()
                               for date in data.get('Attendance', {}).keys()
                               if date is not None))

        headers = ['Roll No.', 'Name', 'Mother/Guardian Contact No.', 'Father/Guardian Contact No.',
                   'Student Contact No.'] + all_dates + ['Total Lectures', 'Presents', 'Absents', 'Percent']

        sheet.append(headers)

        for name, data in merged_data.items():
            row = [
                data.get('Roll No.', ''),
                name,
                data.get('Mother/Guardian Contact No.', ''),
                data.get('Father/Guardian Contact No.', ''),
                data.get('Student Contact No.', '')
            ]
            
            for date in all_dates:
                row.append(data.get('Attendance', {}).get(date, ''))
            
            row.extend([
                data.get('Total Lectures', ''),
                data.get('Presents', ''),
                data.get('Absents', ''),
                data.get('Percent', '')
            ])
            
            sheet.append(row)

        workbook.save(output_file)
    else:  # Write as CSV
        with open(output_file, 'w', newline='') as file:
            writer = csv.writer(file)
            
            # Filter out None values and convert to string
            all_dates = sorted(set(str(date) for data in merged_data.values()
                                   for date in data.get('Attendance', {}).keys()
                                   if date is not None))
            
            headers = ['Roll No.', 'Name', 'Mother/Guardian Contact No.', 'Father/Guardian Contact No.',
                       'Student Contact No.'] + all_dates + ['Total Lectures', 'Presents', 'Absents', 'Percent']
            writer.writerow(headers)
            
            for name, data in merged_data.items():
                row = [
                    data.get('Roll No.', ''),
                    name,
                    data.get('Mother/Guardian Contact No.', ''),
                    data.get('Father/Guardian Contact No.', ''),
                    data.get('Student Contact No.', '')
                ]
                
                for date in all_dates:
                    row.append(data.get('Attendance', {}).get(date, ''))
                
                row.extend([
                    data.get('Total Lectures', ''),
                    data.get('Presents', ''),
                    data.get('Absents', ''),
                    data.get('Percent', '')
                ])
                
                writer.writerow(row)

    print(f"Merged data has been written to {output_file}")

def main():
    print("Welcome to the Attendance Merger!")
    
    attendance_file = get_file_path("Please drag and drop the attendance file (CSV or XLSX):")
    contact_file = get_file_path("Please drag and drop the contact info file (CSV or XLSX):")
    
    output_file = 'ATTENDANCE_MERGER.xlsx'
    
    print("\nProcessing files...")
    print(f"Attendance file: {attendance_file}")
    print(f"Contact info file: {contact_file}")
    
    attendance_data = read_attendance_data(attendance_file)
    contact_info = read_contact_info(contact_file)
    merged_data = merge_data(attendance_data, contact_info)
    write_merged_data(merged_data, output_file)
    
    print(f"\nMerged data has been written to {output_file}")
    print("The merged file is in the same directory as this script.")

if __name__ == "__main__":
    main()
