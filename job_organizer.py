#! python3
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import shutil
import zipfile
import datetime
from concurrent.futures import ThreadPoolExecutor
from openpyxl import load_workbook
from tqdm import tqdm
from pathlib import Path
from pyunpack import Archive
import shutil

subfolders_list = ['CNC','DRAWINGS','EXCEL FILES', 'KSS','SHIPPING AND BILLING','ZIP FILES', 'temp', 'ARCHIVE', 'MATERIAL']

def create_folders(job_path: Path, subfolders_list):
    for subfolder in subfolders_list:
        (job_path / subfolder).mkdir(parents=True, exist_ok=True)

def extract_archive_with_progress(archived_path: Path, job_path: Path, pbar):
    temp_folder = job_path / 'temp'
    temp_folder.mkdir(parents=True, exist_ok=True)
    total_size = archived_path.stat().st_size
    extracted_size = 0

    def update_progress(extracted_size):
        progress = (extracted_size/ total_size) * 100
        pbar.update(progress - pbar.n)

    if archived_path.suffix.lower() == '.rar':
        # Move .rar file to ARCHIVE subfolder
        archive_folder = job_path / 'ARCHIVE'
        archive_folder.mkdir(parents=True, exist_ok=True)
        shutil.move(str(archived_path), str(archive_folder / archived_path.name))
    else:

        with zipfile.ZipFile(archived_path, 'r') as zip_ref:
            with ThreadPoolExecutor(max_workers = 200) as exe:
                for member in zip_ref.infolist():
                        exe.submit(zip_ref.extract, member, temp_folder)
                        extracted_size += member.file_size
                        update_progress(extracted_size)

def move_files(source: Path, destination: Path, extensions):
    for file in source.iterdir():
        if file.suffix.lower() in extensions:
            try:
                shutil.move(str(file), str(destination / file.name))
            except PermissionError:
                print(f"Permission denied when moving '{file.name}'. The file might be in use or inaccessible.")
                
def index_temp(job_path: Path, temp_path: Path):
    extension_to_folder = {
        '.nc1': 'CNC',
        '.nc': 'CNC',
        '.cnc': 'CNC',
        '.step': 'CNC',
        '.stp': 'CNC',
        '.dxf': 'CNC',
        '.pdf': 'DRAWINGS',
        '.zip': 'ZIP FILES',
        '.rar': 'ZIP FILES',
        'master': 'SHIPPING AND BILLING',
        'billing': 'SHIPPING AND BILLING',
        'material': 'MATERIAL',
        '.xlsx': 'EXCEL FILES',
        '.xlsm': 'EXCEL FILES',
        '.xls': 'EXCEL FILES',
        '.kss': 'KSS',
    }

    for file in temp_path.rglob('*'):
        if file.is_file() and file.suffix.lower() in extension_to_folder:
            print(f"Moving '{file.suffix.lower()}' files to '{extension_to_folder[file.suffix.lower()]}'...")
            move_files(file.parents[0], job_path / extension_to_folder[file.suffix.lower()], [file.suffix.lower()])

    (job_path / 'ARCHIVE').mkdir(parents=True, exist_ok=True)
    shutil.copytree(temp_path, job_path / 'ARCHIVE', dirs_exist_ok=True)
    shutil.rmtree(temp_path)

def copy_exe_to_job_folder(job_path: Path):
    exe_source = Path('dist/nc1_drawing_remarks.exe')
    if exe_source.exists():
        shutil.copy(exe_source, job_path)
    else:
        print(f"Executable file '{exe_source}' not found.")

def organize_job(job_path: Path, subfolders_list):
    extension_to_folder = {
        '.nc1': 'CNC',
        '.nc': 'CNC',
        '.cnc': 'CNC',
        '.step': 'CNC',
        '.stp': 'CNC',
        '.dxf': 'CNC',
        '.pdf': 'DRAWINGS',
        '.zip': 'ZIP FILES',
        '.rar': 'ZIP FILES',
        'master': 'SHIPPING AND BILLING',
        'billing': 'SHIPPING AND BILLING',
        'material': 'MATERIAL',
        '.xlsx': 'EXCEL FILES',
        '.xlsm': 'EXCEL FILES',
        '.xls': 'EXCEL FILES',
        '.kss': 'KSS',
    }

    for file in job_path.iterdir():
        if file.is_file() and file.suffix.lower() in ['.zip', '.rar']:
            print(f"Extracting '{file.name}'...")
            with tqdm(total=100, unit="B", unit_scale=True) as pbar:
                extract_archive_with_progress(file, job_path, pbar)

    index_temp(job_path, job_path / 'temp')

    for file in job_path.iterdir():
        if file.suffix.lower() in extension_to_folder:
            print(f"Moving '{file.suffix.lower()}' files to '{extension_to_folder[file.suffix.lower()]}'...")
            move_files(job_path, job_path / extension_to_folder[file.suffix.lower()], [file.suffix.lower()])

    for folder in job_path.iterdir():
        if folder.is_dir() and folder.name not in subfolders_list:
            for file in folder.iterdir():
                if file.suffix.lower() in extension_to_folder:
                    print(f"Moving '{file.suffix.lower()}' files to '{extension_to_folder[file.suffix.lower()]}'...")
                    try:
                        move_files(folder, job_path / extension_to_folder[file.suffix.lower()], [file.suffix.lower()])
                    except (PermissionError, OSError) as e:
                        print(f"Error moving file: {e}")
            try:           
                shutil.move(str(folder), str(job_path / 'temp'))
            except (PermissionError, OSError) as e:
                print(f"Error moving folder: {e}")
    
    print(f"Copying 'nc1_drawing_remarks.exe' to '{job_path}'...")
    copy_exe_to_job_folder(job_path)

def copy_material_spreadsheet(job_path: Path):
    material_spreadsheet = Path('data/Material_Takeoff.xlsm')
    if material_spreadsheet.exists():
        try:
            print(f"Copying 'Material_Takeoff.xlsm' to '{job_path}'...")
            shutil.copy(material_spreadsheet, job_path.joinpath('MATERIAL'))
        except PermissionError:
            print(f"Permission denied when copying 'Material_Takeoff.xlsm'. The file might be in use or inaccessible.")
    else:
        print(f"Material spreadsheet '{material_spreadsheet}' not found.")

def get_google_sheet_jobs(sheet):
    # Get all values from the sheet
    rows = sheet.get_all_values()
    job_list = set()

    client_column = 0  # Change this to the column number where the client name is located
    job_column = 1  # Change this to the column number where the job name is located

    for row in rows[1:]:  # Skip the header row
        client_name = row[client_column].strip()
        job_name = row[job_column].strip() or "Reserve"
        for char in job_name:
            if char in "?.!/;:":
                job_name.replace(char,'')
        if client_name and job_name:
            job_list.add(f"{client_name}/{job_name}")

    return job_list


def report_unmatched_folders(base_directory, google_sheet_jobs):
    # Get all client/job folders from the base directory
    all_folders = set()
    for client_folder in base_directory.iterdir():
        if client_folder.is_dir():
            for job_folder in client_folder.iterdir():
                if job_folder.is_dir():
                    all_folders.add(f"{client_folder.name}/{job_folder.name}")

    # Find unmatched folders
    unmatched_folders = all_folders - google_sheet_jobs

    # Print the report
    if unmatched_folders:
        print("Unmatched folders:")
        for folder in unmatched_folders:
            print(folder)
            # Rename folder to add an asterisk
            client_name, job_name = folder.split('/')
            job_folder_path = base_directory / client_name / job_name
            new_job_folder_path = base_directory / client_name / (job_name + '+++')
            if not new_job_folder_path.exists():  # Ensure no existing folder with asterisk
                job_folder_path.rename(new_job_folder_path)
            else:
                print(f"Folder {new_job_folder_path} already exists. Skipping renaming.")
    else:
        print("All folders match the Google Sheet entries.")

def check_empty_folders(base_directory):
    empty_folders = []
    for client_folder in base_directory.iterdir():
        if client_folder.is_dir():
            for job_folder in client_folder.iterdir():
                if job_folder.is_dir():
                    drawings_folder = job_folder / 'DRAWINGS'
                    kss_folder = job_folder / 'KSS'
                    zip_folder = job_folder / 'ZIP FILES'
                    if (drawings_folder.is_dir() and not any(drawings_folder.iterdir()) and
                        kss_folder.is_dir() and not any(kss_folder.iterdir()) and
                        zip_folder.is_dir() and not any(zip_folder.iterdir())):
                        empty_folders.append(job_folder)

     # Generate and save the report
    report_path = Path(r'\\serverdfi\e\users\bobby\2024\14 PROJECTS\Job Organizer Project\data\empty_folders_report.txt')
    with report_path.open('w') as report_file:
        if empty_folders:
            report_file.write("Job folders with empty DRAWINGS, KSS, and ZIP FILES:\n")
            for folder in empty_folders:
                report_file.write(f"{folder}\n")
        else:
            report_file.write("No job folders have all three DRAWINGS, KSS, and ZIP FILES empty.\n")

    # Print the report
    if empty_folders:
        print("Job folders with empty DRAWINGS, KSS, and ZIP FILES:")
        for folder in empty_folders:
            print(folder)
    else:
        print("No job folders have all three DRAWINGS, KSS, and ZIP FILES empty.")

def main(subfolders_list):
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secrets.json', scope)
    client = gspread.authorize(creds)

    # Open the Google Sheets document
    sheet = client.open("SCHEDULE BOARD").sheet1

    base_directory = Path('Y:/02 Job Files')  # Change this to the base directory where the client folders are located


    # Fetch job list from Google Sheet
    #google_sheet_jobs = get_google_sheet_jobs(sheet)

    # Report unmatched folders
    #report_unmatched_folders(base_directory, google_sheet_jobs)
    
    client_column = 0   # Change this to the column number where the client name is located
    job_column = 1  # Change this to the column number where the job name is located
    due_date_column = 3 # Change this to the column number where the due date is located
    bill_status_colulmn = 11 # Change this to the column number where the bill status is located

    # Get the current date and the date for the start of the week
    now = datetime.datetime.now()
    start_of_week = now - datetime.timedelta(days=now.weekday())

    # Get all values from the sheet
    rows = sheet.get_all_values()

    for row in rows[1:]:  # Skip the header row
        try:
            due_date = datetime.datetime.strptime(row[due_date_column], '%m/%d/%Y')  or datetime.datetime.now()
        except ValueError:
            due_date = datetime.datetime.now()
        client_name = row[client_column].strip() 
        job_name = row[job_column].strip() or "Reserve"
        bill_status = row[bill_status_colulmn].strip() or " "

        if bill_status.lower() == "billed":
            continue

        if due_date <= start_of_week:
            continue

        if client_name is None:
            break

        client_folder = base_directory / client_name.strip()
        job_folder = client_folder / job_name

        if not client_folder.exists():
            response = input(f"Do you want to create a new folder for client '{client_name}'? (yes/no): ")
            if response.lower() == 'yes':
                client_folder.mkdir(parents=True, exist_ok=True)
            else:
                existing_client = input("Enter the name of an existing client: ")
                client_folder = base_directory / existing_client
                job_folder = client_folder / job_name

        print(f"Creating job folder '{job_folder}'...")
        create_folders(job_folder, subfolders_list)
        print(f"Organizing job '{job_folder}'...")
        organize_job(job_folder, subfolders_list)
        print(f"Copying material spreadsheet to '{job_folder}'...")
        copy_material_spreadsheet(job_folder)

    # Check and report empty folders
    check_empty_folders(base_directory)

    print("Done.")
    input("Press Enter to exit...")

if __name__ == "__main__":
    main(subfolders_list)