import os
import shutil
import zipfile
from concurrent.futures import ThreadPoolExecutor
from openpyxl import load_workbook
from tqdm import tqdm
from pathlib import Path

subfolders_list = ['CNC','DRAWINGS','EXCEL FILES', 'KSS','SHIPPING AND BILLING','ZIP FILES', 'temp', 'ARCHIVE']

def create_folders(job_path: Path, subfolders_list):
    for subfolder in subfolders_list:
        (job_path / subfolder).mkdir(parents=True, exist_ok=True)

def extract_zip_archive_with_progress(archived_path: Path, job_path: Path, pbar):
    temp_folder = job_path / 'temp'
    temp_folder.mkdir(parents=True, exist_ok=True)
    total_size = archived_path.stat().st_size
    extracted_size = 0

    def update_progress(extracted_size):
        progress = (extracted_size/ total_size) * 100
        pbar.update(progress - pbar.n)

    with zipfile.ZipFile(archived_path, 'r') as zip_ref:
        with ThreadPoolExecutor(max_workers = 200) as exe:
            for member in zip_ref.infolist():
                    exe.submit(zip_ref.extract, member, temp_folder)
                    extracted_size += member.file_size
                    update_progress(extracted_size)

def move_files(source: Path, destination: Path, extensions):
    for file in source.iterdir():
        if file.suffix.lower() in extensions:
            shutil.move(str(file), str(destination / file.name))

def index_temp(job_path: Path, temp_path: Path):
    extension_to_folder = {
        '.nc1': 'CNC',
        '.cnc': 'CNC',
        '.step': 'CNC',
        '.stp': 'CNC',
        '.dxf': 'CNC',
        '.pdf': 'DRAWINGS',
        '.zip': 'ZIP FILES',
        '.rar': 'ZIP FILES',
        'master': 'SHIPPING AND BILLING',
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
        '.cnc': 'CNC',
        '.step': 'CNC',
        '.stp': 'CNC',
        '.dxf': 'CNC',
        '.pdf': 'DRAWINGS',
        '.zip': 'ZIP FILES',
        '.rar': 'ZIP FILES',
        'master': 'SHIPPING AND BILLING',
        '.xlsx': 'EXCEL FILES',
        '.xlsm': 'EXCEL FILES',
        '.xls': 'EXCEL FILES',
        '.kss': 'KSS',
    }

    for file in job_path.iterdir():
        if file.is_file() and file.suffix.lower() in ['.zip', '.rar']:
            print(f"Extracting '{file.name}'...")
            with tqdm(total=100, unit="B", unit_scale=True) as pbar:
                extract_zip_archive_with_progress(file, job_path, pbar)

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
                    move_files(folder, job_path / extension_to_folder[file.suffix.lower()], [file.suffix.lower()])
            shutil.move(str(folder), str(job_path / 'temp'))
    
    print(f"Copying 'nc1_drawing_remarks.exe' to '{job_path}'...")
    copy_exe_to_job_folder(job_path)

def main(subfolders_list):
    spreadsheet_path = Path('data/workschedule.xlsx')
    workbook = load_workbook(spreadsheet_path)
    base_directory = Path('Y:/02 Job Files')

    client_column = 0
    job_column = 1

    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, values_only=True):
        client_name = row[client_column] 
        job_name = row[job_column] or "Reserve"

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

if __name__ == "__main__":
    main(subfolders_list)