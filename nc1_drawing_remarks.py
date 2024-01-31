# Test program for editting NC1 files. In order to work this file must be deposited into job subfolder.
import os
import sys
import shutil

# Opens target kss file and stores information by line as lists, it then packages these lists into a larger list
def ref_kss(kss_path):
    kss_line_data = []
    try:
        with open(kss_path, 'r') as kss_file:
            kss_data = kss_file.readlines()

            for kss_line in kss_data:
                kss_lines = [kss_line.split(',')]
                kss_line_data = kss_line_data+kss_lines

        return kss_line_data
    except FileNotFoundError:
        print(f"{kss_path} NOT FOUND.")
        answer = input("Would you like to use the kss file located in the KSS folder? (yes/no): ")
        if answer.lower() == "yes":
            kss_dir = os.path.dirname(kss_path)
            if len(os.listdir(kss_dir)) > 1:
                print("More than one .kss file found in directory. Please remove any unwanted .kss files and try again.")
                input("Press any key to exit.")
                exit()
            for file in os.listdir(kss_dir):
                if file.endswith(".kss"):
                    shutil.copy(os.path.join(kss_dir, file), kss_path)
                    print(f"Created 'kss combined.kss' in {kss_dir} from existing kss file. Run program again.")
                    input("Press any key to exit.")
                    exit()
        else:
            print("No .kss file copied.")

# Opens target nc1 file and replaces line 5 value with a value cross referenced from kss file
def edit_nc1(nc1_path, kss_line_data):
    with open(nc1_path, 'r+', encoding='utf-8') as nc1_file:
        nc1_data = nc1_file.readlines()

    with open(nc1_path, 'w', encoding='utf-8') as nc1_file:
        try:
            nc1_filename = os.path.basename(nc1_path)
            nc1_lookup_value = os.path.splitext(nc1_filename)[0]
            kss_return = 0

            for kss_list in kss_line_data:  # Looks through lists created from kss file line data and returns index position of nc1 lookup value
                if kss_return >= len(kss_line_data)-1:
                    continue
                try:
                    kss_list.index(nc1_lookup_value)
                    # print(kss_return)
                    break
                except ValueError:
                    kss_return = kss_return + 1

        except  ValueError:
            nc1_file.writelines(nc1_data)
            print("ERROR LOOKING UP PART.")

        nc1_replace = str(nc1_lookup_value+"_DR"+str(kss_line_data[kss_return]).split(',')[1].replace("'","")).replace(" ","") # searches kss for nc1 file name 
        # and returns the appropriate data. Then removes all quotation and spacing in string.
        try:
            nc1_data[4] = "  "+(nc1_replace)+"\n" # replaces data on line 5 with value found in kss file
            nc1_file.writelines(nc1_data) # writes new data to old nc1
        except IndexError:
            nc1_file.writelines(nc1_data)
            return

def get_working_directory():
    pth = os.getcwd()
    return pth

def main():
    # Get the current script directory
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)  # Use this if running as a PyInstaller bundle
    else:
        script_dir = os.path.dirname(os.path.realpath(__file__))  # Use this if running as a normal script

    # Change the current working directory to the script directory
    os.chdir(script_dir)

    nc1_path = os.path.join(script_dir, 'CNC')  # Folder containing nc1 files in current directory.
    kss_path = os.path.join(script_dir, 'KSS/kss combined.kss')

    kss_line_data = ref_kss(kss_path)

    for nc1_file in os.listdir(nc1_path):
        if any(nc1_file.lower().endswith(ext) for ext in ['.nc1']):
            nc1_filepath = os.path.join(nc1_path, nc1_file)
            edit_nc1(nc1_filepath, kss_line_data) 
            print("EDITED: "+nc1_file)

    print("EDITING COMPLETE.")
    input("Press any key to exit.")
    exit()

if __name__ == "__main__":
    main()