import subprocess
import os

# Directory containing the Python script files
script_dir = r'C:\Users\ahmed\OneDrive\Desktop\Clapgrow-Salasar_Lombard-Reconciliation'

# List of Python script filenames
scripts = [
    'Pol_no+End_no.py',
    'Customer+Policy+Premium.py',
    'Customer+Premium+Tenure.py'
]


def execute_script(script_path):
    # Execute the Python script
    subprocess.run(['python', script_path], check=True)


if __name__ == "__main__":
    # Change the current working directory to the script directory
    os.chdir(script_dir)

    for script in scripts:
        print(f"Executing {script}...")
        execute_script(script)
        print(f"Finished executing {script}")