import os
import json
import shutil
import subprocess
from datetime import datetime

# Get the root directory (parent of version_manager)
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

VERSION_FILE = os.path.join(ROOT_DIR, 'version_manager', 'version.json')
LOG_FILE = os.path.join(ROOT_DIR, 'version_manager', 'version_history.log')
BUILD_DIR = os.path.join(ROOT_DIR, 'build')
VERSIONS_DIR = os.path.join(ROOT_DIR, 'versions')
VENV_DIR = os.path.join(ROOT_DIR, 'env')
SETUP_PY = os.path.join(ROOT_DIR, 'version_manager', 'setup.py')

def load_version():
    if os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, 'r') as f:
            return json.load(f)
    return {'major': 2, 'minor': 0, 'patch': 4}  # Starting from the current version

def save_version(version):
    with open(VERSION_FILE, 'w') as f:
        json.dump(version, f)

def update_version(change_type):
    version = load_version()
    
    if change_type == 'major':
        version['major'] += 1
        version['minor'] = 0
        version['patch'] = 0
    elif change_type == 'minor':
        version['minor'] += 1
        version['patch'] = 0
    elif change_type == 'patch':
        version['patch'] += 1
    
    save_version(version)
    log_version_update(version, change_type)
    return f"{version['major']}.{version['minor']}.{version['patch']}"

def log_version_update(version, change_type):
    version_str = f"{version['major']}.{version['minor']}.{version['patch']}"
    log_entry = f"{datetime.now().isoformat()} - Version {version_str} ({change_type} change)\n"
    
    with open(LOG_FILE, 'a') as f:
        f.write(log_entry)

def get_current_version():
    version = load_version()
    return f"{version['major']}.{version['minor']}.{version['patch']}"

def find_exe_folder():
    for item in os.listdir(BUILD_DIR):
        if item.startswith('exe.'):
            return os.path.join(BUILD_DIR, item)
    return None

def copy_json_files(exe_folder):
    json_files = [
        'platform_config.json',
        'lookup_dictionaries.json',
        'rfm_lookup_dictionaries.json'
    ]
    for json_file in json_files:
        src = os.path.join(ROOT_DIR, json_file)
        dst = os.path.join(exe_folder, json_file)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"Copied {json_file} to exe folder")
        else:
            print(f"Warning: {json_file} not found in root directory")

def build_and_zip():
    version = get_current_version()
    
    # Activate virtual environment and run setup.py build
    if os.name == 'nt':  # Windows
        activate_cmd = os.path.join(VENV_DIR, 'Scripts', 'activate.bat')
        build_cmd = f"call {activate_cmd} && python {SETUP_PY} build"
    else:  # Unix-based systems
        activate_cmd = f"source {os.path.join(VENV_DIR, 'bin', 'activate')}"
        build_cmd = f"{activate_cmd} && python {SETUP_PY} build"
    
    try:
        subprocess.run(build_cmd, shell=True, check=True, cwd=ROOT_DIR)
    except subprocess.CalledProcessError as e:
        print(f"Error during build process: {e}")
        return
    
    # Find the exe folder
    exe_folder = find_exe_folder()
    if not exe_folder:
        print("Error: Could not find the exe folder in the build directory")
        return
    
    # Copy JSON files to exe folder
    copy_json_files(exe_folder)
    
    # Create versions directory if it doesn't exist
    if not os.path.exists(VERSIONS_DIR):
        os.makedirs(VERSIONS_DIR)
    
    # Zip the exe folder
    zip_filename = f"ultimate_transaction_compilerV{version}.zip"
    zip_path = os.path.join(VERSIONS_DIR, zip_filename)
    
    shutil.make_archive(os.path.splitext(zip_path)[0], 'zip', exe_folder)
    
    print(f"Build completed and zipped: {zip_path}")

if __name__ == '__main__':
    print(f"Current version: {get_current_version()}")
    change_type = input("Enter change type (major/minor/patch): ").lower()
    new_version = update_version(change_type)
    print(f"Updated version: {new_version}")
    
    build_and_zip()