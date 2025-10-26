import os
import shutil

def clean_folders(root_dir):
    """
    Recursively clean a folder tree:
    - Keep only files starting with 'CG.' and ending with '.pdf'
    - Ignore .zip files (do not delete them)
    - Delete all other files
    - Remove folders that do not contain any 'CG.' pdf files after cleanup
    """
    for folder_path, subdirs, files in os.walk(root_dir, topdown=False):
        keep_folder = False  # Track if this folder contains any valid files

        for file_name in files:
            file_path = os.path.join(folder_path, file_name)

            # Skip .zip files
            if file_name.lower().endswith('.zip'):
                keep_folder = True
                continue

            # Check if file should be kept
            if file_name.startswith('CG.') and file_name.lower().endswith('.pdf'):
                keep_folder = True
                continue

            # Delete any other files
            try:
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

        # After cleaning files, check if folder still contains valid content
        if not keep_folder:
            # Check if folder is now empty (no files or subfolders)
            if not os.listdir(folder_path):
                try:
                    os.rmdir(folder_path)
                    print(f"Deleted empty folder: {folder_path}")
                except Exception as e:
                    print(f"Error deleting folder {folder_path}: {e}")

if __name__ == "__main__":
    # Example usage:
    # Replace this path with the root directory you want to clean
    root_directory = r"/Users/fahad/Desktop/car_stolen/Cars/SKODA/"
    clean_folders(root_directory)
