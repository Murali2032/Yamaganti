import os
import glob

# List of directories to clean
directories = [r"S:\New File\Output", r"S:\Old File", r"S:\New File"]

# Extensions to delete
file_extensions = ['*.csv', '*.xlsx']

for directory in directories:
    if os.path.exists(directory):
        for ext in file_extensions:
            files = glob.glob(os.path.join(directory, ext))  # Find all matching files
            for file in files:
                try:
                    os.remove(file)  # Delete file
                    print(f"Deleted: {file}")
                except Exception as e:
                    print(f"Error deleting {file}: {e}")
    else:
        print(f"Directory does not exist: {directory}")



