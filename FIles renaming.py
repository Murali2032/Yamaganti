import os

# Updated directory path
root_directory = r"C:\Users\YamagantiMuraliKrish\Downloads\November"

def rename_excel_files(directory):
    """
    Traverse the directory and its subfolders to rename Excel files
    such that each word in the file name is capitalized.
    """
    # Walk through the directory and subdirectories
    for foldername, subfolders, filenames in os.walk(directory):
        for filename in filenames:
            # Check if the file has an Excel extension
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                # Split the file name and extension
                name, ext = os.path.splitext(filename)
                # Convert the file name to title case
                new_name = name.title() + ext
                # Construct full paths
                old_file = os.path.join(foldername, filename)
                new_file = os.path.join(foldername, new_name)
                # Rename the file if the name changes
                if old_file != new_file:
                    os.rename(old_file, new_file)
                    print(f"Renamed: {old_file} -> {new_file}")

if __name__ == "__main__":
    # Call the function to rename Excel files
    print("Starting the renaming process...")
    rename_excel_files(root_directory)
    print("Renaming process completed!")