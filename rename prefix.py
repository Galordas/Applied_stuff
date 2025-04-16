import os

# Define the directory where the files are located
directory = r''  # Replace with the actual path to your directory

# Define the prefix you want to remove from the file names
prefix = "140621_"

try:
    # Iterate through all the files in the given directory
    for filename in os.listdir(directory):
        if filename.startswith(prefix):
            # Construct the new file name by removing the prefix from the original file name
            new_filename = filename[len(prefix):]

            # Construct the full original file path and the new file path
            original_file_path = os.path.join(directory, filename)
            new_file_path = os.path.join(directory, new_filename)

            if os.path.exists(new_file_path):
                # If a file already exists with the new name, handle that case specifically
                print(f"File '{new_filename}' already exists. Skipping renaming for '{filename}'.")
            else:
                # Rename the file from the original path to the new path
                os.rename(original_file_path, new_file_path)
                print(f"Renamed '{filename}' to '{new_filename}'")
except FileNotFoundError:
    print(f"The directory '{directory}' does not exist.")
except Exception as e:
    print(f"An error occurred: {e}")