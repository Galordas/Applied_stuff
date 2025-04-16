import os


def rename_files(directory):
    # Iterate over all files in the given directory
    for filename in os.listdir(directory):
        # Check if the file is an SVG file
        if filename.endswith(".svg"):
            # Split the filename by the first underscore character
            parts = filename.split('_', 1)  # Split only once

            if len(parts) == 2:
                new_filename = parts[
                    1]  # The part after the first underscore should be kept (including any underscores after the first one)
                old_file_path = os.path.join(directory, filename)
                new_file_path = os.path.join(directory, new_filename)

                # Make sure not to overwrite any existing files unless it's the same file
                if old_file_path != new_file_path:
                    os.rename(old_file_path, new_file_path)
                    print(f"Renamed file '{filename}' to '{new_filename}'")

            else:
                print(f"Filename '{filename}' does not contain an underscore or does not split correctly.")


if __name__ == "__main__":
    directory = r""
    rename_files(directory)
