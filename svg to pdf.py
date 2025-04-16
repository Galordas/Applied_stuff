import os
import subprocess
import PyPDF2
import re
import tempfile
import shutil


def check_inkscape_path(inkscape_path):
    if not os.path.exists(inkscape_path):
        raise FileNotFoundError(f"Inkscape was not found at the provided path: {inkscape_path}")


def natural_sort_key(s):
    """
    Key function for natural sorting of strings that contain numbers.
    """
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]


def main(svg_directory, output_pdf_filename, inkscape_path):
    # Step 1: List all .svg files in the given directory
    svg_files = [f for f in os.listdir(svg_directory) if f.endswith('.svg')]

    # Ensuring that the files are sorted based on the numeric part of their names ("1.svg", "2.svg", "10.svg", "22.svg", "201.svg")
    svg_files.sort(key=natural_sort_key)

    # Temporary directory to store the temporary .pdf files
    temp_dir = tempfile.mkdtemp()

    pdf_files = []

    try:
        # Step 2: Iterate through sorted .svg files and convert them to .pdf using Inkscape
        for svg_file in svg_files:
            base_name = os.path.splitext(svg_file)[0]  # Get the part without the extension (e.g., "1" from "1.svg")
            svg_path = os.path.join(svg_directory, svg_file)
            pdf_path = os.path.join(temp_dir, f"{base_name}.pdf")

            # Running Inkscape command to convert .svg files to .pdf
            subprocess.run([inkscape_path, "--export-filename", pdf_path, svg_path], check=True)

            # Append pdf path to the list for later merging
            pdf_files.append(pdf_path)

        # Step 3: Merging temporary .pdf files into a single .pdf file using PyPDF2
        pdf_merger = PyPDF2.PdfMerger()
        for pdf_file in pdf_files:
            pdf_merger.append(pdf_file)

        # Step 4: Write the merged .pdf file to the specified output file path
        pdf_merger.write(output_pdf_filename)
        pdf_merger.close()

    finally:
        # Clean up the temporary directory if needed
        shutil.rmtree(temp_dir)

    print(f"Merged all .svg files into a single PDF file: {output_pdf_filename}")


if __name__ == "__main__":
    svg_directory = r''  # You can change this to your specific SVG files directory
    output_pdf_filename = r''  # You can change this to your desired output PDF filename
    inkscape_path = r''  # Full path to the Inkscape executable

    check_inkscape_path(inkscape_path)

    main(svg_directory, output_pdf_filename, inkscape_path)
