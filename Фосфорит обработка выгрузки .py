import pandas as pd
import re
import os


def parse_excel(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=None)

        # Initialize patterns for object names, time of measurement, and parameters
        time_pattern = re.compile(r'\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}\.\d{3}')  # Timestamps
        parameter_pattern = re.compile(
            r'^[A-Za-zА-Яа-я]+\s+[A-Za-zА-Яа-я]+\s+.*\s+\d+\.?\d*\s*[A-Za-zА-Яа-я]*')  # Parameter rows
        cosf_pattern = re.compile(r'^cosf\s+.*')  # Special pattern for cosf parameters

        # Initialize variables to store the parsed data
        result = []
        current_object = None
        current_time = None

        # Detect object names dynamically from the data
        object_name_pattern = re.compile(r'^([A-Za-zА-Яа-я]+\s*-?\d+).*')  # Object names (e.g., "Ввод-1")

        # Iterate through each row in the DataFrame
        for _, row in df.iterrows():
            # Convert the row to a string for easier pattern matching
            row_str = " ".join(str(cell) for cell in row if pd.notna(cell))

            # Check if the row contains the object name
            if object_name_pattern.match(row_str):
                match = object_name_pattern.match(row_str)
                current_object = match.group(1).strip()
                continue

            # Check if the row contains the time of measurement
            if time_pattern.search(row_str):
                match = time_pattern.search(row_str)
                current_time = match.group(0).strip()
                continue

            # Check if the row contains a cosf parameter
            if cosf_pattern.match(row_str):
                parts = row_str.split()
                parameter_type = "cos"
                parameter_name = ' '.join(parts[1:-1])  # Handle cosf-specific structure
                value = parts[-1].replace(".", ",")  # Replace '.' with ',' for decimal separator
                unit = "-"  # cosf parameters don't have a unit

                # Append the parsed data to the result list
                result.append({
                    'object_name': current_object,
                    'time_of_measurement': current_time,
                    'parameter_type': parameter_type,
                    'parameter_name': parameter_name,
                    'value': value,
                    'unit': unit
                })
                continue

            # Check if the row contains a regular parameter
            if parameter_pattern.match(row_str):
                parts = row_str.split()
                parameter_type = parts[0]
                parameter_name = ' '.join(parts[1:-2])
                value = parts[-2].replace(".", ",")  # Replace '.' with ',' for decimal separator
                unit = parts[-1]

                # Append the parsed data to the result list
                result.append({
                    'object_name': current_object,
                    'time_of_measurement': current_time,
                    'parameter_type': parameter_type,
                    'parameter_name': parameter_name,
                    'value': value,
                    'unit': unit
                })

        # Convert the result list to a DataFrame
        result_df = pd.DataFrame(result)
        return result_df

    except Exception as e:
        print(f"Error parsing file {file_path}: {e}")
        return pd.DataFrame()


def process_directory(input_directory, output_directory):
    # Ensure the output directory exists
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Process all Excel files in the input directory
    for file_name in os.listdir(input_directory):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_directory, file_name)
            print(f"Processing file: {file_path}")

            # Parse the Excel file
            parsed_data = parse_excel(file_path)

            # Save the parsed data to a new Excel file in the output directory
            if not parsed_data.empty:
                output_file_path = os.path.join(output_directory, f"parsed_{file_name}")
                parsed_data.to_excel(output_file_path, index=False)
                print(f"Parsed data saved to: {output_file_path}")
            else:
                print(f"No data parsed from file: {file_path}")


# Example usage
input_directory = r''
output_directory = r''

# Process all Excel files in the directory
process_directory(input_directory, output_directory)
