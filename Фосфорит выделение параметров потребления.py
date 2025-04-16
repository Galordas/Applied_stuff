import pandas as pd
import os

def process_excel(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=None)

        # Print the first few rows to understand the structure
        print("Debugging: First few rows of the file:")
        print(df.head())

        # Temporary storage for averaging
        temp_data = {}

        # Iterate through each row in the DataFrame
        for index, row in df.iterrows():
            # Join the row into a string for easier parsing
            row_str = " ".join(str(cell) for cell in row if pd.notna(cell))

            # Debug: Print the row string
            print(f"Debugging: Row {index}: {row_str}")

            # Extract individual values for averaging
            if "Ia" in row_str or "Ib" in row_str or "Ic" in row_str:
                parts = row_str.split()
                if len(parts) >= 4:  # Ensure there are enough values
                    # Extract Ia, Ib, Ic values
                    ia = float(parts[parts.index("Ia") + 1].replace(",", "."))
                    ib = float(parts[parts.index("Ib") + 1].replace(",", "."))
                    ic = float(parts[parts.index("Ic") + 1].replace(",", "."))
                    iavg = (ia + ib + ic) / 3
                    print(f"Debugging: Calculated Iavg = {iavg}")

                    # Replace Ia, Ib, Ic with Iavg
                    df.loc[index, df.columns[parts.index("Ia") + 1]] = iavg
                    df.loc[index, df.columns[parts.index("Ib") + 1]] = iavg
                    df.loc[index, df.columns[parts.index("Ic") + 1]] = iavg

            # Extract individual values for Ubc, Uca, Uab averaging
            if "Ubc" in row_str or "Uca" in row_str or "Uab" in row_str:
                parts = row_str.split()
                if len(parts) >= 4:  # Ensure there are enough values
                    # Extract Ubc, Uca, Uab values
                    ubc = float(parts[parts.index("Ubc") + 1].replace(",", "."))
                    uca = float(parts[parts.index("Uca") + 1].replace(",", "."))
                    uab = float(parts[parts.index("Uab") + 1].replace(",", "."))
                    uavg = (ubc + uca + uab) / 3
                    print(f"Debugging: Calculated Uavg = {uavg}")

                    # Replace Ubc, Uca, Uab with Uavg
                    df.loc[index, df.columns[parts.index("Ubc") + 1]] = uavg
                    df.loc[index, df.columns[parts.index("Uca") + 1]] = uavg
                    df.loc[index, df.columns[parts.index("Uab") + 1]] = uavg

        # Save the processed data to a new Excel file
        output_file_path = os.path.join(os.path.dirname(file_path), "processed_" + os.path.basename(file_path))
        df.to_excel(output_file_path, index=False)
        print(f"Processed data saved to: {output_file_path}")

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

# Example usage
file_path = r'C:\Users\stazher3\Desktop\Выгрузка Фосфорит\Выгрузка Фосфорит (улучшенная)\РПТ 12 о.xlsx'  # Replace with the actual path
process_excel(file_path)