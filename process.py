import os
import pandas as pd
import config as cfg


def process_file(input_file_path, output_file_path):
    """
    Extracts data from a CSV and appends it to a master Excel file.
    Creates the Excel file if it doesn't exist.
    """
    try:
        # Load the CSV file
        df = pd.read_csv(input_file_path, dtype=str, low_memory=False)

        # Check for target headers
        existing_columns = [
            header for header in cfg.PHISHING_HEADERS if header in df.columns
        ]

        if not existing_columns:
            print(
                f"   [WARNING] No target headers found in {os.path.basename(input_file_path)}. Skipped."
            )
            return

        # Filter the data & Create column for the origin file
        extracted_data = df[existing_columns].copy()
        file_name = os.path.basename(input_file_path)
        extracted_data.insert(0, "Source File", file_name)

        # Append to master Excel file
        if not os.path.exists(output_file_path):
            # If file doestn't exist, create it with the extracted data and selected headers
            extracted_data.to_excel(
                output_file_path, index=False, sheet_name="Raw Data"
            )
            print("   [+] Created master file and saved data.")
        else:
            # If it exists, append the new data without headers
            with pd.ExcelWriter(
                output_file_path, mode="a", engine="openpyxl", if_sheet_exists="overlay"
            ) as writer:
                # Find the last written row to append data without overwriting
                startrow = writer.sheets["Raw Data"].max_row

                # Write the new data starting from the next row
                extracted_data.to_excel(
                    writer,
                    sheet_name="Raw Data",
                    index=False,
                    header=False,
                    startrow=startrow,
                )
            print("   [+] Appended data to master file.")
    except pd.errors.EmptyDataError:
        print(
            f"   [WARNING] The file {os.path.basename(input_file_path)} is empty. Skipped."
        )
    except Exception as e:
        print(f"   [ERROR] Failed to process {os.path.basename(input_file_path)}: {e}")
