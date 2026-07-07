import argparse
import sys
import os
import config as cfg
import process as prc
import analyze as anl


def main():
    """
    Data extraction utilities.
    Handles the extraction of data from a directory of phishing campaigns.
    """
    parser = argparse.ArgumentParser(
        description="Phishing Campaign Batch Data Extraction Utility",
        add_help=False,
    )

    # Flags adapted for directories
    parser.add_argument(
        "-i",
        "--input",
        metavar="<INPUT_DIR>",
        help="Path to the input directory containing the files to be analyzed (Required).",
    )
    parser.add_argument(
        "--output",
        metavar="<OUTPUT_DIR>",
        help="Path to the output directory to save the extracted data (Optional).",
    )
    parser.add_argument(
        "-h", "--help", action="help", help="Show this help message and exit."
    )

    args = parser.parse_args()

    # Validate required input and check if it's a valid directory
    if not args.input:
        print("=" * 54)
        parser.print_help()
        sys.exit(1)

    # Remove trailing slash if the user adds one (e.g., "folder/" becomes "folder")
    args.input = os.path.abspath(args.input)

    if not os.path.isdir(args.input):
        print(f"\n[ERROR] The input path '{args.input}' is not a valid directory!")
        sys.exit(1)

    # Output directory default logic
    if not args.output:
        args.output = f"{args.input}_results"

    # Create the output directory if it does not exist
    os.makedirs(args.output, exist_ok=True)

    print("=" * 54)
    print(f" -> Input Dir:  {args.input}")
    print(f" -> Output Dir: {args.output}")
    print("=" * 54)
    print("Starting batch data extraction...\n")

    # Master output file
    output_file_path = os.path.join(args.output, "phishing_results.xlsx")

    # Iterate over all files in the directory
    for filename in os.listdir(args.input):
        input_file_path = os.path.join(args.input, filename)

        # Ensure we only process files AND that they are CSV files
        if os.path.isfile(input_file_path) and filename.lower().endswith(".csv"):
            print(f"[*] Processing file: {filename}")

            # Chama a função sempre com o mesmo output_file_path
            prc.process_file(input_file_path, output_file_path)

        elif os.path.isfile(input_file_path):
            print(f"[-] Skipping non-CSV file: {filename}")

    print("\n[+] All files processed successfully!")

    print("\n[+] Processing analytics report...")
    anl.generate_report(output_file_path)
    print("\n[+] Analytics report generated successfully!")
    print(f"\n[+] Process completed. Results saved to {output_file_path}.")


if __name__ == "__main__":
    cfg.banner()
    main()
