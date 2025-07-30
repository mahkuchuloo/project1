# RFM Analyzer

## Overview

The RFM Analyzer is a Python-based tool with a graphical user interface (GUI) designed to process transaction data and generate valuable insights for donor analysis. It uses the RFM (Recency, Frequency, Monetary) model to segment donors and provide detailed metrics for each donor.

## Features

- User-friendly GUI for easy interaction
- Imports CSV and Excel files containing transaction data
- Exports results to CSV or Excel files
- Calculates key metrics such as:
  - Total number of gifts
  - Lifetime giving
  - Last gift details
- Implements RFM (Recency, Frequency, Monetary) scoring
- Provides gift amount range categorization
- Identifies digital monthly status for the last gift

## Requirements

- Python 3.6+
- pandas
- numpy
- openpyxl (for Excel file support)
- tkinter (usually comes pre-installed with Python)

## Installation

1. Ensure you have Python 3.6 or later installed on your system.
2. Install the required packages:

   ```
   pip install pandas numpy openpyxl
   ```

## Usage

1. Run the script:

   ```
   python rfm_analyzer.py
   ```

2. The RFM Analyzer GUI will open.

3. Click on "Select Input File" to choose your input file (CSV or Excel) containing transaction data.

4. Once a file is selected, the "Process Data" button will become active.

5. Click "Process Data" to start the RFM analysis.

6. Choose a location and name for the output file (CSV or Excel) when prompted.

7. The tool will process the data and save the results to the specified output file.

8. Once processing is complete, you can click on "Open Output Directory" to view the results.

## Input File Format

The input file (CSV or Excel) should contain the following columns:

- Relationship ID
- Transaction ID
- Giving Platform
- Date Clean
- Amount
- Donor Email
- Donor Phone
- Donor Address Line 1
- Donor City
- Donor State
- Donor ZIP
- MSA
- Recipient
- Recurring ID

## Output

The script generates a file (CSV or Excel) containing detailed donor analysis, including:

- Donor information (ID, contact details)
- RFM Score
- Total Number of Gifts
- Lifetime Giving
- Last Gift Details
- Gift Amount Range
- Digital Monthly Status

## Troubleshooting

If you encounter any issues:

1. Ensure your input file matches the expected format.
2. Check that you have the required Python packages installed.
3. Verify that the file paths for input and output are correct.

For any persistent issues, please check the error messages in the log window of the GUI.

## Contributing

Contributions to improve the RFM Analyzer are welcome. Please feel free to submit pull requests or open issues to discuss potential enhancements.

## License

This project is licensed under the MIT License.
