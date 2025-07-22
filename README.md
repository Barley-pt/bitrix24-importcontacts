# Bitrix24 Contact Importer

This Python script provides a user-friendly graphical interface (GUI) to import contacts from an Excel file (.xlsx or .xls) into your Bitrix24 CRM. It allows for mapping Excel columns to Bitrix24 contact fields and includes an optional duplicate check to avoid creating redundant entries.

## Features

  * **GUI-based:** Easy-to-use interface built with Tkinter.
  * **Excel Support:** Reads contact data from `.xlsx` and `.xls` files.
  * **Dynamic Field Mapping:** Automatically fetches available Bitrix24 contact fields and allows you to map your Excel columns to them.
  * **Optional Duplicate Check:** Prevents duplicate contact creation by checking for existing contacts based on email or phone number before importing.
  * **Import Status Logging:** Provides a summary of successful and failed imports.
  * **Bitrix24 ID Tracking:** Updates the original Excel file with the newly created or found Bitrix24 Contact IDs.

## Requirements

  * Python 3.x
  * `openpyxl` library
  * `requests` library

## Installation

1.  **Install the required Python libraries:**

    ```bash
    pip install openpyxl requests
    ```

## Usage

1.  **Prepare your Excel file:**

      * Ensure your Excel file has a header row in the first row.
      * Make sure the columns containing email and phone numbers are clearly identifiable if you plan to use the duplicate check.

2.  **Run the script:**

    ```bash
    python import_contacts.py
    ```

3.  **Follow the GUI prompts:**

      * **Duplicate Check:** A dialog box will appear asking if you want to check for duplicates. It's highly recommended to enable this.
      * **Select Excel File:** A file dialog will open. Navigate to and select your Excel file (`.xlsx` or `.xls`).
      * **Enter Bitrix24 Webhook URL:** A prompt will appear asking for your Bitrix24 Incoming Webhook URL. This is essential for the script to communicate with your Bitrix24 instance.
          * **How to get your Webhook URL:**
            1.  In your Bitrix24, go to `Applications` (or `Market`) \> `Webhooks`.
            2.  Click `Add webhook` \> `Incoming webhook`.
            3.  Give it a name (e.g., "Contact Importer Webhook").
            4.  Select `CRM` \> `Contact` \> `Add` and `CRM` \> `Contact` \> `Read` permissions.
            5.  Save and copy the `Webhook URL`. It will look something like `https://yourdomain.bitrix24.com/rest/1/xxxxxxxxxxxxxxxx/`.
      * **Field Mapping:** A new window will display your Excel column headers on the left and dropdown menus on the right. Map each Excel header to the corresponding Bitrix24 contact field. Bitrix24 fields are shown with their internal key and a user-friendly label (e.g., `NAME - First Name`, `EMAIL - Email`).
      * **Start Import:** Click the "Start Import" button to begin the process.

4.  **Review Results:** Once the import is complete, a message box will display the number of successful and failed imports. A new Excel file will be saved in the same directory as your original file, with `_bitrix_imported` appended to its name (e.g., `your_file_bitrix_imported.xlsx`). This new file will contain an additional column named `BITRIX_ID` with the respective Bitrix24 contact IDs.

## How it Works

The script performs the following steps:

1.  **Initial Setup:** Asks the user if a duplicate check should be performed.
2.  **File Selection:** Prompts the user to select an Excel file and reads its headers.
3.  **Webhook Input & Field Fetching:** Requests the Bitrix24 webhook URL and then uses it to fetch a list of available contact fields from your Bitrix24 instance. This ensures dynamic and accurate mapping.
4.  **Field Mapping GUI:** Displays a GUI where the user visually maps Excel columns to the fetched Bitrix24 fields.
5.  **Import Logic:**
      * Iterates through each row of the Excel file (starting from the second row to skip headers).
      * Constructs a payload for Bitrix24's `crm.contact.add` method.
      * If duplicate checking is enabled, it uses `crm.contact.list` to search for existing contacts based on email or phone before attempting to add a new one.
      * Sends a POST request to the Bitrix24 API to add the contact.
      * Records the success or failure of each import.
      * Updates the `BITRIX_ID` column in the Excel sheet with the ID of the newly created or found Bitrix24 contact.
6.  **Save Results:** Saves the updated Excel file with the Bitrix24 IDs.
7.  **Summary:** Displays a final summary of the import process.

## Error Handling

  * The script includes basic error handling for fetching Bitrix24 fields and for issues during the import process.
  * Errors during individual contact imports are printed to the console.

## Contributing

Feel free to fork this repository, make improvements, and submit pull requests.

## License

This project is open-source and available under the [MIT License](https://www.google.com/search?q=LICENSE).
