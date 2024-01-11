###  About
What this set of automation does:
- Extract emails from a certain label using Gmail API
- Download the password-encrypted zipped files from the extracted emails
- Reads data from Google Sheets using Sheets API
- Unzip file using password provided in Google Sheet
- Appends csv data in file onto Spreadsheet

### Setup

1. Clone this repository or download the zip file of the code and proceed to download the requirements file.

2. Set up your OAuth2.0 credentials at https://console.cloud.google.com and download the credentials into your computer locally.

3. Create a spreadsheet to input all the relevant details of the different emails and labels you want to process, and strictly follow the exact names of the column headers.

    ![Alt text](https://github.com/sparkfn/EKO-NETS/blob/main/image.png)

4. In the `sample.env` file, replace the variables with the your own details. For `SPREADSHEET_ID` and `PASSWORD_SHEET` in the file, these sheet details are of where your compilation of emails and passwords are. Do remember to include apostrophes `'` before and after all variables are filled. Here is the format on how you should replace local file paths:

    - For Windows: 'C:\\path\\to\\document.pdf' *(Double back slashes)*
    - For macOS: 'C:/path/to/document.pdf' *(Single slash)*

5. Afterwards, rename `sample.env` file to become `.env`.

6. To execute running of script, run `nets_main.py`.

7. For the first time running the script, a pop-up site or a link will show up for authentication. Accept all authentication to start running. Authentication will not be needed for future runs of the same gmail.
