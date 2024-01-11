import helper
import globals
import os

import helper
import globals

def main():

    helper.createFolderIfDoesNotExist(globals.DOWNLOADS_FOLDER)
    # Get 2d array of all data including header row in sheet
    data = helper.read_data_from_spreadsheet(globals.SPREADSHEET_ID, globals.PASSWORD_SHEET)

    allEmails = []
    start = 0
    # Through looping the header row, get the indexes of each variable
    for row in data:
        if start == 0:
            i = 0
            for element in row:
                if element == 'Email':
                    emailIndex = i
                if element == 'Encryption Password':
                    passwordIndex = i
                if element == 'Label To Process':
                    labelToProcessIndex = i
                if element == 'Processed Label':
                    processedLabelIndex = i
                if element == 'Spreadsheet ID':
                    spreadsheetIdIndex = i
                if element == 'Sheet Name':
                    sheetNameIndex = i  
                i += 1
            start = 1
        else:
            # Get 2d array of only actual email accounts
            allEmails.append(row)

    for row in allEmails:
        try:
            # Get indexes of all variables shifting of columns wont be affected
            email = row[emailIndex]
            password = row[passwordIndex]
            labelToProcess = row[labelToProcessIndex]
            processedLabel = row[processedLabelIndex]
            spreadsheetId = row[spreadsheetIdIndex]
            sheetName = row[sheetNameIndex]
            
            messages = helper.get_messages_with_label(labelToProcess)
            # Loop through all messages in label and only process email if the recipient matches
            for message in messages:
                try:
                    # Match email to password and spreadsheet details
                    messageId = message['id']
                    recipient = helper.get_recipient_of_email(messageId)

                    if recipient == email:
                        print(f"Message ID: {messageId}")
                        # Download attachments
                        downloaded_file_paths = helper.download_attachments(messageId, globals.DOWNLOADS_FOLDER)

                        for zip_file_path in downloaded_file_paths:
                            # Unzip attachment and delete zipped file
                            unzippedFilePath = helper.unzip_password_protected_zip(zip_file_path, password, globals.DOWNLOADS_FOLDER)
                            os.remove(zip_file_path)
                            print('Removed zipped file.')

                            # Read and append data into spreadsheet
                            allData2dArray = helper.read_csv_file(unzippedFilePath)            
                            dataToAppend = helper.extractDataFrom2dArray(allData2dArray)
                            helper.appendToSpreadsheet(dataToAppend, spreadsheetId, sheetName)
                            os.remove(unzippedFilePath)
                            print('Removed unzipped file.')

                        # Move messages into different labels
                        helper.add_label_to_message(messageId, processedLabel)
                        print('Added processed label.')
                        helper.remove_label_from_message(messageId, labelToProcess)
                        print('Removed to process label.')

                except Exception as e:
                    print(e)
                    
        except Exception as e:
            helper.sendErrorEmail(globals.EXCEPTION_EMAIL_ADDRESS, '')
            print(e)

    print("The script has finished running.")
