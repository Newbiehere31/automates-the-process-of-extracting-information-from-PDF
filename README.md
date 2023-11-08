This code is a Python script that performs the following tasks:
1. It reads data from an Excel file containing PDF links and report numbers.
2. It downloads PDF files from Google Drive using the provided links, extracts text from these PDFs, and attempts to match the report numbers in the extracted text with the report numbers in the Excel file.
3. It also extracts dates from the PDF text and checks if the extracted dates match a predefined format.
4. Based on the matching status of report numbers and dates, it updates a "Status" column and a "Date Status" column in the Excel file, marking rows with a yellow background if there is no match.
Overall, the code automates the process of extracting information from PDFs, comparing it to data in an Excel file, and updating the file with the matching status.

Note: Currently not working on this project. 
