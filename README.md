Macro Uploader Version 1.0.0

Disclaimer this is NOT AN OFFICAL PEXIP PRODUCT. This is a community-built tool.

This tool is designed to upload Cisco macros in bulk. Created using Python, tkinter, and customtkinter libraries.

The Macro Uploader tool has two modes of operation. Single Mode and Pexip OTJ Mode. 

The Single Mode will upload a single macro to multiple endpoints via a .csv file.

The Pexip OTJ Mode will upload both the required otj-macro.js and otj-macro-settings.js files to their corresponding endpoint. This is done by first browsing to a directory that contains the Pexip OTJ macros. When an OTJ macro package is downloaded from the Pexip Control Center Portal it will download with the naming convention "endpoint_name-otj-macro-latest.zip". This zip file contains two files, the otj-macro.js and the otj-macro-settings.js. Macro Uploader will scan all the zip files within a directory and first ensure they contain both of these files. If they do, it will take the "endpoint_name" portion of the zip file and compare that against the name in the csv file. If they match it will proceed to unzip and add the file paths to the csv. The user can then press the upload macros button which will loop through the csv.

Note the CSV must be formatted with the columns "name", "ip", "username", "password", "macro settings path", and "macro path".

The source code has been compiled into a zip file containing an exe. Included in the Zip file is also a blank template for the csv file.Depending on your antivirus software the exe could be flagged. Microsoft has already cleared this from Windows Defender. 


Instructions for use:

Single Mode
1. Fill out the csv that is provided in the zip file
2. Do not modify the column headers as the tool will look to ensure the values match
3. Ensure to populate the name of the unit, ip, username, and password
4. Macro settings path and macro path column should be left blank for single mode as they are not used
5. Browse out to the Cisco macro-JavaScript file
6. Browse out to the CSV file. Ensure that the file is not open in another application such as excel.
7. Click the Upload Macros Button. The tool will then loop through all rows of the CSV file and attempt to upload. 
8. A .log file will be created in the directory for review

Pexip OTJ Mode
1. You must have a directory containing the macro zip files from Pexip Control Center. Do not unzip them as the tool will unzip them for you.
2. Add all macro zip files to the same directory
3. Do not rename the zip files as the tool will attempt to match the name prefix to the name column in the csv
4. Fill out the csv that is provided in the zip file. Ensure the name column matches the name prefix of the zip files. 
5. Do not modify the column headers as the tool will look to ensure the values match
6. Ensure to populate the name of the unit, ip, username, and password
7. Leave the macro settings path and macro path blank as the tool will write these for us
8. In the tool browse out to the macro directory
9. Browse out to the CSV. Ensure that the file is not open in another application such as excel.
10. Click the Scan Directory button. The tool will scan the directory and let you know how many files were found.
11. Click yes in the pop-up to extract and add the paths to the csv
12. Review the CSV and ensure it populates the paths correctly
13. Click the Upload Macros Button. The tool will then loop through all rows of the CSV file and attempt to upload. 
14. A .log file will be created in the directory for review




