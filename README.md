# Macro Uploader
Tool for uploading Cisco Macros to Cisco endpoints. Created using Python, tkinter, and customtkinter libraries.

The Macro Uploader tool has two modes of opertation. Single Mode and Pexip OTJ Mode. 

The Single Mode will upload a single macro to multiple endpoints via a .csv file.

The Pexip OTJ Mode will upload both the required otj-macro.js and otj-macro-settings.js files to their corresponding endpoint. This is done by first browsing to a directory that contains the Pexip OTJ macros. When an OTJ macro package is downloaded from the Pexip Control Center Portal it will download with the naming convention "endpoint_name-otj-macro-latest.zip". This zip file contains two files, the otj-macro.js and the otj-macro-settings.js. Macro Uploader will scan all the zip files within a directory and first ensure they contain both of these files. If they do, it will take the "endpoint_name" portion of the zip file and compare that against the name in the csv file. If they match it will proceed to unzip and add the file paths to the csv. The user can then press the upload macros button which will loop through the csv.

Note the CSV must be formatted with the columns "name", "ip", "username", "password", "macro settings path", and "macro path".

The source code has been compliled into a zipfile containing an exe. Included in the Zip file is also a blank template for the csv file.


