# Macro Uploader
Tool for uploading Cisco Macros to Cisco endpoints

The Macro Uploader tool has two modes of opertation. Single Mode and Pexip OTJ Mode. 

The Single Mode will upload a single macro to multiple endpoints via a .csv file.

The Pexip OTJ Mode will upload both the required otj-macro.js and otj-macro-settings.js files to their corresponding endpoint. This is done by first browsing to a directory that contains the Pexip OTJ macros. When an OTJ macro package is downloaded from the Pexip Control Center Portal it will download with the naming convention "endpoint_name-otj-macro-latest.zip". This zip file contains two files, the otj-macro.js and the otj-macro-settings.js. Macro Uploader will scan all the zip files within a directory and first ensure they contain both of these files. If they do, it will take the "endpoint_name" portion of the zip file and compare that against the name in the csv file. If they match it will proceed to unzip and add the file paths to the csv. The user can then press the upload macros button which will loop through the csv.

The source code has been compliled into on.py file and included is the csv with the correct format. Additonally, an exe directory package is included so that users don't require python or the libraries. They can simply download the app package and run it using the "macro_uploader.exe" file.


