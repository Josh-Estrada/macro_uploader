import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
import customtkinter
import os
import re
import zipfile
import threading
import pandas as pd
import requests
from lxml import etree
import logging
import time

customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class PexipMacroApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.create_widgets()
        self.single_mode()
        self.create_console()
        self.create_logger()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_logger(self):
        # Create a new logger for this function
        self.logger = logging.getLogger('macro_uploader')
        self.logger.setLevel(logging.DEBUG)
        self.file_handler = logging.FileHandler('macro_uploader.log')
        self.formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.file_handler.setFormatter(self.formatter)
        self.logger.addHandler(self.file_handler)

    def create_widgets(self):
        # configure window
        self.headers = {'Content-Type': 'text/xml'}
        self.title("Macro Uploader")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0), weight=0)
        self.grid_rowconfigure((1, 2, 3, 4), weight=0)
        self.grid_rowconfigure((5), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="MacroUploader", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.macro_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Macro Mode:", anchor="w")
        self.macro_mode_label.grid(row=1, column=0, padx=20, pady=(10, 0))
        self.macro_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Single Macro", "Pexip OTJ"],
                                                                       command=self.handle_mode_selection)
        self.macro_mode_optionemenu.grid(row=2, column=0, padx=20, pady=(10, 10))

        # Add macro label and entry field
        self.macro_entry = customtkinter.CTkEntry(self, placeholder_text="Browse to .js file",state="readonly")
        self.macro_browse_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Browse .js', command=self.browse_js_file)
        # Add CSV label and entry field
        self.csv_entry = customtkinter.CTkEntry(self, placeholder_text="Browse to .csv file", state="readonly")
        self.csv_browse_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Browse .csv', command=self.browse_csv_file)
        # Upload Macros Button in Single Macro Mode
        self.upload_macros_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Upload Macros', command=self.upload_macros)
        # Add OTJ Macro Directory label, buttons, and entry field
        self.otj_check_macros_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Scan Directory', command=self.count_dirs_with_files)
        self.otj_macro_entry = customtkinter.CTkEntry(self, placeholder_text="Browse to otj macro directory",state="readonly")
        self.otj_macro_browse_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Browse Dir', command=self.browse_OTJ_dir_file)
        # Add CSV label and entry field
        self.csv2_entry = customtkinter.CTkEntry(self, placeholder_text="Browse to .csv file", state="readonly")
        self.csv2_browse_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Browse .csv', command=self.browse_csv2_file)
        # Upload Macros Button in OTJ Macro Mode
        self.otj_upload_macros_button = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text='Upload Macros', command=self.upload_macros_otj)

        # set default values
        self.appearance_mode_optionemenu.set("Light")

    def handle_mode_selection(self, selected_mode):
        if selected_mode == "Single Macro":
            self.single_mode()
        elif selected_mode == "Pexip OTJ":
            self.otj_mode()


    def single_mode(self):
        #Single Macro Mode grid widgets
        self.macro_entry.grid(row=1, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
        self.macro_browse_button.grid(row=1, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.csv_entry.grid(row=2, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
        self.csv_browse_button.grid(row=2, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.upload_macros_button.grid(row=3, column=1, padx=(20, 0), pady=(20, 20), sticky="nsew")
        #Forget OTJ macro mode
        self.otj_check_macros_button.grid_forget()
        self.otj_macro_entry.grid_forget()
        self.otj_macro_browse_button.grid_forget()
        self.csv2_entry.grid_forget()
        self.csv2_browse_button.grid_forget()
        self.otj_upload_macros_button.grid_forget()


    def otj_mode(self):
        #OTJ Macro Mode grid widgets
        self.otj_check_macros_button.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.otj_macro_entry.grid(row=1, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
        self.otj_macro_browse_button.grid(row=1, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.csv2_entry.grid(row=2, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
        self.csv2_browse_button.grid(row=2, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.otj_upload_macros_button.grid(row=3, column=1, padx=(20, 0), pady=(20, 20), sticky="nsew")
        self.otj_upload_macros_button.configure(state="disabled")
        #Forget Single macro mode except CSV
        self.macro_entry.grid_forget()
        self.macro_browse_button.grid_forget()
        self.csv_entry.grid_forget()
        self.csv_browse_button.grid_forget()
        self.upload_macros_button.grid_forget()


    def create_console(self):
        # Add console output text widget with scroll bar
        # create textbox
        self.console_text = customtkinter.CTkTextbox(self, width=250, state="normal")
        self.console_text.grid(row=4, column=1, rowspan=5, padx=(20, 0), pady=(20, 20), sticky="nsew")

        # Bind events to prevent editing of text
        def return_break(event):
            return "break"

        self.console_text.bind("<Key>", return_break)
        self.console_text.bind("<FocusIn>", return_break)


    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)
        

    def browse_js_file(self):
        try:
            self.js_file_path = filedialog.askopenfilename(filetypes=[("JavaScript files", "*.js")])
            if self.js_file_path:
                self.macro_entry.configure(state="normal")
                self.macro_entry.delete(0, tk.END)
                self.macro_entry.insert(0, self.js_file_path)
                self.macro_entry.configure(state="readonly")
                self.console_text.insert(tk.END, f"\nSelected JS file: {self.js_file_path}\n")
                self.logger.info(f"\nSelected JS file: {self.js_file_path}\n")
                self.console_text.see(tk.END)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error selecting JS file: {e}")
            self.console_text.insert(tk.END, f"\nError selecting JS file: {e}\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nError selecting JS file: {e}\n")


    def browse_csv_file(self):
        try:
            self.file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
            if self.file_path:
                self.csv_entry.configure(state="normal")
                self.csv_entry.delete(0, tk.END)
                self.csv_entry.insert(0, self.file_path)
                self.csv_entry.configure(state="readonly")
                self.console_text.insert(tk.END, f"\nSelected CSV file: {self.file_path}\n")
                self.console_text.see(tk.END)
                self.logger.info(f"\nSelected CSV file: {self.file_path}\n")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error selecting CSV file: {e}")
            self.console_text.insert(tk.END, f"\nError selecting CSV file: {e}\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nError selecting CSV file: {e}\n")

    def browse_csv2_file(self):
        try:
            self.file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
            if self.file_path:
                self.csv2_entry.configure(state="normal")
                self.csv2_entry.delete(0, tk.END)
                self.csv2_entry.insert(0, self.file_path)
                self.csv2_entry.configure(state="readonly")
                self.console_text.insert(tk.END, f"\nSelected CSV file: {self.file_path}\n")
                self.console_text.see(tk.END)
                self.logger.info(f"\nSelected CSV file: {self.file_path}\n")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error selecting CSV file: {e}")
            self.console_text.insert(tk.END, f"\nError selecting CSV file: {e}\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nError selecting CSV file: {e}\n")


    def browse_OTJ_dir_file(self):
        try:
            self.otj_dir_file_path = filedialog.askdirectory()
            if self.otj_dir_file_path:
                self.otj_macro_entry.configure(state="normal")
                self.otj_macro_entry.delete(0, tk.END)
                self.otj_macro_entry.insert(0, self.otj_dir_file_path)
                self.otj_macro_entry.configure(state="readonly")
                self.console_text.insert(tk.END, f"\nSelected directory: {self.otj_dir_file_path}\n")
                self.console_text.see(tk.END)
                self.logger.info(f"\nSelected directory: {self.otj_dir_file_path}\n")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error selecting directory: {e}")
            self.console_text.insert(tk.END, f"\nError selecting directory: {e}\n")


    def upload_macros(self):
        try:
            filename = self.csv_entry.get()
            js_file_path = self.macro_entry.get()
            if not filename or not js_file_path:
                raise ValueError("Both CSV data file and JS macro file are required")
            macro_name = os.path.splitext(os.path.basename(js_file_path))[0]
            #Disable all buttons while the function runs
            self.upload_macros_button.configure(state="disabled")
            self.csv_browse_button.configure(state="disabled")
            self.macro_browse_button.configure(state="disabled")
            threading.Thread(target=self.read_csv_data, args=(filename, js_file_path, macro_name)).start()
        except ValueError as e:
            tk.messagebox.showerror("Error", str(e))
            self.console_text.insert(tk.END, f"\nError: {e}\n")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error uploading macros: {e}")
            self.console_text.insert(tk.END, f"\nError uploading macros: {e}\n")


    def save_macro(self, endpoint_ip, username, password, macro_name, js_file_path, index):
        try:
            with open(js_file_path, 'r') as f:
                js_code = f.read()
        except FileNotFoundError as e:
            error_msg = f'Error: File {js_file_path} not found'
            self.console_text.insert(tk.END, f"Error file {js_file_path} not found\n")
            self.console_text.see(tk.END)
            self.logger.error(error_msg)
            self.logger.exception(e)
            return

        tags = ['Macros', 'Macro', 'Save']
        macro_params = {'name': macro_name,
                        'body': js_code,
                        'overWrite': "True",
                        'Transpile': "False",
                        }

        xml = root = etree.Element("Command")
        for tag in tags:
            xml = etree.SubElement(xml, tag)
        xml.attrib["command"] = "True"
        for (arg, value) in macro_params.items():
            arg_xml = etree.SubElement(xml, arg)
            arg_xml.text = str(value)

        data = etree.tostring(root)

        # Send the xAPI command to the endpoint
        try:
            with requests.post(f'https://{endpoint_ip}/putxml', auth=(username, password), data=data, headers=self.headers, verify=False, timeout=10) as response:
                if response.status_code == 200:
                    success_msg = f'Row {index}. {macro_name} saved successfully for: {endpoint_ip}'
                    self.console_text.insert(tk.END, f"\n{success_msg}\n")
                    self.console_text.see(tk.END)
                    self.logger.info(success_msg)
                    self.enable_otj_macro(endpoint_ip, username, password, macro_name, index)
                elif response.status_code == 401:
                    error_msg = f'\nRow {index}. 401 error. Connection for: {endpoint_ip} unauthorized. Please check your credentials and try again.\n'
                    self.logger.error(f'{error_msg}')
                    self.console_text.insert(tk.END, f"{error_msg}\n")
                    self.console_text.see(tk.END)
                else:
                    # Log the error message from the HTTP response
                    error_msg = f'Row {index}. Error saving {macro_name} to: {endpoint_ip } {response.text}'
                    self.logger.error(error_msg)
                    self.console_text.insert(tk.END, f"\n{error_msg}\n")
                    self.console_text.see(tk.END)
        except requests.exceptions.Timeout:
            # Log and print the timeout error message
            error_msg = f'Row {index}. Timed out waiting for a response from {endpoint_ip}.'
            self.logger.warning(error_msg)
            self.console_text.insert(tk.END, f"\n{error_msg}\n")
            self.console_text.see(tk.END)
            # Retry the request
        except requests.exceptions.RequestException as e:
            error_msg = f'Row {index}. Error sending request: {e}'
            self.logger.error(error_msg)
            self.console_text.insert(tk.END, f"\n{error_msg}\n")
            self.console_text.see(tk.END)



    #enable Macro and restart the macro runtime
    def enable_otj_macro(self, endpoint_ip, username, password, macro_name, index):
        # Set up the endpoint URL for the macro
        url = f"https://{endpoint_ip}/putxml"
        headers = {'Content-Type': 'application/xml'}

        # Set up the authentication credentials
        auth = (username, password)

        # Set the payload for enabling the macro
        root = etree.Element("Command")
        macros = etree.SubElement(root, "Macros")
        macro = etree.SubElement(macros, "Macro")
        activate = etree.SubElement(macro, "Activate")
        name = etree.SubElement(activate, "Name")
        name.text = macro_name
        payload = etree.tostring(root, encoding='unicode')

        # Send the POST request to the endpoint to enable the macro.
        response = requests.post(url, auth=auth, headers=headers, data=payload, verify=False)
        # Check the status code of the response
        if response.status_code == 200:
            self.console_text.insert(tk.END, f"\nRow {index}. {macro_name} enabled successfully for {endpoint_ip}\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nRow {index}. {macro_name} enabled successfully for {endpoint_ip}\n")
        else:
            self.console_text.insert(tk.END, f"\nRow {index}. Error enabling {macro_name} for {endpoint_ip}. Status code: {response.status_code}")
            self.console_text.see(tk.END)
            self.logger.error(f"\nRow {index}. Error enabling {macro_name} for {endpoint_ip}. Status code: {response.status_code}")

        # Set the payload for restarting the Macro runtime
        tags2 = ['Macros', 'Runtime', 'Restart']
        macro_params1 = {"command": "True"}

        root = etree.Element("Command")
        macros = etree.SubElement(root, tags2[0])
        runtime = etree.SubElement(macros, tags2[1])
        etree.SubElement(runtime, tags2[2], **macro_params1)
        payload2 = etree.tostring(root)

        # Send the POST request to restart the Macro runtime.
        response = requests.post(url, data=payload2, headers=headers, auth=auth, verify=False)

        # Check the status code of the response
        if response.status_code == 200:
            self.console_text.insert(tk.END, f"\nRow {index}. Macro runtime restarted successfully for {endpoint_ip}\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nRow {index}. Macro runtime restarted successfully for {endpoint_ip}\n")
        else:
            self.console_text.insert(tk.END, f"Row {index}. Error restarting Macro runtime for {endpoint_ip}. Status code: {response.status_code}")
            self.console_text.see(tk.END)
            self.logger.error(f"Row {index}. Error restarting Macro runtime for {endpoint_ip}. Status code: {response.status_code}")


    def read_csv_data(self, filename, js_file_path, macro_name):
        try:
            # Replace 'filename.csv' with the actual name of your CSV file
            data = pd.read_csv(filename, delimiter=',', header=0)
            self.console_text.insert(tk.END, f"\nMacro uploads starting... Do not close the application until complete.\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nMacro uploads starting... Do not close the application until complete.\n")
            time.sleep(1)
            self.console_text.insert(tk.END, f"\nParsing {filename} for all endpoints\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nParsing {filename} for all endpoints\n")
            time.sleep(1)

            # Check if the column names match (ignoring case)
            expected_columns = {'name', 'ip', 'username', 'password', 'macro settings path', 'macro path'}
            actual_columns = set(map(str.lower, data.columns))
            self.logger.info(f"Columns in CSV file: {data.columns}")
            if not actual_columns.issuperset(expected_columns):
                missing_columns = expected_columns - actual_columns
                extra_columns = actual_columns - expected_columns
                error_msg = ""
                if missing_columns:
                    error_msg += f"The following columns are missing from the CSV file: {', '.join(missing_columns)}. "
                if extra_columns:
                    error_msg += f"The following extra columns are present in the CSV file: {', '.join(extra_columns)}. "
                error_msg += "Please ensure that the columns are name, ip, username, password"
                self.logger.error(error_msg)
                self.console_text.insert(tk.END, f"\n{error_msg}\n")
                self.console_text.see(tk.END)
                self.enable_buttons()
                # raise an error and stop execution
                raise ValueError(error_msg)
            elif actual_columns != expected_columns:
                # if the actual columns set is not equal to the expected columns set, 
                # but contains all the required columns, then log a warning and continue execution
                extra_columns = actual_columns - expected_columns
                self.logger.warning(f"The following extra columns are present in the CSV file: {', '.join(extra_columns)}.\nPlease check the CSV file and try again.")
                self.console_text.insert(tk.END, f"\nThe following extra columns are present in the CSV file: {', '.join(extra_columns)}.\nPlease check the CSV file and try again.")
                self.console_text.see(tk.END)
                self.enable_buttons()
            else:
                # Loop through each row and read the values for each column
                for index, row in data.iterrows():
                    try:
                        name = row['name']
                        ip = row['ip']
                        username = row['username']
                        password = row['password']

                        # Check for missing values in row
                        if any(pd.isnull([name, ip, username, password])):
                            error_msg = f"Row {index} has blank values. Please ensure that all rows within these columns have values."
                            self.logger.error(error_msg)
                            self.console_text.insert(tk.END, f"\n{error_msg}\n")
                            self.console_text.see(tk.END)
                        else:
                            #option for multithreading
                            #threading.Thread(target=self.save_macro, args=(ip, username, password, macro_name, js_file_path, index)).start()
                            self.save_macro(ip, username, password, macro_name, js_file_path, index)
                    except KeyError as e:
                        self.logger.error(f"Missing column {e} in CSV file")
                        self.console_text.insert(tk.END, f"\nMissing column {e} in CSV file\n")
                        self.console_text.see(tk.END)
                        self.enable_buttons()
                    except Exception as e:
                        self.logger.exception(f"Exception occurred while processing row {index}: {e}")
                        self.console_text.insert(tk.END, f"\nException occurred while processing row {index}: {e}\n")
                        self.console_text.see(tk.END)
                    # Add message to be printed after loop is finished
                msg = "Macro Uploader has completed looping through all rows in the CSV file. Please refer to the logs for further details. You can now close the application or start another upload."
                self.logger.info(msg)
                self.console_text.insert(tk.END, f"\n\n{msg}\n\n")
                self.console_text.see(tk.END)
                self.enable_buttons()
                tk.messagebox.showinfo("Complete", f"{msg}")


        except pd.errors.EmptyDataError:
            self.logger.error(f"{filename} is empty")
            self.console_text.insert(tk.END, f"\n{filename} is empty\n")
            self.console_text.see(tk.END)
            self.enable_buttons()
        except Exception as e:
            self.logger.exception(f"Exception occurred while reading CSV file: {e}")
            self.console_text.insert(tk.END, f"\nException occurred while reading CSV file: {e}\n")
            self.console_text.see(tk.END)
            self.enable_buttons()


    def count_dirs_with_files(self):
        # Get the directory path from the Tkinter entry widget
        dir_path = self.otj_macro_entry.get()
        js_file_path = self.csv2_entry.get()

        if not dir_path or not js_file_path:
            messagebox.showerror("Error", "Both CSV data file and directory are required")
            self.console_text.insert(tk.END, f"\nBoth CSV data file and directory are required\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nBoth CSV data file and directory are required\n")
            return

        try:
            # Check if the directory exists
            self.console_text.insert(tk.END, f"\nScanning directory...\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nScanning directory...\n")
            if not os.path.isdir(dir_path):
                raise ValueError("Directory does not exist")

            # Get a list of all files and directories in the specified directory
            dir_contents = os.listdir(dir_path)

            # Check if any zip files exist in the specified directory
            zip_files = [f for f in dir_contents if f.endswith('.zip')]
            if not zip_files:
                raise ValueError("No zip files found in directory")

            # Count the number of directories that contain the required files
            num_dirs = 0
            for zip_file in zip_files:
                try:
                    with zipfile.ZipFile(os.path.join(dir_path, zip_file), 'r') as zip_ref:
                        zip_contents = zip_ref.namelist()
                        if 'otj-macro.js' in zip_contents and 'otj-macro-settings.js' in zip_contents:
                            num_dirs += 1
                except zipfile.BadZipFile:
                    self.console_text.insert(tk.END, f"Error: {zip_file} is not a valid zip file")
                    self.console_text.see(tk.END)
                    self.logger.error(f"Error: {zip_file} is not a valid zip file")

            self.console_text.insert(tk.END, f"\nNumber of directories containing otj-macro.js and otj-macro-settings.js: {num_dirs}\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nNumber of directories containing otj-macro.js and otj-macro-settings.js: {num_dirs}\n")
            self.show_scan_result(num_dirs)

            return num_dirs

        except Exception as e:
            self.console_text.insert(tk.END, f"Error: {str(e)}")
            self.console_text.see(tk.END)
            self.logger.exception(f"Error: {str(e)}")
            return 0

    def extract_zip_files(self):
        # Get the directory path from the Tkinter entry widget
        dir_path = self.otj_macro_entry.get()

        try:
            # Check if the directory exists
            if not os.path.isdir(dir_path):
                raise ValueError("Directory does not exist")

            # Get a list of all files and directories in the specified directory
            dir_contents = os.listdir(dir_path)

            # Check if any zip files exist in the specified directory
            zip_files = [f for f in dir_contents if f.endswith('.zip')]
            if not zip_files:
                raise ValueError("No zip files found in directory")

            # Extract the contents of each zip file into a subdirectory with the same name as the zip file
            for zip_file in zip_files:
                try:
                    with zipfile.ZipFile(os.path.join(dir_path, zip_file), 'r') as zip_ref:
                        zip_contents = zip_ref.namelist()
                        if 'otj-macro.js' in zip_contents and 'otj-macro-settings.js' in zip_contents:
                            # Create a subdirectory with the same name as the zip file
                            sub_dir = os.path.splitext(zip_file)[0]
                            sub_dir_path = os.path.join(dir_path, sub_dir)
                            os.makedirs(sub_dir_path, exist_ok=True)

                            # Extract the contents of the zip file into the subdirectory
                            zip_ref.extractall(sub_dir_path)
                except zipfile.BadZipFile:
                    self.console_text.insert(tk.END, f"Error: {zip_file} is not a valid zip file")
                    self.console_text.see(tk.END)
                    self.logger.error(f"Error: {zip_file} is not a valid zip file")

            self.console_text.insert(tk.END, "\nZip files extracted successfully\n")
            self.console_text.see(tk.END)
            self.logger.info("\nZip files extracted successfully\n")
            self.update_csv_with_macro_paths_threaded(dir_path)

        except Exception as e:
            self.console_text.insert(tk.END, f"Error: {str(e)}")
            self.console_text.see(tk.END)
            self.logger.error(f"Error: {str(e)}")


    def show_scan_result(self, count):
        result = tk.messagebox.askquestion("Scan Results", f"{count} OTJ ZIP Macro directories found. Do you want to extract these files and modify the CSV?")
        if result == 'yes':
            self.extract_zip_files()
            self.logger.info(f"{count} OTJ ZIP Macro directories found. Do you want to extract these files and modify the CSV?")
            self.logger.info("User chose Yes to extract and update the csv")
            pass
        else:
            self.logger.info("User chose No to extract and update the csv")
            pass


    def update_csv_with_macro_paths(self, dir_path):
        try:
            otj_csv_file = self.csv2_entry.get()
            # regular expression pattern to match the part of the directory name before "-otj-macro-latest"
            pattern = r"^(.*)-otj-macro-latest$"

            # read the CSV file into a DataFrame
            df = pd.read_csv(otj_csv_file)

            # loop through all subdirectories in the directory
            for subdir in os.listdir(dir_path):
                # check if the subdirectory name matches the expected pattern
                match = re.match(pattern, subdir)
                if match:
                    # get the part of the subdirectory name before "-otj-macro-latest"
                    subdir_name = match.group(1)

                    # loop through each row in the "name" column of the DataFrame
                    for index, row in df.iterrows():
                        # check if the value in the "name" column matches the subdir name
                        if row['name'] == subdir_name:
                            # get the full file path of the matching directory
                            full_path = os.path.join(dir_path, subdir)
                            # get the file names in the matching directory
                            file_names = os.listdir(full_path)
                            # loop through each file in the directory
                            for file_name in file_names:
                                # check if the file name matches the expected format
                                if file_name == 'otj-macro.js':
                                    # write the full file path to the "macro path" column for the matching row
                                    df.at[index, 'macro path'] = os.path.join(full_path, file_name)
                                    self.console_text.insert(tk.END, f"\nRow {index}. Successful: Updated 'macro path' column for the value '{row['name']}' in subdirectory '{subdir_name}' with file path '{os.path.join(full_path, file_name)}'\n")
                                    self.console_text.see(tk.END)
                                    self.logger.info(f"\nRow {index}. Successful: Updated 'macro path' column for the value '{row['name']}' in subdirectory '{subdir_name}' with file path '{os.path.join(full_path, file_name)}'\n")
                                elif file_name == 'otj-macro-settings.js':
                                    # write the full file path to the "macro settings path" column for the matching row
                                    df.at[index, 'macro settings path'] = os.path.join(full_path, file_name)
                                    self.console_text.insert(tk.END, f"\nRow {index}. Successful: Updated 'macro settings path' column for the value '{row['name']}' in subdirectory '{subdir_name}' with file path '{os.path.join(full_path, file_name)}'\n")
                                    self.console_text.see(tk.END)
                                    self.logger.info(f"\nRow {index}. Successful: Updated 'macro settings path' column for the value '{row['name']}' in subdirectory '{subdir_name}' with file path '{os.path.join(full_path, file_name)}'\n")
                            if df.loc[index, 'macro path'] and df.loc[index, 'macro settings path']:
                                self.console_text.insert(tk.END, f"\nSuccessful: The value '{row['name']}' matches the subdirectory name '{subdir_name}' in the directory '{full_path}' and the CSV file has been updated\n")
                                self.console_text.see(tk.END)
                                self.logger.info(f"\nSuccessful: The value '{row['name']}' matches the subdirectory name '{subdir_name}' in the directory '{full_path}' and the CSV file has been updated\n")
                            else:
                                self.console_text.insert(tk.END, f"\nError: Failed to write file paths for the value row {index} '{row['name']}' in the subdirectory '{subdir_name}' in the directory '{full_path}'\n")
                                self.console_text.see(tk.END)
                                self.logger.error(f"\nError: Failed to write file paths for the value row {index} '{row['name']}' in the subdirectory '{subdir_name}' in the directory '{full_path}'\n")


            # write the updated DataFrame back to the CSV file
            df.to_csv(otj_csv_file, index=False)
            messagebox.showinfo("Scan Complete", "Complete. Please check the log and CSV file for results.")
            self.logger.info("Scan Complete. Please check the log and CSV file for results.")
            self.otj_upload_macros_button.configure(state="normal")

        except Exception as e:
            self.console_text.insert(tk.END, f"\nError: {str(e)}\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nError: {str(e)}\n")
            messagebox.showerror("Error", str(e))



    def update_csv_with_macro_paths_threaded(self, dir_path):
        # create a new thread for the method
        t = threading.Thread(target=self.update_csv_with_macro_paths, args=(dir_path,))
        # start the thread
        t.start()



    def upload_macros_otj(self):
        try:
            filename = self.csv2_entry.get()
            otj_dir = self.otj_macro_entry.get()
            if not filename or not otj_dir:
                raise ValueError("Both CSV data file and directory are required")
            #Disable all buttons while the function runs
            self.otj_upload_macros_button.configure(state="disabled")
            self.csv2_browse_button.configure(state="disabled")
            self.otj_macro_browse_button.configure(state="disabled")
            self.otj_check_macros_button.configure(state='disabled')
            threading.Thread(target=self.read_csv_data_otj).start()
            self.logger.info(f"Proceeding to threaded read_csv_data_otj method")
        except ValueError as e:
            tk.messagebox.showerror("Error", str(e))
            self.console_text.insert(tk.END, f"\nError: {e}\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nError: {e}\n")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error uploading macros: {e}")
            self.console_text.insert(tk.END, f"\nError uploading macros: {e}\n")
            self.console_text.see(tk.END)
            self.logger.error(f"\nError uploading macros: {e}\n")


    def read_csv_data_otj(self):
        filename = self.csv2_entry.get()
        try:
            # Replace 'filename.csv' with the actual name of your CSV file
            data = pd.read_csv(filename, delimiter=',', header=0)
            self.console_text.insert(tk.END, f"\nMacro uploads starting... Do not close the application until complete.\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nMacro uploads starting... Do not close the application until complete.\n")
            time.sleep(1)
            self.console_text.insert(tk.END, f"\nParsing {filename} for all endpoints\n")
            self.console_text.see(tk.END)
            self.logger.info(f"\nParsing {filename} for all endpoints\n")
            time.sleep(1)

            # Check if the column names match (ignoring case)
            expected_columns = {'name', 'ip', 'username', 'password', 'macro settings path', 'macro path'}
            actual_columns = set(map(str.lower, data.columns))
            if not actual_columns.issuperset(expected_columns):
                missing_columns = expected_columns - actual_columns
                extra_columns = actual_columns - expected_columns
                error_msg = ""
                if missing_columns:
                    error_msg += f"The following columns are missing from the CSV file: {', '.join(missing_columns)}. "
                if extra_columns:
                    error_msg += f"The following extra columns are present in the CSV file: {', '.join(extra_columns)}. "
                error_msg += "Please ensure that the columns are name, ip, username, password"
                self.logger.error(error_msg)
                self.console_text.insert(tk.END, f"\n{error_msg}\n")
                self.console_text.see(tk.END)
                self.enable_buttons()
                # raise an error and stop execution
                raise ValueError(error_msg)
            elif actual_columns != expected_columns:
                # if the actual columns set is not equal to the expected columns set, 
                # but contains all the required columns, then log a warning and continue execution
                extra_columns = actual_columns - expected_columns
                self.logger.warning(f"The following extra columns are present in the CSV file: {', '.join(extra_columns)}.\nPlease check the CSV file and try again.")
                self.console_text.insert(tk.END, f"\nThe following extra columns are present in the CSV file: {', '.join(extra_columns)}.\nPlease check the CSV file and try again.")
                self.console_text.see(tk.END)
                self.enable_buttons()
            else:
                # Loop through each row and read the values for each column
                for index, row in data.iterrows():
                    try:
                        name = row['name']
                        ip = row['ip']
                        username = row['username']
                        password = row['password']
                        macro_settings_path = row['macro settings path']
                        macro_path = row['macro path']

                        if any(pd.isnull([name, ip, username, password, macro_settings_path, macro_path])):
                            error_msg = f"Row {index} has blank values. Please ensure that all rows within these columns have values."
                            self.logger.error(error_msg)
                            self.console_text.insert(tk.END, f"\n{error_msg}\n")
                            self.console_text.see(tk.END)
                        else:
                            # First loop
                            macro_name = os.path.splitext(os.path.basename(macro_settings_path))[0]
                            self.save_macro(ip, username, password, macro_name, macro_settings_path, index)

                            # Second loop
                            macro_name = os.path.splitext(os.path.basename(macro_path))[0]
                            self.save_macro(ip, username, password, macro_name, macro_path, index)
                    except KeyError as e:
                        self.logger.error(f"Missing column {e} in CSV file")
                        self.console_text.insert(tk.END, f"\nMissing column {e} in CSV file\n")
                        self.console_text.see(tk.END)
                        self.enable_buttons_otj()
                    except Exception as e:
                        self.logger.exception(f"Exception occurred while processing row {index}: {e}")
                        self.console_text.insert(tk.END, f"\nException occurred while processing row {index}: {e}\n")
                        self.console_text.see(tk.END)
                    # Add message to be printed after loop is finished
                msg = "Macro Uploader has completed looping through all rows in the CSV file. Please refer to the logs for further details. You can now close the application or start another upload."
                self.logger.info(msg)
                self.console_text.insert(tk.END, f"\n\n{msg}\n\n")
                self.console_text.see(tk.END)
                self.enable_buttons_otj()
                tk.messagebox.showinfo("Complete", f"{msg}")

        except pd.errors.EmptyDataError:
            self.logger.error(f"{filename} is empty")
            self.console_text.insert(tk.END, f"\n{filename} is empty\n")
            self.console_text.see(tk.END)
            self.enable_buttons_otj()
        except Exception as e:
            self.logger.exception(f"Exception occurred while reading CSV file: {e}")
            self.console_text.insert(tk.END, f"\nException occurred while reading CSV file: {e}\n")
            self.console_text.see(tk.END)
            self.enable_buttons_otj()

       
    def enable_buttons(self):
        time.sleep(1)
        self.upload_macros_button.configure(state="normal")
        self.csv_browse_button.configure(state="normal")
        self.macro_browse_button.configure(state="normal")

    def enable_buttons_otj(self):
        time.sleep(1)
        self.otj_upload_macros_button.configure(state="normal")
        self.csv2_browse_button.configure(state="normal")
        self.otj_macro_browse_button.configure(state="normal")
        self.otj_check_macros_button.configure(state='normal')


    def on_closing(self):
        # Do any cleanup or save data here before closing the window
        self.destroy() 
        

# Instantiate and run the class
if __name__ == "__main__":
    app = PexipMacroApp()
    app.mainloop()
