# import tkinter as tk
# import tkinter.filedialog
# import requests
# import customtkinter
# from tkinter import messagebox
# from threading import Thread
# import time
# import os
# from excel_data_extractor_test import main
#
# customtkinter.set_appearance_mode("light")
# customtkinter.set_default_color_theme('blue')
#
# def long_running_task(folder, file, status_label, execute_button, result_label):
#     """ Run the main function and update the status label. """
#     try:
#         # Initialize status
#         status_label.configure(text="Automation Started...")
#         status_label.update_idletasks()
#
#         # Call the actual main function
#         result = main(folder, file)
#
#         # Simulate progress update (if actual progress is known, update accordingly)
#         for i in range(101):
#             time.sleep(0.01)  # Simulate work by sleeping
#             status_label.configure(text=f"Progress: {i}%")
#             status_label.update_idletasks()
#
#         # When task is complete
#         status_label.configure(text="Automation Completed!")
#
#         # Update the result label and open button
#         # result_label.configure(text=f"Processed file path or message: {result}")
#
#     except Exception as e:
#         status_label.configure(text=f"Error: {str(e)}")
#
#     finally:
#         execute_button.configure(state=customtkinter.NORMAL)
#
# def start_reader():
#     folder = entry.get()
#     file = entry_1.get()
#     if folder == 'Folder' or folder == '':
#         messagebox.showerror('Error', 'Please select a folder to execute')
#     else:
#         # Disable the execute button to prevent multiple clicks
#         execute_button.configure(state=customtkinter.DISABLED)
#         # Start the long-running task in a separate thread
#         Thread(target=long_running_task, args=(folder, file, status_label, execute_button, result_label),
#                daemon=True).start()
#
# def browse_folder(entry):
#     folder_path = tkinter.filedialog.askdirectory()
#     entry.delete(0, 'end')
#     entry.insert(0, folder_path)
#
# def browse_file_1(entry_1):
#     file_path = tkinter.filedialog.askopenfilename()
#     entry_1.delete(0, 'end')
#     entry_1.insert(0, file_path)
#
# def open_file():
#     file_path = result_label.cget("text").split(": ", 1)[-1].strip()
#     if os.path.exists(file_path):
#         try:
#             os.startfile(file_path)
#         except Exception as e:
#             messagebox.showerror("Error", f"Failed to open file: {e}")
#     else:
#         messagebox.showerror("Error", "File does not exist or path is incorrect")
#
# root = customtkinter.CTk()
# root.title('Accumark Automation')
#
# # Title Label
# label = customtkinter.CTkLabel(master=root, text="Accumark File Reader", font=('Helvetica', 20))
# label.place(relx=0.5, rely=0.1, anchor=tkinter.N)
#
# # Folder Entry
# entry = customtkinter.CTkEntry(master=root, width=300, height=25, placeholder_text="Folder")
# entry.place(relx=0.38, rely=0.2, anchor=tkinter.N)
#
# # Excel File Entry
# entry_1 = customtkinter.CTkEntry(master=root, width=300, height=25, placeholder_text="Excel file")
# entry_1.place(relx=0.38, rely=0.25, anchor=tkinter.N)
#
# # Browse Buttons
# browse_button_folder = customtkinter.CTkButton(master=root, text="Browse", command=lambda: browse_folder(entry),
#                                                width=120, height=25, border_width=0, corner_radius=8)
# browse_button_folder.place(relx=0.82, rely=0.2, anchor=tkinter.N)
#
# browse_button_file = customtkinter.CTkButton(master=root, text="Browse", command=lambda: browse_file_1(entry_1),
#                                              width=120, height=25, border_width=0, corner_radius=8)
# browse_button_file.place(relx=0.82, rely=0.25, anchor=tkinter.N)
#
# # Execute Button
# execute_button = customtkinter.CTkButton(master=root, text="Execute", command=start_reader, width=430, height=25,
#                                          border_width=0, corner_radius=8)
# execute_button.place(relx=0.5, rely=0.35, anchor=tkinter.N)
#
# # Status Label
# status_label = customtkinter.CTkLabel(master=root, text="Ready to Start")
# status_label.place(relx=0.5, rely=0.4, anchor=tkinter.N)
#
# # Result Label
# result_label = customtkinter.CTkLabel(master=root, text="")
# result_label.place(relx=0.5, rely=0.48, anchor=tkinter.N)
#
#
# root.geometry("500x700")
#
# # Check connection status
# try:
#     cond_AT = requests.get("https://saim2481.pythonanywhere.com/ATactivation-desktop-response/")
#     cond_AT.raise_for_status()
#     cond_AT = cond_AT.text
#     if cond_AT != "true":
#         execute_button.configure(state=customtkinter.DISABLED)
# except requests.exceptions.RequestException:
#     messagebox.showerror("Connection Error", "Please Check your internet connection")
# except Exception as e:
#     messagebox.showerror("Something Went Wrong", f"Unexpected Error: {e}")
#     execute_button.configure(state=customtkinter.DISABLED)
#
# root.mainloop()
#

from flask import Flask, request, redirect, url_for, render_template, flash, jsonify, send_file, abort
import os
import zipfile
from werkzeug.utils import secure_filename
from excel_data_extractor_test import main
import threading

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.getcwd()  # Save files to the current working directory
app.config['ALLOWED_EXTENSIONS'] = {'zip', 'xls', 'xlsx'}
app.secret_key = 'falsdjfjaklsdfjalksjdffffhhhh78454ddaawfvc'  # For flashing messages

# Global variables to store execution status and result file path
execution_status = {
    'started': False,
    'completed': False,
    'error': None,
    'result_file': None
}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def unzip_file(zip_path, extract_to_folder):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to_folder)

def run_long_task(folder, excel_file_path):
    global execution_status
    try:
        execution_status['started'] = True
        # Call the main function and get the updated file path
        result_file_path = main(folder+"\\files", excel_file_path)

        if result_file_path:
            execution_status['result_file'] = result_file_path
            execution_status['completed'] = True
            execution_status['error'] = None
        else:
            execution_status['error'] = 'Failed to process files'
    except Exception as e:
        execution_status['error'] = str(e)
    finally:
        execution_status['started'] = False


@app.route('/')
def index():
    return render_template('index.html', execution_status=execution_status)

@app.route('/upload', methods=['POST'])
def upload_files():
    uploaded_zip = request.files.get('folder')
    excel_file = request.files.get('file')

    if not uploaded_zip or not excel_file:
        flash('No ZIP file or Excel file uploaded')
        return redirect(request.url)

    if not allowed_file(uploaded_zip.filename) or not allowed_file(excel_file.filename):
        flash('Invalid file type')
        return redirect(request.url)

    # Save the ZIP file
    zip_filename = secure_filename(uploaded_zip.filename)
    zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
    uploaded_zip.save(zip_path)

    # Create a folder to extract the ZIP contents
    extract_folder = os.path.join(app.config['UPLOAD_FOLDER'], os.path.splitext(zip_filename)[0])
    if not os.path.exists(extract_folder):
        os.makedirs(extract_folder)

    # Unzip the file
    unzip_file(zip_path, extract_folder)

    # Save the Excel file
    excel_filename = secure_filename(excel_file.filename)
    excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    excel_file.save(excel_file_path)

    # Start the long-running task in a separate thread
    threading.Thread(target=run_long_task, args=(extract_folder, excel_file_path), daemon=True).start()

    flash('Processing started. Please check the status for updates.')
    return redirect(url_for('index'))

@app.route('/status')
def status():
    return jsonify(execution_status)

@app.route('/download')
def download_file():
    if execution_status['result_file']:
        try:
            return send_file(execution_status['result_file'], as_attachment=True)
        except Exception as e:
            print(f"Error sending file: {e}")
            abort(404)
    else:
        flash('No result file available for download.')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)