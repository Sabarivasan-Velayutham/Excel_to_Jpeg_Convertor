import tkinter as tk
from tkinter import filedialog, messagebox
import shutil 
import os
from jpype import *

__cells_jar_dir__ = os.path.dirname(__file__)
addClassPath(os.path.join(__cells_jar_dir__, "aspose-cells-24.4.jar"))
addClassPath(os.path.join(__cells_jar_dir__, "bcprov-jdk15on-1.68.jar"))
addClassPath(os.path.join(__cells_jar_dir__, "bcpkix-jdk15on-1.68.jar"))
addClassPath(os.path.join(__cells_jar_dir__, "JavaClassBridge.jar"))


import  asposecells 
from PIL import Image 
startJVM()     
from asposecells.api import Workbook

# Define colors
bg_color = "#f0f0f0"
btn_color = "#4caf50"
btn_text_color = "#ffffff"
label_text_color = "#333333"
font_name = "Arial"

# Define font styles
label_font = (font_name, 12)
button_font = (font_name, 12, "bold")

# Variables to hold the source and destination folder paths
source_folder = ""
destination_folder = ""
backup_folder = ""

def select_source_folder():
    global source_folder, source_label
    # Use filedialog to select the source folder
    source_folder = filedialog.askdirectory(title="Select the folder containing Excel sheets")
    if source_folder:
        # Update the source folder label with the selected folder path
        source_label.config(text=f"Source Folder: {source_folder}", bg=bg_color, fg=label_text_color)

def select_destination_folder():
    global destination_folder, dest_label
    # Use filedialog to select the destination folder
    destination_folder = filedialog.askdirectory(title="Select the destination folder for images")
    if destination_folder:
        # Update the destination folder label with the selected folder path
        dest_label.config(text=f"Destination Folder: {destination_folder}", bg=bg_color, fg=label_text_color)

def select_backup_folder():
    global backup_folder, backup_label
    # Use filedialog to select the backup folder
    backup_folder = filedialog.askdirectory(title="Select the backup folder for Excel sheets")
    if backup_folder:
        # Update the backup folder label with the selected folder path
        backup_label.config(text=f"Backup Folder: {backup_folder}", bg=bg_color, fg=label_text_color)

def process_excel_sheets():
    # Define trash folder to store intermediate JPEG files
    trash_folder = "trash"
    os.makedirs(trash_folder, exist_ok=True)

    if not source_folder :
        # Show a warning message if either folder is not selected
        messagebox.showwarning("Warning", "Please select source folder ")
        return

    if not destination_folder:
        # Show a warning message if either folder is not selected
        messagebox.showwarning("Warning", "Please select destination folder ")
        return

    if not backup_folder:
        # Show a warning message if either folder is not selected
        messagebox.showwarning("Warning", "Please select backup folder ")
        return
    
    # Create a list of Excel files in the source folder
    excel_files = [f for f in os.listdir(source_folder) if f.endswith(".xlsx") or f.endswith(".xls")]

    # Process each Excel file in the folder
    for excel_file in excel_files:
        # Construct the full file path for the Excel file
        excel_file_path = os.path.join(source_folder, excel_file)
        
        # Create a Workbook instance and load the Excel file
        workbook = Workbook(excel_file_path)
        
        # Save the workbook as an image (JPEG format)
        output_image_path = os.path.join(trash_folder, f"{os.path.splitext(excel_file)[0]}.jpg")
        workbook.save(output_image_path)
        
        # Open the saved image
        im = Image.open(output_image_path)
        
        # Get the size of the image
        width, height = im.size
        
        # Define the cropping region
        left = 0
        top = 30
        right = width
        bottom = height * 0.5
        
        # Crop the image
        im1 =im.crop((left, top, right, bottom))
        
        # Resize the cropped image
        # newsize = (width,height)
        # im1 = im1.resize(newsize)
        
        # Save the cropped and resized image in the destination folder
        filename = f"{os.path.splitext(excel_file)[0]}.jpg" 
        output_path = os.path.join(destination_folder, filename)
        im1.save(output_path)

        im1.show()

    # Show a completion message once all Excel files have been processed
    messagebox.showinfo("Conversion Complete", "All Excel sheets have been converted to images.")

    # Move Excel files from the source folder to the backup folder
    for excel_file in excel_files:
        source_path = os.path.join(source_folder, excel_file)
        backup_path = os.path.join(backup_folder, excel_file)
        # Move the file from source folder to backup folder
        shutil.move(source_path, backup_path)

    # Clean up the trash folder
    shutil.rmtree(trash_folder, ignore_errors=True)
    
    # Shutdown the JVM
    # jpype.shutdownJVM()

def main():
    # Create a tkinter window
    root = tk.Tk()
    root.title("Excel to Image Converter")

    # Set window background color
    root.configure(bg=bg_color)

    # Increase the window size
    root.geometry("600x400")

    # Create buttons for selecting the source and destination folders
    source_button = tk.Button(root, text="Select Source Folder", command=select_source_folder,
                                    bg=btn_color, fg=btn_text_color, font=button_font)
    source_button.pack(pady=10, fill=tk.X)

    destination_button = tk.Button(root, text="Select Destination Folder", command=select_destination_folder,
                                    bg=btn_color, fg=btn_text_color, font=button_font)
    destination_button.pack(pady=10, fill=tk.X)

    backup_button = tk.Button(root, text="Select Backup Folder", command=select_backup_folder,
                                    bg=btn_color, fg=btn_text_color, font=button_font)
    backup_button.pack(pady=10, fill=tk.X)


    label_font = (font_name, 12)
    # Create labels to display the selected folders
    global source_label, dest_label, backup_label
    source_label = tk.Label(root, text="Source Folder: Not Selected", wraplength=500, bg=bg_color, fg=label_text_color, font=label_font)
    source_label.pack(pady=5)

    dest_label = tk.Label(root, text="Destination Folder: Not Selected", wraplength=500, bg=bg_color, fg=label_text_color, font=label_font)
    dest_label.pack(pady=5)

    backup_label = tk.Label(root, text="Backup Folder: Not Selected", wraplength=500, bg=bg_color, fg=label_text_color, font=label_font)
    backup_label.pack(pady=5)

    # Create a button for starting the Excel to JPEG conversion
    convert_button = tk.Button(root, text="Convert Excel to Images", command=process_excel_sheets,
                           bg=btn_color, fg=btn_text_color, font=button_font)    
    convert_button.pack(pady=10, fill=tk.X)

    # Start the tkinter main loop
    root.mainloop()

if __name__ == "__main__":
    main()
