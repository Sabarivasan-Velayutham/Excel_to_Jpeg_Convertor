# Excel to Image Converter

The Excel to Image Converter is a Python application that allows users to convert Excel sheets into images. This can be useful for various purposes such as sharing data in image format or creating previews of Excel files.

## Features

- Select a folder containing Excel sheets.
- Select a destination folder for saving the converted images.
- Convert each Excel sheet to an image (JPEG format).
- Automatically create a backup of the original Excel sheets after conversion.

## How to Use

1. Run the `excel_to_image_converter.py` script.
2. Click on the "Select Source Folder" button to choose the folder containing Excel sheets.
3. Click on the "Select Destination Folder" button to choose the folder where the converted images will be saved.
4. Click on the "Convert Excel to Images" button to start the conversion process.
5. Once the conversion is complete, a message will be displayed indicating that all Excel sheets have been converted to images.
6. The original Excel sheets will be moved to a backup folder, which will be created if it does not already exist.

## Requirements

- Python 3.x
- tkinter (Python's standard GUI library)
- Pillow (Python Imaging Library, used for image processing)
- JPype (Java integration library)
- Aspose.Cells for Java (Java library for working with Excel files)

## Installation

Install required Python packages using pip:
```pip install pyinstaller tkinter pillow jpype aspose.cells```

Create the .exe file :
```pyinstaller --onefile -w 'gui.py'```    
```pyinstaller gui.spec```

