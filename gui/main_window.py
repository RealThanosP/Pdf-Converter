import customtkinter as ctk
from tkinter import (filedialog, END, Canvas)
from pdflib import word2pdf
from pathlib import WindowsPath, PosixPath
import os.path
from PIL import Image
import threading

# Size and style constants
WIN_WIDTH = 600
WIN_HEIGHT = 500

FILE_WIDTH = 100
FILE_HEIGHT = 100

DASH = (5,2)

WHITE = "white"

# Orientation constants
READONLY = "readonly"
NORMAL = "normal"

class MainWindow(ctk.CTk):
    '''Main window of the app. Redirects to all the other features'''
    def __init__(self):
        super().__init__()
        self.title("Pdf Manager")
        self.resizable(False, False) # NOT Resizable
        self.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        ctk.set_appearance_mode("System")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("dark-blue")

        # Main Frame
        self.mainFrame = ctk.CTkFrame(master=self, width=WIN_WIDTH, height=WIN_HEIGHT, corner_radius=15)
        self.mainFrame.pack(pady=10, padx=10, fill="both", expand=True)

        # App Title
        self.titleLabel = ctk.CTkLabel(
            master=self.mainFrame,
            text="PDF Manager",
            font=("Helvetica", 24, "bold"),
            text_color="#FFFFFF",
        )
        self.titleLabel.pack(pady=20)

        # Open DOCX button
        self.openDocxButton = ctk.CTkButton(
            master=self.mainFrame,
            text="Open Docx File",
            font=("Helvetica", 16),
            command=self.open_file,
        )
        self.openDocxButton.pack(pady=10)

        # File Preview Frame
        self.fileFrame = ctk.CTkFrame(master=self.mainFrame, corner_radius=10, fg_color="#2A2A2A")
        self.fileFrame.pack(pady=20, padx=20, fill="both", expand=True)

        # File Canvas
        self.fileCanvas = Canvas(master=self.fileFrame, width=FILE_WIDTH, height=FILE_HEIGHT, bg="#3E3E3E", highlightthickness=0)
        self.fileCanvas.pack(pady=10)

        # File Image Placeholder
        self.fileImageLabel = ctk.CTkLabel(
            master=self.fileCanvas,
            text="Upload File Here",
            text_color="grey",
            width=FILE_WIDTH,
            height=FILE_HEIGHT,
        )
        self.fileImageLabel.pack()

        # File Path Label
        self.paths_to_convert = []
        self.filePathLabel = ctk.CTkLabel(
            master=self.fileFrame,
            text="",
            font=("Helvetica", 14),
            text_color="#FFFFFF",
        )
        self.filePathLabel.pack(pady=10)

        # Convert Button
        self.convertButton = ctk.CTkButton(
            master=self.mainFrame,
            text="Convert to PDF",
            font=("Helvetica", 16),
            fg_color="#FF4B4B",
            hover_color="#FF6B6B",
            command=self.convert_to_pdf,
        )
        self.convertButton.pack(pady=10)

        # Error Label
        self.errorLabel = ctk.CTkLabel(
            master=self.mainFrame,
            text="",
            font=("Helvetica", 12),
            text_color="#FF4B4B",
            width=WIN_WIDTH,
        )
        self.errorLabel.pack(pady=10)

        # Add this to the `__init__` method of MainWindow, at the bottom:
        self.progressBar = ctk.CTkProgressBar(
            master=self.mainFrame,
            width=WIN_WIDTH - 20,  # Slightly smaller than window width for padding
            height=10,            # Thin progress bar
            progress_color="#4CAF50",  # Light green color
            fg_color="#F0F0F0"     # Slightly darker grey background
        )
        self.progressBar.set(0)  # Initial progress set to 0
        self.progressBar.pack(pady=10, side="bottom")  # Positioned at the bottom


    # TODO: Load the image of the file in a box inside the open_file method
    def open_file(self):
        """Opens file dialog and changes the the self.loadBox to hold an image of the file"""

        # Define file types
        filetypes = (
            ("Word Files", "*.docx"),
            ("Word Files", "*.doc"),
            ("PDF Files", "*.pdf"),
            ("All Files", "*.*")
        )

        # Show file dialog
        file_paths = filedialog.askopenfilenames(filetypes=filetypes)

        if file_paths:
            # Define the image object
            file_image = ctk.CTkImage(light_image=Image.open("gui/assets/doc.png"),
                                    dark_image=Image.open("gui/assets/doc.png"),
                                    size=(FILE_WIDTH, FILE_HEIGHT))
            
            # Add the image to the canvas if file_paths is not empty
            self.fileImageLabel.configure(image=file_image, text="")
        
 
        # Save the paths_to_convert to label
        end_path_list = []
        for i, path in enumerate(file_paths):
            path_obj = WindowsPath(path)

            self.paths_to_convert.append(path_obj)

            if path_obj.is_dir():
                continue

            end_path_list.append(f"({i + 1}) {path_obj.name}")
        
        end_text = "\n".join(end_path_list)

        self.filePathLabel.configure(text=end_text)


    # TODO: Make the convert_to_pdf compatible for Linux and MACOS
    # TODO: Handle the errors in the convert_word_to_pdf Function (catch Exceptions fr)
    # TODO: Make the convert_word_to_pdf function compatible with macos
    #* TODO: Make the convert_to_pdf convert more than 1 file DONE
    def convert_to_pdf(self):
        '''Gets the path of the file to conversion by the widget on the screen and 
            converts to pdf to the same folder as the source file.'''

        file_paths = self.paths_to_convert

        if not file_paths:
            self.errorLabel.configure(text="Please enter a path to convert to pdf", text_color="#FF4B4B")
            return
        
        # Runs the conversion on multi-threading to update the progressBar in parallel. Makes the app FEEL? fast. It's not...
        conversion_thread = threading.Thread(target=self.run_conversion_docx_to_pdf, daemon=True)
        conversion_thread.start()


    def run_conversion_docx_to_pdf(self):
        ''' Runs the actual 2pdf conversion'''
        file_paths = self.paths_to_convert
        total_files:int = len(file_paths)
        progress_step = 1 / total_files

        for i, file_path in enumerate(file_paths):
            try: 
                root_path, file = os.path.split(file_path)
                file_name, extension = os.path.splitext(file)

                # Initialize a string with the same name and the .pdf extension
                new_extension = ".pdf"
                new_file_name = file_name + new_extension

                # Have the file in the same as the initial file
                new_file_path  = WindowsPath(root_path, new_file_name)

                # Actual Conversion
                word2pdf.convert_word_to_pdf(fr"{file_path}", fr"{new_file_path}")
                
                # Update the progressBar
                good_message = f"{i + 1} Conversions done successfully"
                self.after(0, self.update_progress, (i + 1) * progress_step, good_message)

            except Exception as error:
                bad_message = f"Error Converting {file_path}: {error}"
                self.after(0, self.update_error, bad_message)

        # At last clear the paths_to_convert
        self.after(0, self.reset_after_conversion)

    def smooth_progress(self, target_progress):
        """Smoothly increments the progress bar to a target value."""

        current_progress = self.progressBar.get()
        step = 0.01  # Adjust for smoother or faster animation

        if current_progress < target_progress:
            self.progressBar.set(current_progress + step)
            self.after(50, lambda: self.smooth_progress(target_progress))  # Recursive call

    def update_progress(self, progress, message):
        """Updates the progress bar and success message."""
        self.progressBar.set(progress)
        self.errorLabel.configure(text=message, text_color="#4CAF50")

    def update_error(self, message):
        """Updates the error label."""
        self.errorLabel.configure(text=message, text_color="#FF4B4B")

    def reset_after_conversion(self):
        """Resets paths and progress bar after all conversions are done."""
        self.paths_to_convert.clear()
        self.progressBar.set(0)





        



        


        
