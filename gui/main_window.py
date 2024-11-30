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

FILE_WIDTH = 50
FILE_HEIGHT = 50

DASH = (5,2)

WHITE = "#F3F3F3"
BLACK = "black"
LIGHT_GREY = "#e0e0e0"
GREY = "#d0d0d0"
DARK_GREY = "#b3b3b3"

# Orientation constants
READONLY = "readonly"
NORMAL = "normal"
RIGHT = "right"
LEFT = "left"
BOTTOM = "bottom"
BOTH = "both"
class MainWindow(ctk.CTk):
    '''Main window of the app. Redirects to all the other features'''
    def __init__(self):
        super().__init__()
        self.title("Pdf Converter")
        self.resizable(False, False) # NOT Resizable
        self.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        ctk.set_appearance_mode("System")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("dark-blue")

        # Class variable needed
        self.paths_to_convert = []

        # Configure the main window properties
        self.title("PDF Converter App")
        self.geometry("800x500")
        
        self.mainFrame = ctk.CTkFrame(self, fg_color=WHITE)
        self.mainFrame.pack(expand=True, fill="both")

        # Title label
        self.TitleLabel = ctk.CTkLabel(
            self.mainFrame, 
            text="PDF Converter", 
            text_color="black",
            font=("Arial", 24))
        self.TitleLabel.pack(pady=20)

        # Create a frame for the file selection
        self.FileFrame = ctk.CTkFrame(self.mainFrame, fg_color=WHITE)  # Light blue background for the frame
        self.FileFrame.pack(pady=10, padx=20, fill="x")

        # Create a label for file path selection
        self.FileLabel = ctk.CTkLabel(self.FileFrame, 
                                      text="Select a file:", 
                                      text_color=BLACK,
                                      )
        self.FileLabel.pack(side="left", padx=(0, 10))

        # Create an entry field for the selected file path
        self.FileEntry = ctk.CTkEntry(
            self.FileFrame, 
            placeholder_text="Choose a file", 
            fg_color="#FFFFFF",  # White background for the entry field
            border_color="#A1C2F1",  # Light blue border color
            text_color="#2C2C54"  # Deep blue text
        )
        self.FileEntry.pack(side="left", fill="x", expand=True)

        # Create a button for browsing files
        self.BrowseButton = ctk.CTkButton(
            self.FileFrame, 
            text="Browse", 
            fg_color="#4A90E2",  # Bright blue background
            hover_color="#357ABD",  # Slightly darker blue when hovered
            text_color="#FFFFFF",  # White text
            command=self.open_file
        )
        self.BrowseButton.pack(side="right", padx=8)

        # Create a convert button
        self.ConvertButton = ctk.CTkButton(
            self.mainFrame, 
            text="Convert PDF", 
            width=15,
            fg_color="#50C878",  # Green background
            hover_color="#3A9A5D",  # Darker green when hovered
            text_color="#FFFFFF",
            command=self.convert_to_pdf
        )
        self.ConvertButton.pack(pady=10)

        # Create a quit button
        self.QuitButton = ctk.CTkButton(
            self.mainFrame, 
            text="Quit", 
            command=self.quit,
            fg_color="#D9534F",  # Red background
            hover_color="#C9302C",  # Darker red when hovered
            text_color="#FFFFFF"  # White text
        )
        self.QuitButton.pack(pady=10)

        # Create a status label
        self.StatusLabel = ctk.CTkLabel(
            self.mainFrame, 
            text="Status: Ready", 
            font=("Arial", 12),
            text_color=BLACK
        )
        self.StatusLabel.pack(pady=10)

        # File Preview Frame
        self.fileShowingFrame = ctk.CTkFrame(master=self.mainFrame, corner_radius=10, fg_color=LIGHT_GREY)
        self.fileShowingFrame.pack(pady=20, padx=20, fill="both", expand=True)

        # # File Canvas
        # self.fileCanvas = Canvas(master=self.fileShowingFrame, width=FILE_WIDTH, height=FILE_HEIGHT, bg=GREY, highlightthickness=0)
        # self.fileCanvas.pack(pady=10)

        # File Image Placeholder
        self.fileImageLabel = ctk.CTkLabel(
            master=self.fileShowingFrame,
            text="Upload File Here",
            text_color=BLACK,
            width=FILE_WIDTH,
            height=FILE_HEIGHT,
        )
        self.fileImageLabel.pack()

        # File Path Label
        self.filePathLabel = ctk.CTkLabel(
            master=self.fileShowingFrame,
            text="",
            font=("Helvetica", 14),
            text_color=BLACK,
        )
        self.filePathLabel.pack(pady=10)

        # Error Label
        self.errorLabel = ctk.CTkLabel(
            master=self.mainFrame,
            text="",
            font=("Helvetica", 12),
            text_color="#FF4B4B",
            width=WIN_WIDTH,
        )
        self.errorLabel.pack(pady=10)

        # Progress Bar frame
        self.progressFrame = ctk.CTkFrame(self.mainFrame, fg_color=WHITE)
        self.progressFrame.pack(side=BOTTOM)

        # Progress bar
        self.progressBar = ctk.CTkProgressBar(
            master=self.progressFrame,
            width=WIN_WIDTH - 20,  # Slightly smaller than window width for padding
            height=10,            # Thin progress bar
            progress_color="#4CAF50",  # Light green color
            fg_color=DARK_GREY,    # Slightly darker grey background
            determinate_speed=2,
        )
        self.progressBar.set(0)  # Initial progress set to 0
        self.progressBar.grid(row=0, column=0)  # Positioned at the bottom
        
        # Progress bar percentage label
        self.progressBarLabel = ctk.CTkLabel(self.progressFrame, 
                                             text="0%",
                                             text_color=BLACK,
                                             bg_color=WHITE)
        self.progressBarLabel.grid(row=0, column=1, padx=10)

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
            
            # 
            self.fileImageLabel.pack_forget()
        
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
        self.reset_after_conversion()
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
                if i == 0: good_message = f"{i + 1} Conversion done successfully"
                else: good_message = f"{i + 1} Conversions done successfully"
                self.after(0, self.update_progress, (i + 1) * progress_step, good_message)

            except Exception as error:
                bad_message = f"Error Converting {file_path}: {error}"
                self.after(0, self.update_error, bad_message)

        self.paths_to_convert.clear()

    def update_progress(self, progress, message):
        """Updates the progress bar and success message."""
        
        self.progressBar.set(progress)
        self.progressBarLabel.configure(text=f"{int(progress * 100)}%")
        self.errorLabel.configure(text=message, text_color="#4CAF50")

    def update_error(self, message):
        """Updates the error label."""
        self.errorLabel.configure(text=message, text_color="#FF4B4B")

    def reset_after_conversion(self):
        """Resets paths and progress bar after all conversions are done."""
        self.progressBar.set(0)
        self.errorLabel.configure(text="")
        self.progressBarLabel.configure(text="0%")




        



        


        
