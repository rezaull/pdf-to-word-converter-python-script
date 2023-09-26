import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
import sys
from pdf2docx import Converter

# Get the path to the directory containing the script or executable
if getattr(sys, 'frozen', False):
    # Executable path in PyInstaller
    base_path = sys._MEIPASS
else:
    # Script path
    base_path = os.path.dirname(os.path.abspath(__file__))

# Global variables
input_files = []
output_folder = ""

# Function to handle file selection
def select_files():
    global input_files
    input_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    files_label.config(text=f"Selected Files: {', '.join(input_files)}")

# Function to handle output folder selection
def select_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory()
    output_label.config(text=f"Output Folder: {output_folder}")

# Function to convert PDF to Word
def convert_to_word():
    global input_files, output_folder

    if input_files and output_folder:
        try:
            # Create the output folder if it doesn't exist
            os.makedirs(output_folder, exist_ok=True)

            # Calculate the total number of files
            total_files = len(input_files)

            # Iterate through each input file
            for i, input_file in enumerate(input_files):
                # Get the base file name
                file_name = os.path.basename(input_file)
                # Remove the file extension
                file_name = os.path.splitext(file_name)[0]
                # Create the output file path with the same name as the input file
                output_file = os.path.join(output_folder, f"{file_name}.docx")

                # Create a Converter object
                cv = Converter(input_file)

                # Convert PDF to Word document
                cv.convert(output_file)

                # Close the Converter object
                cv.close()

                # Calculate the progress percentage
                progress_percentage = (i + 1) / total_files * 100

                # Update the progress label
                progress_label.config(text=f"Conversion Progress: {int(progress_percentage)}%")
                window.update()

            messagebox.showinfo("Conversion Complete", "PDF to Word conversion completed successfully!")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    else:
        messagebox.showwarning("Warning", "Please select input files and output folder.")

# Create the main window
window = tk.Tk()
window.title("PDF to Word Converter v1.0")
window.geometry("400x440")
# Set the window icon
icon_path = os.path.join(base_path, "windowicon.ico")
window.iconbitmap(icon_path)

# Set the window background to light blue
window.config(bg="#E1F5FE")

# Title
output_label = tk.Label(window, text="PDF to Word Converter", bg="#E1F5FE", font=("Arial", 16), fg="green")
output_label.pack()

# Load the logo image
logo_image = tk.PhotoImage(file=os.path.join(base_path, "logo.png"))  # Logo into the window. Replace "logo.png" with the path to your logo image
top_frame = tk.Frame(window, pady=10)  # Adjust the padding value (in pixels) to add more or less space
# Resize the image
logo_image = logo_image.subsample(1, 1) 

# Create a label widget to display the logo
logo_label = tk.Label(window, image=logo_image, bg="#E1F5FE",)
logo_label.pack()
top_frame.pack(side=tk.TOP)

# Create a file selection button
file_button = ttk.Button(window, text="Select PDF Files", command=select_files, style="Custom1.TButton")
file_button.pack(pady=10)

# Create a label to display the selected file paths
files_label = tk.Label(window, text="Selected Files: None", bg="#E1F5FE")
files_label.pack()

# Create an output folder selection button
output_button = ttk.Button(window, text="Select Output Folder", command=select_output_folder, style="Custom1.TButton")
output_button.pack(pady=10)

# Create a label to display the output folder path
output_label = tk.Label(window, text="Output Folder: None", bg="#E1F5FE")
output_label.pack()

# Create a convert button
convert_button = ttk.Button(window, text="Convert", command=convert_to_word, style="Custom.TButton")
convert_button.pack(pady=10)

# Create a progress label
progress_label = tk.Label(window, text="Conversion Progress: 0%", bg="#E1F5FE", font=("Arial", 12))
progress_label.pack()

# Create a footer with credits
footer_label = tk.Label(window, text="Developed by Reza", bg="#E1F5FE", fg="#b1b8ba")
footer_label.pack(pady=10)

# Define custom button style
style = ttk.Style()
style.theme_use("clam")
style.configure("Custom.TButton",
                background="#4caf50",
                foreground="#ffffff",
                relief="flat",
                padding=10,
                font=("Arial", 14),
                width=20)
style.configure("Custom1.TButton",
                background="#039aab",
                foreground="#ffffff",
                relief="flat",
                padding=10,
                font=("Arial", 14),
                width=20)


# Apply hover effect
style.map("Custom.TButton",
          background=[("active", "#45a049")],
          relief=[("hover", "groove"), ("active", "sunken")])
style.map("Custom1.TButton",
          background=[("active", "#04818f")],
          relief=[("hover", "groove"), ("active", "sunken")])

# Center the window on the desktop
window.update_idletasks()
window_width = window.winfo_width()
window_height = window.winfo_height()
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
window.geometry(f"+{x}+{y}")

# Start the Tkinter event loop
window.mainloop()
