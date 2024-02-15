import pyautogui
from docx import Document
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import time

# Function to type text onto the web app
def type_text(text):
    try:
        chunk_size = 100  # Adjust the chunk size as needed
        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
        for chunk in chunks:
            pyautogui.typewrite(chunk)
            time.sleep(0.2)  # Adjust the sleep time as needed
    except Exception as e:
        print(f"An error occurred while typing: {e}")

# Function to read text from a Word document
def read_word_document(file_path):
    document = Document(file_path)
    text = ''
    for paragraph in document.paragraphs:
        text += paragraph.text + '\n'
    return text

# Function to handle file selection
def select_file():
    root = tk.Tk()
    root.title("Select File")
    root.geometry("300x150")  # Set dimensions
    root.configure(bg="white")  # Set background color

    selected_file = tk.StringVar()

    def browse_file():
        file_path = filedialog.askopenfilename(title="Select Word Document", filetypes=[("Word files", "*.docx")])
        if file_path:
            selected_file.set(file_path)

    browse_button = ttk.Button(root, text="Browse", command=browse_file)
    browse_button.pack(pady=10)

    def select():
        nonlocal root
        root.destroy()  # Close the window

    select_button = ttk.Button(root, text="Select", command=select)
    select_button.pack()

    root.mainloop()

    file_path = selected_file.get()

    if file_path:
        return file_path
    else:
        print("No file selected.")
        return None

# Function to handle window selection
def select_window():
    root = tk.Tk()
    root.title("Select Window")
    root.geometry("300x150")  # Set dimensions
    root.configure(bg="white")  # Set background color

    window_titles = [window.title for window in pyautogui.getAllWindows()]

    window_selection = tk.StringVar()
    window_selection.set(window_titles[0])  # Set default value

    window_label = ttk.Label(root, text="Select a Window:")
    window_label.pack()

    window_dropdown = ttk.Combobox(root, textvariable=window_selection, values=window_titles)
    window_dropdown.pack()

    def select():
        nonlocal root
        root.destroy()  # Close the window

    select_button = ttk.Button(root, text="Select", command=select)
    select_button.pack()

    root.mainloop()

    selected_window = window_selection.get()

    if selected_window:
        return selected_window
    else:
        print("No window selected.")
        return None

# Main function
def main():
    # Prompt user to select a file
    file_path = select_file()
    if not file_path:
        return

    # Read text from Word document
    text_to_type = read_word_document(file_path)

    # Prompt user to select a window
    window_title = select_window()
    if not window_title:
        return

    # Switch to the selected window
    win_title = pyautogui.getWindowsWithTitle(window_title)[0].maximize()
    
    # Wait for a few seconds to ensure the window is focused
    pyautogui.sleep(5)

    # Type the text onto the window
    type_text(text_to_type)

if __name__ == "__main__":
    main()
