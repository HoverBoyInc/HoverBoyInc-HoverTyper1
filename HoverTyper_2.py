import tkinter as tk
import random
from tkinter import filedialog
from tkinter import ttk
import time
import pyautogui
import docx
from docx import Document
import PyPDF2
import asyncio
import os

root = tk.Tk()
root.title('Hover Typer 2')

# Set window size and position
#root.geometry("564x500+100+100")

# Make the window borderless
root.overrideredirect(True)

# Load a mask image with rounded corners
mask_image = tk.PhotoImage(file="htp2logoD.png")
photo_img = tk.PhotoImage(file="icons8-button-102.png")
close_img = tk.PhotoImage(file="icons8-close-windowa-25.png")
mini_img = tk.PhotoImage(file="icons8-minimizeb-window-25.png")
mouse_emo = tk.PhotoImage(file="icons8-mouseemoji-40.png")
ikey = tk.PhotoImage(file="icons8-keyboard-25.png")
iplay = tk.PhotoImage(file="icons8-play-button-25.png")
ipause = tk.PhotoImage(file="icons8-pause-25.png")
istop = tk.PhotoImage(file="icons8-stop-button-25.png")
ikei = tk.PhotoImage(file="icons8-keyboardt-25.png")
# Set the window shape using the mask image
root.attributes("-transparentcolor", "black")
root.configure(bg='black')

def animate():
    global current_char_index
    if current_char_index < len(label_text):
        update_keyboard_colors(label_text[current_char_index])
        label.config(text=label_text[:current_char_index+1])
        label.config(fg=random.choice(colors))
        current_char_index += 1
        root.after(150, animate)

def update_keyboard_colors(current_letter):
    for widget in keyboard_frame.winfo_children():
        if isinstance(widget, tk.Label) and widget.cget("text") == current_letter:
            widget.config(fg=random.choice(colors))
            root.after(140, lambda: widget.config(fg="black"))

selected_file = tk.StringVar()

def browse_file():
    file_path = filedialog.askopenfilename(title="Select File", filetypes=[
        ("Word files", "*.docx"),
        ("PDF files", "*.pdf"),
        ("Text files", "*.txt") # Include an option for all file types
    ])
    if file_path:
        file_name = os.path.basename(file_path)  # Get the base name (filename)
        file_extension = os.path.splitext(file_name)[1]  # Get the extension
        selected_file.set(file_name)
        label = tk.Label(mask_frame, textvariable=selected_file, fg='#0C3919', bg='#90A74F',font=('Arial', 9))
        label.place(relx=0.5, rely=0.96, anchor='s')
        button_frame.place_forget()
        keyboard_frame.place_forget()
        emo_frame.place_forget()
        window_titles = [window.title for window in pyautogui.getAllWindows()]

        window_selection = tk.StringVar()
        #window_selection.set(window_titles[0])  # Set default value

        window_label = tk.Label(mask_frame, text="Select a Window:", font=('Helvetica', 10,'bold'), fg='#0C3919', bg='#90A74F')
        window_label.place(relx=0.5, rely=0.35,anchor='center')

        window_dropdown = ttk.Combobox(mask_frame, textvariable=window_selection, values=window_titles, font=('helvetica', 12), background='orange', justify='center')
        window_dropdown.place(relx=0.5,rely=0.45,anchor='center')
        window_dropdown.configure(state='readonly')
        window_dropdown['background'] = 'orange'

        async def read_file(file_path):
            try:
                if file_path.endswith('.docx'):
                    doc = Document(file_path)
                    text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
                    return text
                elif file_path.endswith('.pdf'):
                    encodings_to_try = ['utf-8', 'latin-1', 'windows-1252']
                    text = ''
                    with open(file_path, 'rb') as file:
                        pdf_reader = PyPDF2.PdfReader(file)
                        for page in pdf_reader.pages:
                            extracted_text = page.extract_text()
                            for encoding in encodings_to_try:
                                try:
                                    decoded_text = extracted_text.encode('latin-1', 'replace').decode(encoding)
                                    text += decoded_text
                                    break
                                except UnicodeDecodeError:
                                    continue
                    return text
                else:
                    with open(file_path, 'r', encoding='utf-8') as file:
                        text = file.read()
                    return text
            except Exception as e:
                print(f"Error reading file: {e}")
                err_frame= tk.Frame(mask_frame)
                err_frame.place(relx=0.4, rely=0.8)
                err_label = tk.Label(err_frame, text='Error Reading File', fg='red', font=('helvetica', 9, 'bold'))
                err_label.pack()
                def fade():
                    err_frame.place_forget()
                err_label.after(4000, fade)
                return None                
        async def reading():
            global read
            read = await read_file(file_path)
            return read

        asyncio.run(reading())
        
        def select():
            selected_window = window_selection.get()
            if selected_window:
                print(selected_window) 
                # Switch to the selected window
                win_title = pyautogui.getWindowsWithTitle(selected_window)[0].maximize()
                canvas.pack_forget()
                # Get the screen resolution
                screen_width = root.winfo_screenwidth()
                screen_height = root.winfo_screenheight()

                # Calculate the position for the root window
                root_width = root.winfo_width()
                root_height = root.winfo_height()
                x_position = screen_width - root_width  # Place at the very far right
                y_position = (screen_height - root_height) // 2  # Center vertically
                # Place the root window
                root.geometry(f"{screen_width}x{screen_height}")
                root.columnconfigure(0, weight=1)
                root.rowconfigure(0, weight=1)
                #Create Home Frame
                home_frame = tk.Frame(root, bg='black')
                home_frame.grid(row=0,column=0, sticky='nsew')
                #Create Icon Frame
                ico_frame = tk.Frame(root)
                ico_frame.grid(row=0, column=0, sticky='e')
                ico_key = tk.Button(ico_frame, image=ikey)
                ico_key.pack()
                ico_play = tk.Button(ico_frame, image=iplay)
                ico_play.pack()
                ico_pause = tk.Button(ico_frame, image=ipause)
                ico_pause.pack()
                ico_stop = tk.Button(ico_frame, image=istop)
                ico_stop.pack()
                
                def key_color():
                    def key_back():
                        ico_key.configure(image=ikey)
                        ico_key.after(1000,key_color)
                    ico_key.configure(image=ikei)
                    ico_key.after(1000,key_back)
                

                root.lift()
                key_color()                
                # Wait for a few seconds to ensure the window is focused
                pyautogui.sleep(5)                
                def type_text(read):
                    try:
                        typing_in_progress = True  # Set flag to indicate typing is in progress
                        start_time = time.time() #Record the start time
                        chunk_size = 100  # Adjust the chunk size as needed
                        chunks = [read[i:i+chunk_size] for i in range(0, len(read), chunk_size)]
                        for chunk in chunks:
                            pyautogui.typewrite(chunk)
                            time.sleep(0.1)  # Adjust the sleep time as needed
                        def key_stop():
                            global typing_in_progress
                            typing_in_progress = False
                            ico_key.configure(image=ikei)
                        key_stop()
                    except Exception as e:
                        print(f"An error occurred while typing: {e}")
                    end_time = time.time() #Record the end time
                    duration = end_time - start_time #Calculate the duration
                    def kill():
                        root.quit()
                    root.after(int(duration), kill)
                ico_key.after(1000, lambda:type_text(read))
            else:
                print("No window selected.")
                err_frame= tk.Frame(mask_frame)
                err_frame.place(relx=0.37, rely=0.8)
                err_label = tk.Label(err_frame, text='No Window Selected', fg='red', font=('helvetica', 9, 'bold'))
                err_label.pack()
                def fade():
                    err_frame.place_forget()
                err_label.after(4000, fade)
                return None
        select_frame = tk.Frame(mask_frame)
        select_frame.place(relx=0.5, rely=0.6,anchor='center')
        select_button = tk.Button(select_frame, text="Start Typing", command=select, image=photo_img, font= ("Helvetica", 7, "bold italic"), compound=tk.CENTER, bd=0, fg='#0C3919', bg='#0C3919')
        select_button.pack()

# Create a canvas to draw the rounded window shape
canvas = tk.Canvas(root, bg='black', highlightthickness=0)
canvas.create_image(0, 0, anchor="nw", image=mask_image)
canvas.pack()

#Create frame for mask_label
mask_frame = tk.Frame(canvas, bg='white')
mask_frame.pack()

# Create a label with the mask image
mask_label = tk.Label(mask_frame, image=mask_image, bg='black')
mask_label.pack()

# Add widgets or other elements to the window
'''label = tk.Label(root, text="Hello, Rounded Window!", font=("Helvetica", 16))
label.pack(pady=20)'''
#Create button frame
mini_frame = tk.Frame(mask_frame, bg='#8746DA')
mini_frame.place(relx=0.93, rely=0.05, anchor='ne')  # Place button_frame in the center of mask_frame

def on_restore(event):
    root.overrideredirect(True)

def minimize_window():
    root.overrideredirect(False)
    root.iconify()
    root.bind("<Map>", on_restore)

#Create minimize button
mini = tk.Button(mini_frame, image=mini_img, compound=tk.CENTER, bd=0, bg="#8746DA", command=minimize_window)
mini.pack()

close_frame = tk.Frame(mask_frame, bg='#8746DA')
close_frame.place(relx=0.99, rely=0.05, anchor='ne')  # Place button_frame in the center of mask_frame
def close():
    root.quit()
#Create close button
close = tk.Button(close_frame, image=close_img, compound=tk.CENTER, bd=0, bg="#8746DA", command=close)
close.pack()

colors = ["#161477", "#DE1FCB", "#21A43E", "orange", '#CEAA2C', '#6D1016']

label_text = "Hover Typer 2"
current_char_index = 0

label = tk.Label(mask_frame, text="", bg="#8746DA", font=("Garamond", 16, 'bold'))
label.place(relx=0.35, rely=0.15)

keyboard_frame = tk.Frame(mask_frame,bg='#0C3919')
keyboard_frame.place(relx=0.241, rely=0.33)

# QWERTY keyboard layout
keyboard_layout = [
    "QWERTYUIOP",
    "ASDFGHJKL",
    "ZXCVBNM"
]

# Keyboard letters
keyboard_letters = [letter for row in keyboard_layout for letter in row]

for row_idx, row in enumerate(keyboard_layout):
    for col_idx, letter in enumerate(row):
        letter_label = tk.Label(keyboard_frame, text=letter,bg='#90A74F', padx=5, pady=10, font=("Garamond", 9), borderwidth=1, relief="raised")
        letter_label.grid(row=row_idx, column=col_idx, padx=1)

#Create emoji frame
emo_frame = tk.Frame(mask_frame, bg='#8746DA')
emo_frame.place(relx=0.74, rely=0.715, anchor='se')  # Place button_frame in the center of mask_frame
#Place image emoji
mouse = tk.Button(emo_frame, image=mouse_emo, compound=tk.CENTER, bd=0, bg="#0C3919")
mouse.pack()

animate()  # Start animation

#Create button frame
button_frame = tk.Frame(mask_frame, bg='#8746DA', highlightbackground='green')
button_frame.place(relx=0.5, rely=0.98, anchor='s')  # Place button_frame in the center of mask_frame

# Create a custom font
custom_font = ("Helvetica", 7, "bold italic")
# Create a button with both text and a transparent image
button = tk.Button(button_frame, text="Read File", image=photo_img, font=custom_font, compound=tk.CENTER, bd=0, fg='#0C3919', bg="#8746DA", command=browse_file)
button.pack()

# Drag window function
def move_window(event):
    root.geometry(f"+{event.x_root}+{event.y_root}")

mask_label.bind("<B1-Motion>", move_window)

root.mainloop()
