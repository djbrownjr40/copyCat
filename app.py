import pyperclip
import time
import pickle
import google_auth_oauthlib.flow
import googleapiclient.discovery
import googleapiclient.errors
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from PIL import ImageGrab, ImageTk
from ttkthemes import ThemedTk
import os
import threading
import mimetypes

clipboard_history = []
current_image = None

SCOPES = ['https://www.googleapis.com/auth/drive.file']

CLIENT_SECRET_FILE = 'credentials.json'

FOLDER_ID = '1iEkF5uH9-vi92UvfEZ-HiiCc9yG1hB97'

def get_mime_type(file_path):
    mime_type, _ = mimetypes.guess_type(file_path)
    return mime_type or 'application.octet-stream'

def authenticate_google_drive():
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)  
    return creds

def upload_to_drive(file_path, file_name):
    try:
        creds = authenticate_google_drive()
        service = googleapiclient.discovery.build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': file_name,  # file_name is used here
            'parents': [FOLDER_ID]
        }

        mime_type = 'application/octet-stream'  # Default MIME type, change based on file type
        media = MediaFileUpload(file_path, mimetype=get_mime_type(file_path))

        # Upload the file to Google Drive
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print('File uploaded with ID: %s' % file.get('id'))
        messagebox.showinfo("Success", f"File uploaded to Google Drive with ID: {file.get('id')}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to upload file: {e}")

# List files in Google Drive
def list_files_from_drive():
    try:
        creds = authenticate_google_drive()
        service = googleapiclient.discovery.build('drive', 'v3', credentials=creds)

        # List the first 10 files in Google Drive
        results = service.files().list(pageSize=10, fields="nextPageToken, files(id, name)").execute()
        items = results.get('files', [])

        if not items:
            messagebox.showinfo("No files", "No files found in Google Drive.")
        else:
            file_names = "\n".join([f"{item['name']} (ID: {item['id']})" for item in items])
            messagebox.showinfo("Files in Google Drive", file_names)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to list files: {e}")

# Tkinter GUI integration (you can add buttons to upload or list files)
def gui_upload_file():
    file_path = tk.filedialog.askopenfilename(title="Select a File")
    if file_path:
        file_name = os.path.basename(file_path)
        upload_to_drive(file_path, file_name)

def gui_list_files():
    list_files_from_drive()


# Toggle the window between always on top and normal behavior
def toggle_sticky():
    # Check if the window is "always on top" and toggle the behavior
    if root.attributes("-topmost"):
        root.attributes("-topmost", False)
    else:
        root.attributes("-topmost", True)


def save_as():
    current_tab = tab_control.index(tab_control.select())  # Get the index of the current tab
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
    if file_path:
        try:
            with open(file_path, "w") as file:
                file.write(tabs_text_widgets[current_tab].get("1.0", tk.END))
            messagebox.showinfo("Saved", f"ファイルは保存されました。")
        except Exception as e:
            messagebox.showerror("Error", f"ファイルの保存に失敗しました: {e}")


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
    if file_path:
        try:
            # Create a new tab for the file before the '+' tab
            new_tab = ttk.Frame(tab_control)
            text_widget = tk.Text(new_tab, wrap="word", padx=10, pady=10, width=125, height=20)
            text_widget.pack(side="left", expand=True, fill="both", padx=10, pady=10)


            # Create the vertical scrollbar
            scrollbar = tk.Scrollbar(new_tab, orient="vertical", command=text_widget.yview)
            scrollbar.pack(side="right", fill="y")


            # Link the scrollbar to the text widget
            text_widget.config(yscrollcommand=scrollbar.set)


            # Read and insert the file content into the new text widget
            with open(file_path, "r") as file:
                text_widget.delete("1.0", tk.END)
                text_widget.insert(tk.END, file.read())

            # Insert the new tab just before the last tab (the "+" tab)
            index = len(tab_control.tabs()) - 1  # Get the index of the last tab (which is "+")
            tab_control.insert(index, new_tab, text=file_path)  # Insert new tab before "+" tab
            tabs_text_widgets.append(text_widget)  # Keep track of text widgets in each tab
            tabs_file_paths.append(file_path)  # Keep track of the file paths

        except Exception as e:
            messagebox.showerror("Error", f"ファイルを開くのに失敗しました:: {e}")


# Monitor the clipboard every second
def monitor_clipboard():
    last_clipboard = ""
    while True:
        current_clipboard = pyperclip.paste()
        if current_clipboard != last_clipboard:
            clipboard_history.clear()
            clipboard_history.append(current_clipboard)
            last_clipboard = current_clipboard
            update_text_widget()
        time.sleep(1)


def update_text_widget():
    active_tab_index = tab_control.index(tab_control.select())
    active_text_widget = tabs_text_widgets[active_tab_index]

    if clipboard_history:
        active_text_widget.insert(tk.END, clipboard_history[0] + "\n")


def copy_to_clipboard():
    try:
        # Get the selected text
        selected_text = initial_text_widget.get("sel.first", "sel.last")
        if selected_text:
            pyperclip.copy(selected_text)  # Copy the selected text back to clipboard
            messagebox.showinfo("コピーされました。", f"コピーされました: {selected_text}")
        else:
            messagebox.showwarning("選択がありません。", "コピーしたいアイテムを選択してください。")
    except tk.TclError:
        # If no text is selected, a TclError will be raised
        messagebox.showwarning("選択がありません。", "コピーしたいアイテムを選択してください。")


# Handling tab change (used to create new tabs when + tab is selected)
def handleTabChange(event):
    if tab_control.select() == tab_control.tabs()[-1]:  # Check if the last tab (the '+') is selected
        index = len(tab_control.tabs()) - 1  # Get the index of the last tab
        new_tab = ttk.Frame(tab_control)  # Create a new tab frame
       
        text_widget = tk.Text(new_tab, wrap="word", padx=10, pady=10, width=125, height=20)  # Create a text widget
        text_widget.pack(side="left", expand=True, fill="both", padx=10, pady=10)  # Pack the text widget into the tab
       
        # Create and configure the scrollbar
        scrollbar = tk.Scrollbar(new_tab, orient="vertical", command=text_widget.yview)
        scrollbar.pack(side="right", fill="y")
        text_widget.config(yscrollcommand=scrollbar.set)  # Link the scrollbar to the text widget


        tab_control.insert(index, new_tab, text="新しいタブ")  # Insert the new tab before the last tab
        tabs_text_widgets.append(text_widget)  # Keep track of text widgets in each tab
        tab_control.select(index)  # Select the newly created tab


# Main app window setup
root = ThemedTk()
root.get_themes()
root.set_theme("ubuntu")
root.title("コピーキャット")

icon_image = tk.PhotoImage(file="copycat.png")
root.iconphoto(True, icon_image)

# Creating the notebook widget
tab_control = ttk.Notebook(root)
tab_control.pack(expand=True, fill="both")

tabs_text_widgets = []
tabs_file_paths = []

# Create the initial "new" tab with a text widget
initial_tab = ttk.Frame(tab_control)
tab_control.add(initial_tab, text="新しいタブ")

# Create the text widget inside the initial tab
initial_text_widget = tk.Text(initial_tab, wrap="word", padx=10, pady=10, width=125, height=20)
initial_text_widget.pack(side="left", expand=True, fill="both", padx=10, pady=10)

# Create the vertical scrollbar for the initial tab
scrollbar = tk.Scrollbar(initial_tab, orient="vertical", command=initial_text_widget.yview)
scrollbar.pack(side="right", fill="y")

# Link the scrollbar to the text widget
initial_text_widget.config(yscrollcommand=scrollbar.set)

tabs_text_widgets.append(initial_text_widget)
tabs_file_paths.append(None)

# Add the "+" tab which will trigger the creation of new tabs when clicked
plus_tab = ttk.Frame(tab_control)
tab_control.add(plus_tab, text="+")  # The last tab is a '+' tab

# Bind the event that will create new tabs
tab_control.bind("<<NotebookTabChanged>>", handleTabChange)

# Menu Bar Setup
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="ファイル", menu=file_menu)
file_menu.add_command(label="タブで開く", command=open_file)
file_menu.add_command(label="名前を付けて保存", command=save_as)

window_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="ウインドウ", menu=window_menu)

button_frame = tk.Frame(root)
button_frame.pack(pady=10, fill="x", expand=True)

# Assign the sticky_index here
sticky_index = window_menu.add_checkbutton(label="ステッキー", command=toggle_sticky)

copy_button = tk.Button(button_frame, text="コピー", command=copy_to_clipboard)
copy_button.pack(side="left", padx=5)

upload_button = tk.Button(button_frame, text="Google ドライブにファイルをアップロード", command=gui_upload_file)
upload_button.pack(side="left", padx=5)

list_button = tk.Button(button_frame, text="Google ドライブからファイルを一覧表示する", command=gui_list_files)
list_button.pack(side="left", padx=5)

# Start clipboard monitoring in a separate thread
monitor_thread = threading.Thread(target=monitor_clipboard, daemon=True)
monitor_thread.start()

root.mainloop()













