# Shortcut-Creator.py
import os
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, Text, END
from win32com.client import Dispatch

def create_shortcut(target_path, shortcut_path, working_dir, arguments, icon_path):
    """
    Creates a .lnk (shortcut) file.
    """
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.WorkingDirectory = working_dir or os.path.dirname(target_path)
    shortcut.Arguments = arguments
    if icon_path:
        shortcut.IconLocation = icon_path
    shortcut.save()

def browse_file(entry_field, file_types=(("All Files", "*.*"),)):
    """
    Opens a file dialog and sets the selected file path to the given entry field.
    """
    file_path = filedialog.askopenfilename(filetypes=file_types)
    if file_path:
        entry_field.delete(0, "end")
        entry_field.insert(0, file_path)

def browse_directory(entry_field):
    """
    Opens a directory dialog and sets the selected directory to the given entry field.
    """
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_field.delete(0, "end")
        entry_field.insert(0, folder_path)

def preview_shortcut():
    """
    Previews the shortcut details in the preview box.
    """
    target = target_entry.get()
    shortcut = shortcut_entry.get()
    working_dir = working_dir_entry.get()
    arguments = arguments_entry.get()
    icon = icon_entry.get()

    preview_box.delete(1.0, END)
    if not target:
        preview_box.insert(END, "Error: Target file/command is required!\n")
        return
    if not shortcut:
        preview_box.insert(END, "Error: Shortcut location is required!\n")
        return

    preview_box.insert(END, f"Shortcut Details:\n")
    preview_box.insert(END, f"Target Path: {target}\n")
    preview_box.insert(END, f"Shortcut Path: {shortcut}\n")
    if working_dir:
        preview_box.insert(END, f"Working Directory: {working_dir}\n")
    if arguments:
        preview_box.insert(END, f"Arguments: {arguments}\n")
    if icon:
        preview_box.insert(END, f"Icon Path: {icon}\n")

def generate_shortcut():
    """
    Generates a shortcut file based on user input.
    """
    target = target_entry.get()
    shortcut = shortcut_entry.get()
    working_dir = working_dir_entry.get()
    arguments = arguments_entry.get()
    icon = icon_entry.get()

    if not target or not shortcut:
        messagebox.showerror("Error", "Target file and shortcut location are required!")
        return

    try:
        create_shortcut(target, shortcut, working_dir, arguments, icon)
        messagebox.showinfo("Success", f"Shortcut created successfully at: {shortcut}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create shortcut:\n{str(e)}")

# GUI Setup
app = Tk()
app.title("Shortcut Creator - by Coder333-A")
app.geometry("600x500")

# Labels and Entry Fields
Label(app, text="Target File/Command:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
target_entry = Entry(app, width=60)
target_entry.grid(row=0, column=1, padx=10, pady=5)
Button(app, text="Browse", command=lambda: browse_file(target_entry)).grid(row=0, column=2, padx=10)

Label(app, text="Shortcut Location (.lnk):").grid(row=1, column=0, sticky="w", padx=10, pady=5)
shortcut_entry = Entry(app, width=60)
shortcut_entry.grid(row=1, column=1, padx=10, pady=5)
Button(app, text="Browse", command=lambda: browse_file(shortcut_entry, file_types=(("Shortcut Files", "*.lnk"),))).grid(row=1, column=2, padx=10)

Label(app, text="Working Directory (Optional):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
working_dir_entry = Entry(app, width=60)
working_dir_entry.grid(row=2, column=1, padx=10, pady=5)
Button(app, text="Browse", command=lambda: browse_directory(working_dir_entry)).grid(row=2, column=2, padx=10)

Label(app, text="Arguments (Optional):").grid(row=3, column=0, sticky="w", padx=10, pady=5)
arguments_entry = Entry(app, width=60)
arguments_entry.grid(row=3, column=1, padx=10, pady=5)

Label(app, text="Icon File (Optional):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
icon_entry = Entry(app, width=60)
icon_entry.grid(row=4, column=1, padx=10, pady=5)
Button(app, text="Browse", command=lambda: browse_file(icon_entry)).grid(row=4, column=2, padx=10)

# Preview Box
Label(app, text="Shortcut Preview:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
preview_box = Text(app, height=10, width=80)
preview_box.grid(row=6, column=0, columnspan=3, padx=10, pady=5)

# Buttons
Button(app, text="Preview Shortcut", command=preview_shortcut, width=20).grid(row=7, column=0, padx=10, pady=20)
Button(app, text="Create Shortcut", command=generate_shortcut, width=20).grid(row=7, column=1, padx=10, pady=20)

app.mainloop()
