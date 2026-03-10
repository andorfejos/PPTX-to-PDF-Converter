import os
import tkinter as tk
from tkinter import filedialog, messagebox
import comtypes.client

def convert_pptx_to_pdf(input_paths, output_folder, custom_name):
    try:
        # Initialize PowerPoint once for the whole batch
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        
        for path in input_paths:
            abs_input = os.path.abspath(path)
            base_name = os.path.basename(path)
            name_without_ext = os.path.splitext(base_name)[0]

            # Logic for naming the output file
            if len(input_paths) == 1 and custom_name.strip():
                final_name = f"{custom_name.strip()}.pdf"
            else:
                final_name = f"{name_without_ext}.pdf"

            abs_output = os.path.abspath(os.path.join(output_folder, final_name))

            # Open and Save
            presentation = powerpoint.Presentations.Open(abs_input, WithWindow=False)
            presentation.SaveAs(abs_output, 32) # 32 = PDF format
            presentation.Close()

        powerpoint.Quit()
        messagebox.showinfo("Success", f"Converted {len(input_paths)} file(s) successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# --- GUI Setup ---
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("PowerPoint files", "*.pptx")])
    if files:
        file_list_label.config(text=f"{len(files)} files selected")
        root.selected_files = files

def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_label.config(text=folder)
        root.output_folder = folder

def start_conversion():
    if not hasattr(root, 'selected_files') or not root.selected_files:
        return messagebox.showwarning("Input Missing", "Please select PPTX files first.")
    if not hasattr(root, 'output_folder') or not root.output_folder:
        return messagebox.showwarning("Output Missing", "Please select a destination folder.")
    
    convert_pptx_to_pdf(root.selected_files, root.output_folder, name_entry.get())

# --- Window Layout ---
root = tk.Tk()
root.title("PPTX to PDF Converter")
root.geometry("400x350")
root.padx = 20

tk.Label(root, text="Step 1: Select PowerPoint Files", font=('Arial', 10, 'bold')).pack(pady=(10,0))
tk.Button(root, text="Browse Files", command=select_files).pack(pady=5)
file_list_label = tk.Label(root, text="No files selected", fg="gray")
file_list_label.pack()

tk.Label(root, text="Step 2: Select Output Folder", font=('Arial', 10, 'bold')).pack(pady=(15,0))
tk.Button(root, text="Choose Folder", command=select_folder).pack(pady=5)
folder_label = tk.Label(root, text="No folder selected", fg="gray")
folder_label.pack()

tk.Label(root, text="Step 3: Custom PDF Name (Optional)", font=('Arial', 10, 'bold')).pack(pady=(15,0))
tk.Label(root, text="(Only works if 1 file is selected)", font=('Arial', 8)).pack()
name_entry = tk.Entry(root, width=40)
name_entry.pack(pady=5)

tk.Frame(root, height=2, bd=1, relief="sunken").pack(fill="x", pady=10)

tk.Button(root, text="CONVERT TO PDF", bg="#2ecc71", fg="white", font=('Arial', 12, 'bold'), 
          padx=20, command=start_conversion).pack(pady=10)

root.mainloop()