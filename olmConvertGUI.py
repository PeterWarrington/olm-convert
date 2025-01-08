import olmConvert
import tkinter as tk
from tkinter import ttk, font, filedialog, messagebox
import webbrowser
import sys
from threading import Thread
import re 
import os
import platform
import subprocess
import time

# Window setup
root = tk.Tk()
try:
    os.chdir(sys._MEIPASS)
except AttributeError:
    pass
root.wm_iconphoto(False, tk.PhotoImage(file = 'olmConvert.png'))
root.title("OLM Convert")
root.resizable(False, False)
frame = ttk.Frame(root, padding=10)
frame.grid()

# Information UI setup

ttk.Label(frame, text="OLM Convert", font=font.Font(weight="bold")).grid(column=0, row=0, columnspan=3)
ttk.Label(frame, text="By Peter Warrington").grid(column=0, row=1, columnspan=3)

web_link = "https://www.lilpete.me/olm-convert"
web_link_label = ttk.Label(frame, text=web_link, foreground="#0071e3", cursor="hand2")
web_link_label.grid(column=0, row=2, columnspan=3)
web_link_label.bind("<Button-1>", lambda e: webbrowser.open(web_link))

# OLM File selector setup

ttk.Label(frame, text="OLM File:").grid(column=0, row=3, sticky="E")
olm_path_entry = ttk.Entry(frame, width=30)
olm_path_entry.grid(column=1, row=3)

def olm_browse():
    filepath = filedialog.askopenfilename(
        title='Select Outlook For Mac Archive',
        filetypes=[("Outlook For Mac Archive", "*.olm")]
    )
    olm_path_entry.delete(0, tk.END)
    olm_path_entry.insert(0, filepath)

olm_browse_btn = ttk.Button(frame, text="Browse...", command=olm_browse)
olm_browse_btn.grid(column=2, row=3)

# Output directory selector setup

ttk.Label(frame, text="Output directory:").grid(column=0, row=4, sticky="E")
output_path_entry = ttk.Entry(frame, width=30)
output_path_entry.grid(column=1, row=4)

def output_browse():
    folderpath = filedialog.askdirectory(title="Select EML output directory")
    output_path_entry.delete(0, tk.END)
    output_path_entry.insert(0, folderpath)

folder_browse_btn = ttk.Button(frame, text="Browse...", command=output_browse)
folder_browse_btn.grid(column=2, row=4)

# Include attachments checkbox

attachments_var = tk.BooleanVar(value = True)
attachments_checkbox = ttk.Checkbutton(frame, text="Include attachments", variable=attachments_var)
attachments_checkbox.grid(column=1, row=5, sticky="E", padx=5)

# Convert button

def run_convert():
    def convert_wrapper():
        try:
            olmConvert.convertOLM(olm_path_entry.get(), output_path_entry.get(), not attachments_var.get(), True)
            messagebox.showinfo("Conversion complete", "The conversion has finished successfully.")
            try:
                if platform.system() == "Darwin":
                    subprocess.run(["open", output_path_entry.get()])
                elif platform.system() == "Windows":
                    subprocess.run(["explorer", output_path_entry.get()])
            except:
                pass
        except Exception as e:
            messagebox.showerror("Conversion unsuccessful", f"Conversion unsuccessful: {e}")
            print(e.with_traceback())
        
        convert_btn.config(state=tk.NORMAL)

    if (olm_path_entry.get().strip() == ""):
        messagebox.showerror("No OLM file selected", "No OLM file selected.")
    elif (not os.path.isfile(olm_path_entry.get())):
        messagebox.showerror("OLM File not found", f"'{olm_path_entry.get()}' does not exist.")
    elif (output_path_entry.get().strip() == ""):
        messagebox.showerror("No output directory selected", "No output directory selected.")
    elif (not os.path.isdir(output_path_entry.get())):
        messagebox.showerror("OLM directory not found", f"'{output_path_entry.get()}' does not exist.")
    else:
        convert_btn.config(state=tk.DISABLED)
        convert_thread = Thread(target=convert_wrapper)
        convert_thread.start()

convert_btn = ttk.Button(frame, text="Convert!", command=run_convert)
convert_btn.grid(column=2, row=5)

# Progress bar

progress = tk.DoubleVar()
progressbar = ttk.Progressbar(frame, variable=progress, length=500)
progressbar.grid(column=0, row=6, columnspan=3)

# Output entry

output_shown_flag = tk.BooleanVar(value=False)

def toggle_output():
    if (not output_shown_flag.get()):
        output_shown_flag.set(True)
        output_label.config(text="▼ Output (advanced):")
        output_text.grid(column=0, row=8, columnspan=3)
    else:
        output_shown_flag.set(False)
        output_label.config(text="▶ Show output (advanced):")
        output_text.grid_forget()

output_label = ttk.Label(frame, text="▶ Show output (advanced):")
output_label.grid(column=0, row=7, sticky="W", columnspan=3)
output_label.bind("<Button-1>", lambda e: toggle_output())
output_text = tk.Text(frame, width=70, height=20)

# Output configure

class FakeStdout():
    def __init__(self, old_stdout):
        self.old_stdout = old_stdout

    def write(self, string):
        progress_match = re.search("^\[(\d+)\/(\d+)]\:", string)
        if (progress_match):
            index = int(progress_match.group(1))
            out_of = int(progress_match.group(2))
            progress.set((index / out_of) * 100)
        else:
            self.old_stdout.write(string)
            output_text.insert(tk.END, string)
            output_text.see(tk.END)

    def flush(self):
        pass

class FakeStderr():
    def __init__(self, old_stderr):
        self.old_stderr = old_stderr

    def write(self, string):
        self.old_stderr.write(string)
        output_text.insert(tk.END, string)
        output_text.see(tk.END)

    def flush(self):
        pass

sys.stdout = FakeStdout(sys.stdout)
sys.stderr = FakeStderr(sys.stderr)

# Mainloop

frame.focus_force()
root.mainloop()