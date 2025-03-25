
import traceback
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from HMI_Translator_V02 import Translate  # This is where you'd import your translation script

def load_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, filepath)

def start_translation():
    api_key = api_key_entry.get()
    file_path = file_path_entry.get()
    num_rows = num_rows_entry.get()  # Get the number of rows from the entry widget
    target_language = target_language_entry.get()
    num_startrow = num_startrow_entry.get()
    try:
        num_rows = int(num_rows)  # Convert the string to an integer
        num_startrow = int(num_startrow )
    except ValueError:
        messagebox.showwarning("Warning", "Please enter a valid integer for the number of rows")
        return
    if not api_key or not file_path:
        messagebox.showwarning("Warning", "Please fill in all fields")
        return
    if not num_startrow:
        num_startrow =2
    
    # Call your translation function here
    try:
        Translate(file_path, api_key,num_rows,target_language,num_startrow)  # Adapt this call to your module's function
        
        messagebox.showinfo("Success", "Translation completed successfully!")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        #traceback.print_exc()


def show_help():
    # Create a new window with the title "Help"
    help_window = tk.Toplevel(app)
    help_window.title("Help")

    # Define the help text
    help_text = "For the API key, you need to create an account on DeepL's website, "\
                "Navigate to profile, and then under API copy and paste the key in "\
                "in the appropriate section."\
                "\n\n"\
                "For The target langue please use the universal Characters i.e:\n"\
                "English: EN   French: FR ...\n\n"\
                "Set starting row to 2 for default\n"\
                "This program is made for Project texts(.xlsx) Generated "\
                "by Tia Portal"\
                
                

    # Create a label in the help_window to display the help text
    tk.Label(help_window, text=help_text, wraplength=400).pack(padx=10, pady=10)

    # Optional: You can also add a button to close the help window
    tk.Button(help_window, text="Close", command=help_window.destroy).pack(pady=(0, 10))
    

app = tk.Tk()
app.title("Excel Translator")

tk.Label(app, text="API Key:").pack()
api_key_entry = tk.Entry(app)
api_key_entry.pack()

tk.Label(app, text="Excel File:").pack()
file_path_entry = tk.Entry(app)
file_path_entry.pack()
tk.Button(app, text="Browse", command=load_file).pack()

tk.Label(app, text="Start Row:").pack()  # Label for the number of rows
num_startrow_entry = tk.Entry(app)  # Entry widget to input the number of rows
num_startrow_entry.pack()

tk.Label(app, text="Number of Rows to translate:").pack()  # Label for the number of rows
num_rows_entry = tk.Entry(app)  # Entry widget to input the number of rows
num_rows_entry.pack()

tk.Label(app, text="Target Language").pack()
target_language_entry = tk.Entry(app)
target_language_entry.pack()


tk.Button(app, text="Translate", command=start_translation).pack()
tk.Button(app, text="Read ME!", command=show_help).pack(pady=10)


# Create a Label widget to display the text
#text_label = tk.Label('Line N°',,'Done in ',f"{app:.3f}",'ms')
text_label = tk.Label(text="KK")
text_label.pack()
app.mainloop()
