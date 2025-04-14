import pandas as pd
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk
from sqlalchemy import create_engine

def browse_file():
    path = filedialog.askopenfilename(filetypes=[("Exel files", "*.xlsx *.xls")])
    if path:
        file_path.set(path)

        priview_exel_sheets(path)

def priview_exel_sheets(path):
    try:
        xls = pd.read_excel(path, sheet_name=None)
        sheets_names = list(xls.keys())

        sheets_listbox.delete(0, tk.END)

        for i, name in enumerate(sheets_names, 1):
            sheets_listbox.insert(tk.END, f"{i}. {name}")

    except Exception as ex:
        messagebox.showerror("Error reading the file", str(ex))

def migrate_exel_to_psql_conf():
    try:
        db = connection_string.get()
        if not db or not file_path.get():
            messagebox.showerror("Error", "Invalid connection string or file path")
            return
        
        path = file_path.get()

        xls = pd.read_excel(path, sheet_name=None)
        sheet_names = list(xls.keys())
        confirmed = messagebox.askyesno("Are you Agree?", f"{len(sheet_names)} tables will be loaded. Continue?")
        if not confirmed:
            return
        
        engine = create_engine(db)
        log_scroll.delete(1.0, tk.END)

        for sheet in sheet_names:
            df = xls[sheet]
            df.to_sql(sheet, engine, if_exists="append", index=False)
            log_scroll.insert(tk.END, f"Loaded {sheet} with {len(df)} entries")
        log_scroll.insert(tk.END, "Successful download.")
    except Exception as ex:
        messagebox.showerror("Migration Error", str(ex))
        


##GUI
root = tk.Tk()
root.title("Import exel to psql")
root.geometry("600x500")

file_path = tk.StringVar()
connection_string = tk.StringVar()

tk.Label(root, text="Select exel file").pack()
tk.Entry(root, textvariable=file_path, width=60).pack()
tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

tk.Label(root, text="Enter db-connection string")
tk.Entry(root, textvariable=connection_string, width=60).pack()
tk.Label(root, text="Example: postgersql+psycopg2://user:pass@localhost/dbname", fg="gray").pack(pady=5)

tk.Label(root, text="Table import order(top to bottom)", font=('Arial', 10, 'bold')).pack()
sheets_listbox = tk.Listbox(root, height=8, width=50)
sheets_listbox.pack()

tk.Button(root, text="Import data", bg="#2acaea", fg="#F5F5F5", command=migrate_exel_to_psql_conf).pack(pady=10)

tk.Label(root, text="Logs:").pack()
log_scroll = scrolledtext.ScrolledText(root, height=10)
log_scroll.pack(fill="both", expand=True)

root.mainloop()