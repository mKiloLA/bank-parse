import tkinter as tk
from tkinter import ttk
from data_types.DataTypes import Category

class CategoryPopup(tk.Toplevel):
    def __init__(self, master=None, desc_default: str = "", amount: float = 0.0, date: str = ""):
        super().__init__(master)
        self.title("Add New Item")
        popup_w, popup_h = 400, 250
        if master is not None:
            master.update_idletasks()
            main_x = master.winfo_x()
            main_y = master.winfo_y()
            main_w = master.winfo_width()
            main_h = master.winfo_height()
            x = main_x + (main_w // 2) - (popup_w // 2)
            y = main_y + (main_h // 2) - (popup_h // 2)
            self.geometry(f"{popup_w}x{popup_h}+{x}+{y}")
        else:
            self.geometry(f"{popup_w}x{popup_h}")
        self.result = None
        self.desc_var = tk.StringVar(value=desc_default)
        self.title_var = tk.StringVar()
        self.cat_var = tk.StringVar()
        info_text = f"Add a category for the purchase at {desc_default} for {amount} on {date}."
        tk.Label(self, text=info_text, wraplength=380, justify="left").pack(pady=10)
        tk.Label(self, text="Description:").pack()
        tk.Entry(self, textvariable=self.desc_var, width=40).pack()
        tk.Label(self, text="Title:").pack()
        tk.Entry(self, textvariable=self.title_var, width=40).pack()
        tk.Label(self, text="Category:").pack()
        cat_combo = ttk.Combobox(self, textvariable=self.cat_var, values=[c.value for c in Category], state="readonly")
        cat_combo.pack()
        tk.Button(self, text="Submit", command=self.submit).pack(pady=10)
        self.protocol("WM_DELETE_WINDOW", self.submit)

    def submit(self):
        self.result = (
            self.desc_var.get().strip().lower(),
            self.title_var.get().strip(),
            self.cat_var.get().strip()
        )
        self.destroy()
