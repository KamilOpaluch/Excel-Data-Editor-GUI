import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import pandas as pd

excel_file='Excel_example.xlsx'

class DateEntry(tk.Entry):
    def __init__(self, parent, callback, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.parent = parent
        self.callback = callback
        self.format = list("__-__-____")
        self.var = tk.StringVar(value="".join(self.format))
        self.config(textvariable=self.var)
        self.bind("<KeyPress>", self._on_key_press)
        self.bind("<Return>", self._on_return_press)
        self.config(fg="grey")

    def _on_key_press(self, event):
        if event.char.isdigit():
            pos = self.index(tk.INSERT)
            while pos < len(self.format) and self.format[pos] in ["-", " "]:
                pos += 1
            if pos < len(self.format):
                self.format[pos] = event.char
                self.var.set("".join(self.format))
                self.icursor(pos + 1)
            return "break"
        elif event.keysym == "BackSpace":
            pos = self.index(tk.INSERT) - 1
            while pos >= 0 and pos < len(self.format) and self.format[pos] in ["-", " "]:
                pos -= 1
            if pos >= 0:
                self.format[pos] = "_"
                self.var.set("".join(self.format))
                self.icursor(pos)
            return "break"

    def _on_return_press(self, event):
        self.callback()


class ExcelGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel editor")
        self.master.resizable(False, False)

        self.master.protocol("WM_DELETE_WINDOW", self.close_window)

        self.df = pd.read_excel(excel_file, engine='openpyxl')
        self.df['DATE'] = pd.to_datetime(self.df['DATE']).apply(lambda x: x.strftime('%d-%m-%Y'))

        self.top_frame = ttk.Frame(master)
        self.top_frame.pack(pady=10, fill=tk.X)

        self.label = ttk.Label(self.top_frame, text="Select exception type:")
        self.label.pack(side=tk.LEFT, padx=5)

        self.options = ["Event", "Caa", "Rach"]
        self.dropdown_var = tk.StringVar()
        self.dropdown = ttk.Combobox(self.top_frame, textvariable=self.dropdown_var, values=self.options)
        self.dropdown.bind("<<ComboboxSelected>>", self.filter_table)
        self.dropdown.pack(side=tk.LEFT, padx=5)

        self.tree_frame = ttk.Frame(master)
        self.tree_frame.pack(pady=20, fill=tk.X)

        self.tree = ttk.Treeview(self.tree_frame, columns=("CKK", "NAME", "DATE", "Value"), show='headings')
        self.tree.heading("CKK", text="CKK")
        self.tree.heading("NAME", text="NAME")
        self.tree.heading("DATE", text="DATE")
        self.tree.heading("Value", text="Value")
        self.tree.pack(side=tk.LEFT)

        self.tree.bind("<Button-1>", self.on_cell_click)

        self.scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        self.commit_btn = ttk.Button(master, text="Commit", command=self.commit_changes)
        self.commit_btn.pack(side="left", padx=10)

        self.add_btn = ttk.Button(master, text="Add Line", command=self.add_line)
        self.add_btn.pack(side="left", padx=10)

        self.restart_btn = ttk.Button(master, text="Restart changes", command=self.restart_changes)
        self.restart_btn.pack(side="left", padx=10)

    def close_window(self):
        self.master.destroy()

    def filter_table(self, event):
        for row in self.tree.get_children():
            self.tree.delete(row)

        filtered_df = self.df[self.df["Value"] == self.dropdown_var.get()]

        for index, row in filtered_df.iterrows():
            self.tree.insert("", "end", values=(row["CKK"], row["NAME"], row["DATE"], row["Value"]))

    def add_line(self):
        new_tree_row = self.tree.insert("", 0, values=("", "", "", ""))
        self.tree.selection_set(new_tree_row)
        new_row_index = 0
        self.df.loc[-1] = {"CKK": "", "NAME": "", "DATE": "", "Value": ""}
        self.df.index = self.df.index + 1
        self.df = self.df.sort_index()

    def commit_changes(self):
        self.df.to_excel(excel_file, index=False, engine='openpyxl')
        self.df = pd.read_excel(excel_file, engine='openpyxl')
        self.filter_table(None)

    def restart_changes(self):
        self.df = pd.read_excel(excel_file, engine='openpyxl')
        self.filter_table(None)

    def on_cell_click(self, event):
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)

        if not row_id:
            return

        if hasattr(self, 'edit_widget'):
            self.edit_widget.destroy()

        column_no = int(column_id.replace('#', '')) - 1
        cell_value = self.tree.item(row_id, 'values')[column_no]
        x, y, width, height = self.tree.bbox(row_id, column_id)

        if self.tree.heading(column_id)["text"] == "Value":
            self.edit_widget = ttk.Combobox(self.tree_frame, values=["Event", "Rach", "caa", "dodatk"],
                                            state='readonly')
            self.edit_widget.place(x=x, y=y, width=width, height=height)
            self.edit_widget.set(cell_value)
            self.edit_widget.bind("<<ComboboxSelected>>", lambda e: self.finish_editing(row_id, column_no))
            self.edit_widget.bind('<FocusOut>', lambda e: self.finish_editing(row_id, column_no))
            self.edit_widget.focus_set()
        elif self.tree.heading(column_id)["text"] == "DATE":
            self.edit_widget = DateEntry(self.tree_frame, lambda: self.finish_editing(row_id, column_no))
            self.edit_widget.place(x=x, y=y, width=width, height=height)
            self.edit_widget.bind('<FocusOut>', lambda e: self.finish_editing(row_id, column_no))
            self.edit_widget.delete(0, tk.END)
            self.edit_widget.insert(0, cell_value)
            self.edit_widget.focus_set()
        else:
            self.edit_widget = tk.Entry(self.tree_frame, text=cell_value)
            self.edit_widget.place(x=x, y=y, width=width, height=height)
            self.edit_widget.insert(0, cell_value)
            self.edit_widget.bind('<Return>', lambda e: self.finish_editing(row_id, column_no))
            self.edit_widget.bind('<FocusOut>', lambda e: self.finish_editing(row_id, column_no))
            self.edit_widget.focus_set()

    def finish_editing(self, row_id, column_no):
        new_value = self.edit_widget.get()
        self.tree.set(row_id, self.tree["columns"][column_no], new_value)
        df_row_index = self.tree.index(row_id)
        column_name = self.tree["columns"][column_no]
        self.df.at[df_row_index, column_name] = new_value
        self.edit_widget.destroy()


if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = ExcelGUI(root)
    root.mainloop()
