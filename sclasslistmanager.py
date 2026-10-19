from pclasses import cstoragemanager as csm
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

#Config directionary
dir = "../teststorage"

listthings = csm.getFileNames(dir)


class ListManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("List Manager")
        self.root.geometry("600x400")

        self.current_manager = None

        # Main Layout: Split into Left (Files) and Right (Names)
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Left Panel: File Selection
        left_frame = ttk.LabelFrame(main_frame, text="Select File", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.file_listbox = tk.Listbox(left_frame, selectmode=tk.SINGLE)
        self.file_listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        file_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=file_scrollbar.set)

        # Populate file list
        for filename in listthings:
            self.file_listbox.insert(tk.END, filename)

        self.file_listbox.bind("<<ListboxSelect>>", self.on_file_select)

        # Right Panel: Names Manager
        right_frame = ttk.LabelFrame(main_frame, text="Names in Selected File", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.names_listbox = tk.Listbox(right_frame, selectmode=tk.SINGLE)
        self.names_listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        names_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.names_listbox.yview)
        names_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.names_listbox.config(yscrollcommand=names_scrollbar.set)

        # Action Buttons Panel
        btn_frame = ttk.Frame(main_frame, padding=5)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.add_btn = ttk.Button(btn_frame, text="Add Name", command=self.add_name, state=tk.DISABLED)
        self.add_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.remove_btn = ttk.Button(btn_frame, text="Remove Name", command=self.remove_name, state=tk.DISABLED)
        self.remove_btn.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)

    def on_file_select(self, event):
        selection = self.file_listbox.curselection()
        if not selection:
            return

        selected_file = self.file_listbox.get(selection[0])
        self.current_manager = csm.ExcelColumnManager(selected_file, dir)

        # Enable buttons once a file is selected
        self.add_btn.config(state=tk.NORMAL)
        self.remove_btn.config(state=tk.NORMAL)

        self.refresh_names()

    def refresh_names(self):
        """Loads and updates names displayed in the names Listbox."""
        self.names_listbox.delete(0, tk.END)
        if not self.current_manager:
            return

        try:
            names = self.current_manager.getList("A")
            for name in names:
                if name is not None:  # Filter out empty/None cells
                    self.names_listbox.insert(tk.END, str(name))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file contents:\n{e}")

    def add_name(self):
        if not self.current_manager:
            return

        new_name = simpledialog.askstring("Add Name", "Enter the name to add:")
        if new_name and new_name.strip():
            try:
                self.current_manager.add2List(new_name.strip())
                self.refresh_names()
            except Exception as e:
                messagebox.showerror("Error", f"Could not add name:\n{e}")

    def remove_name(self):
        if not self.current_manager:
            return

        selection = self.names_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a name to remove.")
            return

        selected_name = self.names_listbox.get(selection[0])

        confirm = messagebox.askyesno("Confirm", f"Are you sure you want to remove '{selected_name}'?")
        if confirm:
            try:
                # Working around the loop typo in cstoragemanager.removeFromList:
                names = [n for n in self.current_manager.getList("A") if n != selected_name]
                self.current_manager.overight_list("A", names)
                self.refresh_names()
            except Exception as e:
                messagebox.showerror("Error", f"Could not remove name:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ListManagerApp(root)
    root.mainloop()