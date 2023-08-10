import tkinter as tk


class ExcelPasteWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.num_fields_entry = None
        self.create_fields_button = None
        self.clear_fields_button = None
        self.entries = []

        self.create_ui()

    def create_ui(self):
        label = tk.Label(self.root, text="Enter the number of fields:")
        label.pack()

        self.num_fields_entry = tk.Entry(self.root)
        self.num_fields_entry.pack()

        self.create_fields_button = tk.Button(self.root, text="Create Fields", command=self.create_input_fields)
        self.create_fields_button.pack()

        self.clear_fields_button = tk.Button(self.root, text="Clear Fields", command=self.clear_input_fields)
        self.clear_fields_button.pack()

        self.root.bind("<Control-v>", self.handle_paste)

        self.root.mainloop()

    def create_input_fields(self):
        num_fields = int(self.num_fields_entry.get())

        # Destroy previously created fields (if any)
        self.clear_input_fields()

        # Create new input fields
        for _ in range(num_fields):
            entry = tk.Entry(self.root, width=30)
            entry.pack()
            self.entries.append(entry)

    def handle_paste(self, event):
        content = self.root.clipboard_get()

        lines = content.split("\n")  # Split content by newline

        for i, line in enumerate(lines):
            if i < len(self.entries):
                self.entries[i].delete(0, tk.END)  # Clear existing entry
                self.entries[i].insert(0, line)  # Paste content into entry

    def clear_input_fields(self):
        for entry in self.entries:
            entry.destroy()
        self.entries = []


if __name__ == "__main__":
    app = ExcelPasteWindow()
