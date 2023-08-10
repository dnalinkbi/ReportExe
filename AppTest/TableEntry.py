import tkinter as tk


class DynamicInputWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.num_fields_entry = None
        self.submit_button = None
        self.input_fields = []

        self.create_ui()

    def create_ui(self):
        label = tk.Label(self.root, text="Enter the number of fields:")
        label.pack()

        self.num_fields_entry = tk.Entry(self.root)
        self.num_fields_entry.pack()

        self.submit_button = tk.Button(self.root, text="Create Fields", command=self.create_input_fields)
        self.submit_button.pack()

        self.root.mainloop()

    def create_input_fields(self):
        num_fields = int(self.num_fields_entry.get())

        # Destroy previously created fields (if any)
        for field in self.input_fields:
            field.destroy()
        self.input_fields = []

        # Create new input fields
        for i in range(num_fields):
            field = tk.Entry(self.root)
            field.pack()
            self.input_fields.append(field)


if __name__ == "__main__":
    app = DynamicInputWindow()
