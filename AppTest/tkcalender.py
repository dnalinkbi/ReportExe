import tkinter as tk
from datetime import date
from tkcalendar import Calendar


class DateApp:
    def __init__(self):
        self.root = tk.Tk()
        self.date_var = tk.StringVar()
        self.date_var.set(date.today().strftime("%Y-%m-%d"))

    def create_ui(self):
        date_entry = tk.Entry(self.root, textvariable=self.date_var)
        date_entry.pack()

        pick_date_button = tk.Button(self.root, text="Pick Date", command=self.pick_date)
        pick_date_button.pack()

        self.root.mainloop()

    def pick_date(self):
        # Create a new Tkinter window for the date picker
        self.picker_window = tk.Toplevel(self.root)
        self.picker_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))

        # Create a Calendar widget and associate it with the date_var variable
        self.calendar = Calendar(self.picker_window, selectmode="day", datevar=self.date_var, firstweekday='sunday')
        self.calendar.pack()

        # Create a "Select" button to update the Entry widget with the selected date
        select_button = tk.Button(self.picker_window, text="Select", command=self.update_entry)
        select_button.pack()

    def update_entry(self):
        selected_date = self.calendar.selection_get().strftime("%Y-%m-%d")
        self.date_var.set(selected_date)
        self.root.focus_set()
        self.picker_window.destroy()

        

    def validate_date_input(self, input_text):
        if input_text == "":
            return True

        if input_text.isdigit() and len(input_text) <= 10:
            # Check if the input is in the format YYYY-MM-DD
            parts = input_text.split("-")
            if len(parts) == 3 and len(parts[0]) == 4 and len(parts[1]) == 2 and len(parts[2]) == 2:
                return True

        return False


if __name__ == "__main__":
    app = DateApp()
    app.create_ui()
