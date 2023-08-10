import tkinter as tk

class InformationWindow:
    def __init__(self):
        self.window = tk.Toplevel()
        self.window.title("Information")
        
        self.entries = []
        for i in range(10):
            label = tk.Label(self.window, text="Information {}: ".format(i+1))
            entry = tk.Entry(self.window)
            label.pack()
            entry.pack()
            self.entries.append(entry)
        
        self.button_confirm = tk.Button(self.window, text="Confirm", command=self.get_information)
        self.button_confirm.pack()
    
    def get_information(self):
        information_list = [entry.get() for entry in self.entries]
        print("Information List:", information_list)


class LoginWindow:
    def __init__(self):
        self.window = tk.Toplevel()
        self.window.title("Login")
        
        self.label_id = tk.Label(self.window, text="ID:")
        self.entry_id = tk.Entry(self.window)
        
        self.label_pw = tk.Label(self.window, text="Password:")
        self.entry_pw = tk.Entry(self.window, show="*")
        
        self.button_login = tk.Button(self.window, text="Login", command=self.login)
        
        self.label_id.pack()
        self.entry_id.pack()
        self.label_pw.pack()
        self.entry_pw.pack()
        self.button_login.pack()
    
    def login(self):
        username = self.entry_id.get()
        password = self.entry_pw.get()
        
        # Add your login logic here
        # For demonstration purposes, let's just print the entered credentials
        print("Username:", username)
        print("Password:", password)
        
        if username and password:
            self.window.destroy()  # Close the login window
            information_window = InformationWindow()


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Main Window")
        
        self.button_open_login = tk.Button(self.root, text="Open Login Window", command=self.open_login_window)
        
        self.button_open_login.pack()
    
    def open_login_window(self):
        login_window = LoginWindow()


# Create an instance of the MainWindow class
main_window = MainWindow()

# Start the Tkinter event loop
main_window.root.mainloop()
