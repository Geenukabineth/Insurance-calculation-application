import tkinter as tk
from tkinter import messagebox
import subprocess


class LoginApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Login Interface")
        self.root.geometry("300x250")

        # Center the window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 300) // 2
        y = (screen_height - 250) // 2
        self.root.geometry(f"300x250+{x}+{y}")

        # Create main frame
        self.frame = tk.Frame(self.root, padx=20, pady=20)
        self.frame.pack(expand=True, fill="both")

        # Create and pack the title label
        title_label = tk.Label(self.frame, text="Login", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # Username field
        self.username_label = tk.Label(self.frame, text="Username:")
        self.username_label.pack(pady=(0, 5))

        self.username_entry = tk.Entry(self.frame)
        self.username_entry.pack(pady=(0, 10))

        # Password field
        self.password_label = tk.Label(self.frame, text="Password:")
        self.password_label.pack(pady=(0, 5))

        self.password_entry = tk.Entry(self.frame, show="*")
        self.password_entry.pack(pady=(0, 20))

        # Login button
        self.login_button = tk.Button(self.frame, text="Login", command=self.login, width=20, height=2)
        self.login_button.pack()

        # Bind Enter key to login function
        self.root.bind("<Return>", lambda event: self.login())

        # Focus on username entry
        self.username_entry.focus()

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if not username or not password:
            messagebox.showerror("Error", "Please enter both username and password")
            return

        if username == "admin" and password == "12345":
            messagebox.showinfo("Success", "Welcome!")
            self.root.destroy()
            self.launch_interface2()
        else:
            messagebox.showerror("Error", "Invalid username or password")
            self.password_entry.delete(0, tk.END)
            self.password_entry.focus()

    def launch_interface2(self):
        try:
            subprocess.Popen(
                ["python", r"C:\GI_Automation\GUI\interface2.py"],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch interface2.py: {str(e)}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        app = LoginApp()
        app.run()
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
