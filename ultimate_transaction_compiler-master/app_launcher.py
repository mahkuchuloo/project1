import tkinter as tk
from tkinter import ttk
from dynamic_transaction_compiler import DynamicTransactionCompiler
from giving_dashboard import GivingDashboard
# from rfm_analyzer.rfm_analyzer import RFMAnalyzer  # Old RFM Analyzer
from rfm_analyzer.final_rfm_analyzer import FinalRFMAnalyzer
# from abstract_rfm.abstract_rfm_analyzer import AbstractRFMAnalyzer  # New RFM Analyzer
from rfm_analyzer_helper import RFMAnalyzerHelper
from abstract_rfm.final_rfm_analyzer import FinalRFMAnalyzer as AbstractRFMAnalyzer
from benevity_transaction_compiler import BenevityTransactionCompiler

class AppLauncher:
    def __init__(self, master):
        self.master = master
        self.master.title("Application Launcher")
        self.master.geometry("300x200")

        self.create_widgets()

    def create_widgets(self):
        self.app_listbox = tk.Listbox(self.master, height=5)
        self.app_listbox.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

        self.apps = [
            ("Dynamic Transaction Compiler", DynamicTransactionCompiler),
            ("Benevity Transaction Compiler", BenevityTransactionCompiler),
            ("Giving Dashboard", GivingDashboard),
            ("RFM Analyzer Helper", RFMAnalyzerHelper),
            ("RFM Analyzer", FinalRFMAnalyzer)  # Using new RFM Analyzer
        ]

        for app_name, _ in self.apps:
            self.app_listbox.insert(tk.END, app_name)

        launch_button = ttk.Button(self.master, text="Launch", command=self.launch_app)
        launch_button.pack(pady=10)

    def launch_app(self):
        selection = self.app_listbox.curselection()
        if selection:
            index = selection[0]
            app_name, app_class = self.apps[index]
            self.master.withdraw()  # Hide the main window
            app_window = tk.Toplevel(self.master)
            app = app_class(app_window)
            app_window.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app_window))

    def on_app_close(self, app_window):
        app_window.destroy()
        self.master.deiconify()  # Show the main window again

if __name__ == "__main__":
    root = tk.Tk()
    app = AppLauncher(root)
    root.mainloop()
