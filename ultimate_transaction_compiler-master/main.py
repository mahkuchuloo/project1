import tkinter as tk
import logging
import traceback
import sys
from app_launcher import AppLauncher
from custom_theme import apply_custom_dark_theme, apply_clean_ui_theme, apply_minimal_round_theme

class ThemeManager:
    def __init__(self, root):
        self.root = root
        self.current_theme = "minimal"  # Set Minimal Round as default
        
        # Create menu bar
        self.menu_bar = tk.Menu(root)
        root.config(menu=self.menu_bar)
        
        # Create Theme menu
        self.theme_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Theme", menu=self.theme_menu)
        self.theme_menu.add_command(label="Minimal Round", command=self.set_minimal_theme)
        self.theme_menu.add_command(label="Clean UI", command=self.set_clean_theme)
        self.theme_menu.add_command(label="Dark Theme", command=self.set_dark_theme)
        
        # Apply default theme (Minimal Round)
        self.style = apply_minimal_round_theme()
    
    def set_dark_theme(self):
        if self.current_theme != "dark":
            self.style = apply_custom_dark_theme()
            self.current_theme = "dark"
    
    def set_clean_theme(self):
        if self.current_theme != "clean":
            self.style = apply_clean_ui_theme()
            self.current_theme = "clean"
            
    def set_minimal_theme(self):
        if self.current_theme != "minimal":
            self.style = apply_minimal_round_theme()
            self.current_theme = "minimal"

if __name__ == "__main__":
    logging.basicConfig(filename='application.log', level=logging.DEBUG, 
                        format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info("Starting Application Launcher")
    try:
        root = tk.Tk()
        root.title("Transaction Processing Suite")
        root.geometry("800x600")
        
        # Initialize theme manager with Minimal Round as default
        theme_manager = ThemeManager(root)
        
        app_launcher = AppLauncher(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Unhandled exception: {str(e)}")
        logging.error(traceback.format_exc())
        print(f"An unhandled exception occurred: {str(e)}")
        print("Please check the log file for more details.")
        sys.exit(1)
