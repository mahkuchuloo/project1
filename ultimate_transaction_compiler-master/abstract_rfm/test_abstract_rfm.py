import tkinter as tk
from abstract_rfm_analyzer import AbstractRFMAnalyzer

def main():
    root = tk.Tk()
    root.title("Abstract RFM Analyzer")
    root.geometry("1024x768")
    
    app = AbstractRFMAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
