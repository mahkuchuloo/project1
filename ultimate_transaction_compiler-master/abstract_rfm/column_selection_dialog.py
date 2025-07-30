import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from .column_config_manager import ColumnConfigManager

class ColumnSelectionDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Column Selection")
        self.geometry("600x800")
        
        # Modern checkbox symbols
        self.UNCHECKED = "⬜"  # White square
        self.CHECKED = "✅"    # Green checkmark
        self.MIXED = "❎"      # Half-filled square
        
        # Initialize column config manager
        self.config_manager = ColumnConfigManager()
        
        # Create main frame
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add description
        description = """
        Select which columns to include in the RFM analysis output.
        Columns are organized by category for easier selection.
        """
        desc_label = ttk.Label(main_frame, text=description, wraplength=550, justify=tk.LEFT)
        desc_label.pack(fill=tk.X, pady=(0, 10))
        
        # Create tree view with custom style
        style = ttk.Style()
        style.configure("Custom.Treeview",
            font=('Segoe UI', 11),  # Larger font
            rowheight=35            # Taller rows for better visibility
        )
        
        # Configure tag for better visibility
        style.configure("Custom.Treeview", indent=20)  # Increase indentation
        
        self.tree = ttk.Treeview(main_frame, selectmode='none', style="Custom.Treeview")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Dictionary to store checkbox states
        self.checkboxes = {}
        
        # Initialize tree structure
        self._initialize_tree()
        
        # Add buttons frame with modern style
        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Style for modern buttons
        style.configure("Action.TButton", 
                       padding=10, 
                       font=('Segoe UI', 10))
        
        ttk.Button(button_frame, text="Select All", 
                  command=self.select_all, 
                  style="Action.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Deselect All", 
                  command=self.deselect_all, 
                  style="Action.TButton").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Save", 
                  command=self.save_selection, 
                  style="Action.TButton").pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(button_frame, text="Cancel", 
                  command=self.destroy, 
                  style="Action.TButton").pack(side=tk.RIGHT, padx=5)
        
        # Bind checkbox click
        self.tree.tag_bind('checkbox', '<Button-1>', self.toggle_checkbox)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Load existing configuration
        self.load_config()
        
    def _initialize_tree(self):
        """Initialize the tree structure with all column categories"""
        self.column_structure = {
            "Basic Customer Info": [
                "Email", "Phone", "Address", "City", "State", "Zip",
                "Total Number of Gifts", "Lifetime Giving"
            ],
            "RFM Criteria": [
                "Recency Criteria", "Frequency Criteria", "Monetary Criteria",
                "RFM Score", "RFM Percentile", "Recency Score", "Frequency Score",
                "Monetary Score", "Recency Percentile", "Frequency Percentile",
                "Monetary Percentile"
            ],
            "First Gift Information": [
                "First Gift Date", "First Gift Amount", "First Gift Platform",
                "First Gift Campaign Name", "First Gift Appeal Name",
                "First Gift Date Range A", "First Gift Date Range B",
                "First Gift Amount Range"
            ],
            "Last Gift Information": [
                "Last Gift Date", "Last Gift Amount", "Last Gift Platform",
                "Last Gift Campaign Name", "Last Gift Appeal Name",
                "Last Gift Date Range", "Last Gift Amount Range"
            ],
            "Largest Gift Information": [
                "Largest Gift Date", "Largest Gift Amount", "Largest Gift Platform",
                "Largest Gift Campaign Name", "Largest Gift Appeal Name",
                "Largest Gift Date Range", "Largest Gift Amount Range"
            ],
            "Monthly Gift Information": [
                "Last Monthly Gift Date", "Last Monthly Gift Amount",
                "Last Monthly Gift Date Range", "Last Monthly Giving Amount Range",
                "Digital Monthly Indicator"
            ],
            "Platform Information": [
                "Primary Giving Platform", "Primary Giving Platform %"
            ],
            "Giving Segments": [
                "Giving Segment A", "Giving Segment B"
            ]
        }
        
        # Add categories and their columns
        for category, columns in self.column_structure.items():
            category_id = self.tree.insert('', 'end', text=f'{self.UNCHECKED} {category}', tags=('checkbox', 'category'))
            self.checkboxes[category_id] = False
            
            for column in columns:
                column_id = self.tree.insert(category_id, 'end', text=f'{self.UNCHECKED} {column}', tags=('checkbox', 'item'))
                self.checkboxes[column_id] = False
    
    def toggle_checkbox(self, event):
        """Toggle checkbox state when clicked"""
        item_id = self.tree.identify('item', event.x, event.y)
        self.checkboxes[item_id] = not self.checkboxes[item_id]
        
        # Update checkbox display
        current_text = self.tree.item(item_id)['text']
        new_text = f'{self.CHECKED} {current_text[2:]}' if self.checkboxes[item_id] else f'{self.UNCHECKED} {current_text[2:]}'
        self.tree.item(item_id, text=new_text)
        
        # If it's a category, update all children
        for child in self.tree.get_children(item_id):
            self.checkboxes[child] = self.checkboxes[item_id]
            child_text = self.tree.item(child)['text']
            new_child_text = f'{self.CHECKED} {child_text[2:]}' if self.checkboxes[item_id] else f'{self.UNCHECKED} {child_text[2:]}'
            self.tree.item(child, text=new_child_text)
        
        # Update parent state based on children
        parent = self.tree.parent(item_id)
        if parent:
            children = self.tree.get_children(parent)
            all_checked = all(self.checkboxes[child] for child in children)
            any_checked = any(self.checkboxes[child] for child in children)
            
            parent_text = self.tree.item(parent)['text']
            if all_checked:
                new_parent_text = f'{self.CHECKED} {parent_text[2:]}'
                self.checkboxes[parent] = True
            elif any_checked:
                new_parent_text = f'{self.CHECKED} {parent_text[2:]}'
                self.checkboxes[parent] = None
            else:
                new_parent_text = f'{self.UNCHECKED} {parent_text[2:]}'
                self.checkboxes[parent] = False
            self.tree.item(parent, text=new_parent_text)
    
    def select_all(self):
        """Select all columns"""
        for item_id in self.checkboxes:
            self.checkboxes[item_id] = True
            text = self.tree.item(item_id)['text']
            self.tree.item(item_id, text=f'{self.CHECKED} {text[2:]}')
    
    def deselect_all(self):
        """Deselect all columns"""
        for item_id in self.checkboxes:
            self.checkboxes[item_id] = False
            text = self.tree.item(item_id)['text']
            self.tree.item(item_id, text=f'{self.UNCHECKED} {text[2:]}')
    
    def get_selected_columns(self):
        """Get list of selected columns"""
        selected = []
        for item_id, checked in self.checkboxes.items():
            if checked and not self.tree.get_children(item_id):  # Only include leaf nodes
                selected.append(self.tree.item(item_id)['text'][2:])  # Remove checkbox symbol
        return selected
    
    def load_config(self):
        """Load existing configuration"""
        selected_columns = self.config_manager.get_columns()
        if selected_columns:
            # Update checkboxes based on configuration
            for item_id in self.checkboxes:
                text = self.tree.item(item_id)['text'][2:]  # Remove checkbox symbol
                if text in selected_columns:
                    self.checkboxes[item_id] = True
                    self.tree.item(item_id, text=f'{self.CHECKED} {text}')
                    
                    # Update parent state
                    parent = self.tree.parent(item_id)
                    if parent:
                        parent_text = self.tree.item(parent)['text'][2:]
                        self.tree.item(parent, text=f'{self.CHECKED} {parent_text}')
                        self.checkboxes[parent] = None
    
    def save_selection(self):
        """Save selected columns to config file"""
        selected = self.get_selected_columns()
        self.config_manager.set_columns(selected)
        # messagebox.showinfo("Success", "Column selection saved successfully!")
        self.destroy()
