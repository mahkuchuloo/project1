from ttkbootstrap import Style
import tkinter as tk

def safe_style_init(theme_name):
    try:
        return Style(theme=theme_name)
    except tk._tkinter.TclError:
        # If there's an error with the theme, fallback to default
        return Style()

def apply_custom_dark_theme():
    style = safe_style_init("darkly")
    
    try:
        style.configure(".", 
            background="#2b3e50",
            foreground="#ffffff",
            fieldbackground="#1a2632",
            font=("TkDefaultFont", 10)
        )
        
        style.configure("TButton",
            padding=[10, 5],
            background="#007bff",
            foreground="#ffffff"
        )
        
        style.configure("TEntry",
            fieldbackground="#1a2632",
            foreground="#ffffff"
        )
        
        style.configure("TLabel",
            background="#2b3e50",
            foreground="#ffffff"
        )
        
        # Enhanced Treeview styling
        style.configure("Treeview",
            background="#1a2632",
            foreground="#ffffff",
            fieldbackground="#1a2632",
            rowheight=30,  # Increased for better checkbox visibility
            font=("TkDefaultFont", 10)
        )
        
        # Checkbox item styling
        style.configure("Treeview.Item",
            padding=[5, 2],  # Add padding around items
            background="#1a2632",
            foreground="#ffffff"
        )
        
        style.map("Treeview.Item",
            background=[("selected", "#007bff")],
            foreground=[("selected", "#ffffff")]
        )
        
        style.configure("TProgressbar",
            background="#007bff",
            troughcolor="#1a2632",
            bordercolor="#495057"
        )
        
        style.configure("TFrame",
            background="#2b3e50"
        )
        
        style.map("TButton",
            background=[("active", "#0056b3"), ("hover", "#0069d9")],
            foreground=[("active", "#ffffff"), ("hover", "#ffffff")]
        )
        
        style.map("Treeview",
            background=[("selected", "#007bff")],
            foreground=[("selected", "#ffffff")]
        )
    except Exception:
        # If there's an error applying specific styles, continue with basic styling
        pass
    
    return style

def apply_clean_ui_theme():
    style = safe_style_init("cosmo")
    
    try:
        # Expanded soft color palette
        colors = {
            'bg': '#f0f4f8',
            'fg': '#4a5568',
            'field': '#ffffff',
            'primary': '#63b3ed',
            'primary_hover': '#90cdf4',
            'primary_active': '#4299e1',
            'success': '#68d391',
            'success_hover': '#9ae6b4',
            'success_active': '#48bb78',
            'info': '#b794f4',
            'info_hover': '#d6bcfa',
            'info_active': '#9f7aea',
            'warning': '#f6ad55',
            'warning_hover': '#fbd38d',
            'warning_active': '#ed8936',
            'danger': '#fc8181',
            'danger_hover': '#feb2b2',
            'danger_active': '#f56565',
            'border': '#e2e8f0',
            'hover': '#edf2f7',
            'disabled': '#e2e8f0',
            'muted': '#718096'
        }
        
        # Global styles
        style.configure(".", 
            background=colors['bg'],
            foreground=colors['fg'],
            fieldbackground=colors['field'],
            font=("Segoe UI", 9),
            relief="flat"
        )
        
        # Soft button styles
        style.configure("TButton",
            padding=[14, 8],
            background=colors['primary'],
            foreground='white',
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 9)
        )
        
        # Button variants
        for variant, color in [
            ("success", 'success'),
            ("info", 'info'),
            ("warning", 'warning'),
            ("danger", 'danger')
        ]:
            style.configure(f"{variant}.TButton",
                background=colors[color],
                foreground='white'
            )
            style.map(f"{variant}.TButton",
                background=[
                    ("active", colors[f"{color}_active"]),
                    ("hover", colors[f"{color}_hover"]),
                    ("disabled", colors['disabled'])
                ],
                foreground=[("disabled", colors['muted'])]
            )
        
        # Clean input fields
        style.configure("TEntry",
            fieldbackground=colors['field'],
            foreground=colors['fg'],
            borderwidth=1,
            relief="solid",
            padding=[6, 4]
        )
        
        # Labels
        style.configure("TLabel",
            background=colors['bg'],
            foreground=colors['fg'],
            padding=[2, 1]
        )
        
        # Enhanced Treeview styling
        style.configure("Treeview",
            background=colors['field'],
            foreground=colors['fg'],
            fieldbackground=colors['field'],
            rowheight=30,  # Increased for better checkbox visibility
            borderwidth=1,
            relief="solid",
            padding=[4, 2],
            font=("Segoe UI", 9)
        )
        
        # Checkbox item styling
        style.configure("Treeview.Item",
            padding=[5, 2],
            background=colors['field'],
            foreground=colors['fg']
        )
        
        style.map("Treeview.Item",
            background=[
                ("selected", colors['primary_hover']),
                ("hover", colors['hover'])
            ],
            foreground=[
                ("selected", colors['primary_active']),
                ("hover", colors['fg'])
            ]
        )
        
        # Progress bar
        style.configure("TProgressbar",
            background=colors['primary'],
            troughcolor=colors['border'],
            borderwidth=0,
            thickness=6
        )
        
        # Frames
        style.configure("TFrame",
            background=colors['bg'],
            borderwidth=0
        )
        
        # Button states
        style.map("TButton",
            background=[
                ("active", colors['primary_active']),
                ("hover", colors['primary_hover']),
                ("disabled", colors['disabled'])
            ],
            foreground=[("disabled", colors['muted'])]
        )
        
        # Treeview selection
        style.map("Treeview",
            background=[
                ("selected", colors['primary_hover']),
                ("hover", colors['hover'])
            ],
            foreground=[
                ("selected", colors['primary_active']),
                ("hover", colors['fg'])
            ]
        )
        
        # Notebook tabs
        style.configure("TNotebook.Tab",
            padding=[10, 6],
            background=colors['bg'],
            foreground=colors['fg'],
            borderwidth=0
        )
        
        style.map("TNotebook.Tab",
            background=[
                ("selected", colors['field']),
                ("active", colors['hover'])
            ],
            foreground=[
                ("selected", colors['primary']),
                ("active", colors['fg'])
            ]
        )
        
        # Title styles
        style.configure("Title.TLabel",
            font=("Segoe UI Semibold", 12),
            foreground=colors['fg'],
            padding=[0, 5]
        )
        
        style.configure("Subtitle.TLabel",
            font=("Segoe UI", 10),
            foreground=colors['muted'],
            padding=[0, 3]
        )
        
        # Card frame style
        style.configure("Card.TFrame",
            background=colors['field'],
            borderwidth=1,
            relief="solid",
            bordercolor=colors['border']
        )
    except Exception:
        # If there's an error applying specific styles, continue with basic styling
        pass
    
    return style

def apply_minimal_round_theme():
    style = safe_style_init("cosmo")
    
    try:
        # Modern, minimal color palette
        colors = {
            'bg': '#ffffff',
            'fg': '#2c3e50',
            'accent': '#8b5cf6',
            'button': {
                'primary': '#8b5cf6',
                'secondary': '#64748b',
                'success': '#10b981',
                'info': '#0ea5e9',
                'warning': '#f59e0b',
                'danger': '#ef4444'
            },
            'hover': {
                'primary': '#a78bfa',
                'secondary': '#94a3b8',
                'success': '#34d399',
                'info': '#38bdf8',
                'warning': '#fbbf24',
                'danger': '#f87171'
            },
            'border': '#e2e8f0',
            'input_bg': '#f8fafc',
            'disabled': '#f1f5f9',
            'selected': '#e0e7ff'
        }
        
        # Global styles
        style.configure(".",
            background=colors['bg'],
            foreground=colors['fg'],
            font=("Segoe UI", 9),
            relief="flat"
        )
        
        # Rounded button base style
        button_base = {
            "padding": [16, 8],
            "relief": "flat",
            "borderwidth": 0,
            "font": ("Segoe UI", 9),
        }
        
        # Configure button variants
        button_variants = {
            "": colors['button']['primary'],
            "Secondary": colors['button']['secondary'],
            "Success": colors['button']['success'],
            "Info": colors['button']['info'],
            "Warning": colors['button']['warning'],
            "Danger": colors['button']['danger']
        }
        
        for variant, color in button_variants.items():
            style_name = f"{variant}.TButton" if variant else "TButton"
            hover_color = colors['hover'][variant.lower() or 'primary']
            
            style.configure(style_name,
                **button_base,
                background=color,
                foreground='white'
            )
            
            style.map(style_name,
                background=[
                    ("active", hover_color),
                    ("hover", hover_color),
                    ("disabled", colors['disabled'])
                ],
                foreground=[("disabled", "#94a3b8")]
            )
        
        # Modern entry fields
        style.configure("TEntry",
            fieldbackground=colors['input_bg'],
            foreground=colors['fg'],
            padding=[8, 6],
            relief="flat",
            borderwidth=1
        )
        
        # Clean labels
        style.configure("TLabel",
            background=colors['bg'],
            foreground=colors['fg'],
            padding=[4, 2]
        )
        
        # Enhanced Treeview styling
        style.configure("Treeview",
            background=colors['bg'],
            foreground=colors['fg'],
            fieldbackground=colors['bg'],
            rowheight=32,  # Increased for better checkbox visibility
            padding=[8, 4],  # Increased padding
            relief="flat",
            borderwidth=0,
            font=("Segoe UI", 9)
        )
        
        # Checkbox item styling with modern look
        style.configure("Treeview.Item",
            padding=[8, 4],
            background=colors['bg'],
            foreground=colors['fg']
        )
        
        style.map("Treeview.Item",
            background=[
                ("selected", colors['selected']),
                ("hover", colors['input_bg'])
            ],
            foreground=[
                ("selected", colors['accent']),
                ("hover", colors['fg'])
            ]
        )
        
        # Sleek progress bar
        style.configure("TProgressbar",
            background=colors['accent'],
            troughcolor=colors['border'],
            borderwidth=0,
            thickness=4
        )
        
        # Clean frames
        style.configure("TFrame",
            background=colors['bg'],
            borderwidth=0
        )
        
        # Modern notebook tabs
        style.configure("TNotebook.Tab",
            padding=[12, 6],
            background=colors['bg'],
            foreground=colors['fg'],
            borderwidth=0
        )
        
        style.map("Treeview",
            background=[
                ("selected", colors['selected']),
                ("hover", colors['input_bg'])
            ],
            foreground=[
                ("selected", colors['accent']),
                ("hover", colors['fg'])
            ]
        )
        
        style.map("TNotebook.Tab",
            background=[
                ("selected", colors['bg']),
                ("active", colors['input_bg'])
            ],
            foreground=[
                ("selected", colors['accent']),
                ("active", colors['fg'])
            ]
        )
        
        # Title styles
        style.configure("Title.TLabel",
            font=("Segoe UI", 16, "normal"),
            foreground=colors['fg'],
            padding=[0, 8]
        )
        
        style.configure("Subtitle.TLabel",
            font=("Segoe UI", 11),
            foreground=colors['button']['secondary'],
            padding=[0, 4]
        )
    except Exception:
        # If there's an error applying specific styles, continue with basic styling
        pass
    
    return style
