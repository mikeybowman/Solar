
"""
Solar - Open Source Game Design Document Builder
Because every great game starts with great documentation

Originally built by: Mikey LaBrecque
Created for: Dead Orbit Studios
License: Open Source - feel free to modify and share!
Compatibility: Cross-platform Python (Windows, Mac, Linux)
Dependencies: python-docx for Word export functionality

This tool was built with love for the indie game dev community.
Got ideas for improvements? The code is yours to hack!
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import os
from datetime import datetime
from pathlib import Path
import webbrowser
import sys

try:
    from docx import Document
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

class CollapsibleEntry:
    """
    This handles the expand/collapse behavior for individual entries.
    I wanted each entry to be collapsible so you don't get overwhelmed
    when you have tons of characters or items to manage.
    """
    
    def __init__(self, parent, title, entry_data, delete_callback, fields):
        self.parent = parent
        self.delete_callback = delete_callback
        self.fields = fields
        self.is_expanded = False
        self.widgets = {}
        
        # Main container for this entry
        self.main_frame = tk.Frame(parent, relief=tk.RIDGE, bd=2, bg='#f0f0f0')
        self.main_frame.pack(fill=tk.X, padx=5, pady=3)
        
        # Header section - this stays visible when collapsed
        self.header_frame = tk.Frame(self.main_frame, bg='#e0e0e0', relief=tk.RAISED, bd=1)
        self.header_frame.pack(fill=tk.X)
        
        # Toggle button - using classic Windows style
        self.toggle_button = tk.Button(self.header_frame, text="►", width=3,
                                     command=self.toggle_expanded, bg='#d0d0d0', 
                                     relief=tk.RAISED, font=('Arial', 8))
        self.toggle_button.pack(side=tk.LEFT, padx=3, pady=3)
        
        # Entry title - updates automatically when you type in the name field
        self.title_var = tk.StringVar()
        self.update_title(title)
        self.title_label = tk.Label(self.header_frame, textvariable=self.title_var, 
                                  font=('Arial', 9, 'bold'), bg='#e0e0e0', fg='#000080')
        self.title_label.pack(side=tk.LEFT, padx=8, pady=3)
        
        # Delete button - keeping it simple and functional
        delete_btn = tk.Button(self.header_frame, text="Delete", 
                             command=self.delete_entry, bg='#ff6666', 
                             relief=tk.RAISED, font=('Arial', 8))
        delete_btn.pack(side=tk.RIGHT, padx=3, pady=3)
        
        # Content area - this gets hidden/shown when collapsing
        self.content_frame = tk.LabelFrame(self.main_frame, text="", 
                                         bg='#f8f8f8', relief=tk.GROOVE, bd=2)
        
        # Build all the input fields
        self.create_content_widgets(entry_data)
        
        # Hook up the name field to update the title automatically
        if 'name' in self.widgets:
            self.widgets['name'].bind('<KeyRelease>', self.on_name_changed)
    
    def create_content_widgets(self, entry_data=None):
        """Build all the input fields for this entry"""
        for i, (label, key, widget_type) in enumerate(self.fields):
            # Field label
            tk.Label(self.content_frame, text=label, bg='#f8f8f8', 
                    font=('Arial', 8, 'bold')).grid(row=i*2, column=0, sticky='w', 
                                                   pady=(5, 2), padx=8)
            
            # Field input widget
            if widget_type == "entry":
                widget = tk.Entry(self.content_frame, width=70, relief=tk.SUNKEN, bd=1)
                widget.grid(row=i*2+1, column=0, sticky='ew', padx=8, pady=(0, 5))
                if entry_data and key in entry_data:
                    widget.insert(0, entry_data[key])
            elif widget_type == "text":
                widget = scrolledtext.ScrolledText(self.content_frame, height=4, width=70,
                                                 bg='white', relief=tk.SUNKEN, bd=1,
                                                 wrap=tk.WORD)
                widget.grid(row=i*2+1, column=0, sticky='ew', padx=8, pady=(0, 5))
                if entry_data and key in entry_data:
                    widget.insert('1.0', entry_data[key])
            
            self.widgets[key] = widget
        
        self.content_frame.columnconfigure(0, weight=1)
    
    def toggle_expanded(self):
        """Handle the expand/collapse button click"""
        if self.is_expanded:
            self.collapse()
        else:
            self.expand()
    
    def expand(self):
        """Show all the fields for editing"""
        self.content_frame.pack(fill=tk.X, pady=(0, 3))
        self.toggle_button.config(text="▼")
        self.is_expanded = True
    
    def collapse(self):
        """Hide the fields, show just the header"""
        self.content_frame.pack_forget()
        self.toggle_button.config(text="►")
        self.is_expanded = False
    
    def update_title(self, title):
        """Update the title shown in the header"""
        if not title.strip():
            title = "Untitled Entry"
        # Keep it reasonably short so the UI doesn't get wonky
        if len(title) > 40:
            title = title[:37] + "..."
        self.title_var.set(title)
    
    def on_name_changed(self, event=None):
        """Called whenever someone types in the name field"""
        if 'name' in self.widgets:
            name = self.widgets['name'].get()
            self.update_title(name)
    
    def delete_entry(self):
        """Handle the delete button - ask for confirmation first"""
        if messagebox.askyesno("Delete Entry", 
                              "Are you sure you want to delete this entry?\n\nThis cannot be undone."):
            self.main_frame.destroy()
            self.delete_callback(self)
    
    def get_data(self):
        """Extract all the data from the input fields"""
        data = {}
        for key, widget in self.widgets.items():
            if isinstance(widget, tk.Entry):
                data[key] = widget.get()
            elif isinstance(widget, scrolledtext.ScrolledText):
                # Get all the text, but strip off the automatic newline at the end
                data[key] = widget.get('1.0', tk.END).rstrip('\n')
        return data

class SettingsManager:
    """
    Handles saving and loading user preferences.
    Things like window size, last project folder, etc.
    Also handles the configurable studio name and author info.
    """
    
    def __init__(self):
        self.settings_file = self.get_settings_path()
        self.settings = self.load_settings()
    
    def get_settings_path(self):
        """Figure out where to save the settings file"""
        # If running as an exe, save next to the executable
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            # If running as a script, save next to the script
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        return os.path.join(app_dir, 'solar_settings.json')
    
    def load_settings(self):
        """Load settings from disk, or create defaults if none exist"""
        default_settings = {
            'last_project_path': '',
            'window_geometry': '1000x700',  # Slightly smaller for that classic feel
            'recent_projects': [],
            'default_studio': 'Your Game Studio',  # User can change this
            'default_author': 'Game Designer',     # User can change this too
            'theme': 'classic'  # Keeping it retro!
        }
        
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    # Merge with defaults to make sure we have all the keys
                    default_settings.update(loaded_settings)
            return default_settings
        except Exception:
            # If something goes wrong, just use defaults
            return default_settings
    
    def save_settings(self):
        """Write settings back to disk"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2)
        except Exception:
            # Fail silently - don't want to annoy users with error dialogs
            pass
    
    def get(self, key, default=None):
        """Get a setting value"""
        return self.settings.get(key, default)
    
    def set(self, key, value):
        """Set a setting value and save immediately"""
        self.settings[key] = value
        self.save_settings()

class SolarGDDBuilder:
    def __init__(self, root):
        self.root = root
        self.root.title("Solar - Open Source GDD Builder")
        
        # Set the window icon
        try:
            # Try to load the icon file
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'solar.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            # If icon loading fails, continue without it
            pass
        
        # Load up our settings and preferences
        self.settings = SettingsManager()
        
        # Set window size and position from saved settings
        geometry = self.settings.get('window_geometry', '1000x700')
        self.root.geometry(geometry)
        
        # Classic Windows look - no fancy dark themes here!
        self.root.configure(bg='#f0f0f0')  # That classic light gray
        
        # Set up our data structures
        self.current_project_path = self.settings.get('last_project_path', None)
        self.project_data = self.create_empty_project()
        
        # Keep track of all our collapsible entries
        self.collapsible_entries = {
            'design_pillars': [],
            'mechanics': [],
            'levels': [],
            'characters': [],
            'items': [],
            'audio': [],
            'art_style': [],
            'technical': []
        }
        
        # Set up the classic Windows styling
        self.setup_classic_styling()
        
        # Build the main interface
        self.create_main_interface()
        
        # If we had a project path saved, show it
        if self.current_project_path and os.path.exists(self.current_project_path):
            self.update_project_display()
        
        # Save settings when the user closes the window
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Let the user know if Word export won't work
        if not DOCX_AVAILABLE:
            messagebox.showwarning("Word Export Unavailable", 
                                 "python-docx module not found.\n\n"
                                 "Word document export will be disabled.\n"
                                 "To enable it, install with: pip install python-docx")
    
    def on_closing(self):
        """Clean up when the user closes the app"""
        # Save the current window size and position
        self.settings.set('window_geometry', self.root.geometry())
        
        # Save the project path if we have one
        if self.current_project_path:
            self.settings.set('last_project_path', self.current_project_path)
        
        self.root.destroy()
    
    def setup_classic_styling(self):
        """Set up that classic early 2000s Windows look"""
        # Using the built-in tkinter widgets with classic styling
        # None of that fancy ttk stuff - we want buttons that look like buttons!
        
        # Set default button styling
        self.button_style = {
            'relief': tk.RAISED,
            'bd': 2,
            'bg': '#d0d0d0',
            'font': ('Arial', 8),
            'padx': 8,
            'pady': 2
        }
        
        # Entry field styling
        self.entry_style = {
            'relief': tk.SUNKEN,
            'bd': 1,
            'bg': 'white'
        }
        
        # Label styling
        self.label_style = {
            'bg': '#f0f0f0',
            'font': ('Arial', 8)
        }
        
        # Title label styling
        self.title_style = {
            'bg': '#f0f0f0',
            'font': ('Arial', 10, 'bold'),
            'fg': '#000080'  # That classic blue color
        }
    
    def create_empty_project(self):
        """Set up the data structure for a new project"""
        return {
            "metadata": {
                "project_name": "",
                "version": "1.1.0",
                "created_date": datetime.now().isoformat(),
                "last_modified": datetime.now().isoformat(),
                "author": self.settings.get('default_author', 'Game Designer'),
                "studio": self.settings.get('default_studio', 'Your Game Studio')
            },
            "introduction": {
                "project_name": "",
                "game_title": "",
                "genre": "",
                "platform": "",
                "target_audience": "",
                "game_overview": "",
                "unique_selling_points": "",
                "inspiration": "",
                "development_timeline": ""
            },
            "design_pillars": [],
            "mechanics": [],
            "levels": [],
            "characters": [],
            "items": [],
            "audio": [],
            "art_style": [],
            "technical": []
        }
    
    def create_main_interface(self):
        """Build the main window interface"""
        # Create the menu bar first
        self.create_menu_bar()
        
        # Main container frame
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=5)
        
        # Project info section at the top
        self.create_project_info_frame(main_frame)
        
        # Tabbed interface for the different sections
        self.create_notebook(main_frame)
        
        # Create all the tabs
        self.create_tabs()
    
    def create_menu_bar(self):
        """Set up the classic menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project...", command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label="Save Project", command=self.save_project)
        file_menu.add_command(label="Save Project As...", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Set Project Folder...", command=self.set_project_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Export menu
        export_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Export", menu=export_menu)
        export_menu.add_command(label="Export to Word Document...", command=self.export_to_word)
        export_menu.add_command(label="Export to Text File...", command=self.export_to_text)
        export_menu.add_separator()
        export_menu.add_command(label="Export as File Structure...", command=self.export_file_structure)
        
        # Settings menu for studio and author configuration
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Configure Studio & Author...", command=self.configure_studio_author)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About Solar", command=self.show_about)
    
    def configure_studio_author(self):
        """Let the user configure their default studio name and author"""
        # Create a simple dialog for this
        dialog = tk.Toplevel(self.root)
        dialog.title("Configure Studio & Author")
        dialog.geometry("400x200")
        dialog.configure(bg='#f0f0f0')
        dialog.resizable(False, False)
        
        # Center the dialog on the parent window
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Studio name
        tk.Label(dialog, text="Default Studio Name:", **self.label_style).grid(
            row=0, column=0, sticky='w', padx=10, pady=(10, 5))
        studio_var = tk.StringVar(value=self.settings.get('default_studio', 'Your Game Studio'))
        studio_entry = tk.Entry(dialog, textvariable=studio_var, width=40, **self.entry_style)
        studio_entry.grid(row=1, column=0, padx=10, pady=(0, 10), sticky='ew')
        
        # Author name
        tk.Label(dialog, text="Default Author Name:", **self.label_style).grid(
            row=2, column=0, sticky='w', padx=10, pady=(0, 5))
        author_var = tk.StringVar(value=self.settings.get('default_author', 'Game Designer'))
        author_entry = tk.Entry(dialog, textvariable=author_var, width=40, **self.entry_style)
        author_entry.grid(row=3, column=0, padx=10, pady=(0, 20), sticky='ew')
        
        # Buttons
        button_frame = tk.Frame(dialog, bg='#f0f0f0')
        button_frame.grid(row=4, column=0, pady=10)
        
        def save_and_close():
            self.settings.set('default_studio', studio_var.get())
            self.settings.set('default_author', author_var.get())
            # Update current project if it's using defaults
            if self.project_data['metadata']['studio'] == 'Your Game Studio':
                self.project_data['metadata']['studio'] = studio_var.get()
            if self.project_data['metadata']['author'] == 'Game Designer':
                self.project_data['metadata']['author'] = author_var.get()
            dialog.destroy()
        
        tk.Button(button_frame, text="Save", command=save_and_close, 
                 **self.button_style).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(button_frame, text="Cancel", command=dialog.destroy, 
                 **self.button_style).pack(side=tk.LEFT)
        
        dialog.columnconfigure(0, weight=1)
        studio_entry.focus()
    
    def create_project_info_frame(self, parent):
        """Show current project information at the top"""
        info_frame = tk.Frame(parent, bg='#f0f0f0', relief=tk.SUNKEN, bd=1)
        info_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Project name
        tk.Label(info_frame, text="Project:", **self.title_style).pack(side=tk.LEFT, padx=8, pady=5)
        self.project_label = tk.Label(info_frame, text="No project loaded", 
                                     bg='#f0f0f0', font=('Arial', 8), fg='#666666')
        self.project_label.pack(side=tk.LEFT, padx=(5, 20))
        
        # Project path
        tk.Label(info_frame, text="Folder:", **self.title_style).pack(side=tk.LEFT, padx=(0, 5))
        self.path_label = tk.Label(info_frame, text="No folder set", 
                                  bg='#f0f0f0', font=('Arial', 8), fg='#666666')
        self.path_label.pack(side=tk.LEFT, padx=5)
    
    def create_notebook(self, parent):
        """Create the tabbed interface using classic styling"""
        # Using a simple frame-based "notebook" instead of ttk.Notebook for that classic look
        notebook_frame = tk.Frame(parent, bg='#f0f0f0')
        notebook_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tab buttons frame
        self.tab_frame = tk.Frame(notebook_frame, bg='#f0f0f0')
        self.tab_frame.pack(fill=tk.X, pady=(0, 3))
        
        # Content area
        self.content_frame = tk.Frame(notebook_frame, bg='#f0f0f0', relief=tk.SUNKEN, bd=2)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Keep track of tab buttons and content frames
        self.tab_buttons = {}
        self.tab_contents = {}
        self.current_tab = None
    
    def add_tab(self, name, display_text):
        """Add a new tab to our classic tab interface"""
        # Create the tab button
        tab_button = tk.Button(self.tab_frame, text=display_text, 
                              command=lambda: self.switch_tab(name),
                              relief=tk.RAISED, bd=2, bg='#e0e0e0', 
                              font=('Arial', 8, 'bold'), padx=12, pady=4)
        tab_button.pack(side=tk.LEFT, padx=1)
        self.tab_buttons[name] = tab_button
        
        # Create the content frame for this tab
        content = tk.Frame(self.content_frame, bg='#f8f8f8')
        self.tab_contents[name] = content
        
        # If this is the first tab, make it active
        if self.current_tab is None:
            self.switch_tab(name)
        
        return content
    
    def switch_tab(self, tab_name):
        """Switch to a different tab"""
        # Hide current tab content
        if self.current_tab:
            self.tab_contents[self.current_tab].pack_forget()
            self.tab_buttons[self.current_tab].config(relief=tk.RAISED, bg='#e0e0e0')
        
        # Show new tab content
        self.tab_contents[tab_name].pack(fill=tk.BOTH, expand=True)
        self.tab_buttons[tab_name].config(relief=tk.SUNKEN, bg='#c0c0c0')
        self.current_tab = tab_name
    
    def create_tabs(self):
        """Set up all the different sections of the GDD"""
        # Introduction tab - basic game info
        self.tabs = {}
        self.tabs['introduction'] = self.create_introduction_tab()
        
        # Design Pillars tab
        self.tabs['design_pillars'] = self.create_design_pillars_tab()
        
        # All the dynamic content tabs
        self.tabs['mechanics'] = self.create_mechanics_tab()
        self.tabs['levels'] = self.create_levels_tab() 
        self.tabs['characters'] = self.create_characters_tab()
        self.tabs['items'] = self.create_items_tab()
        self.tabs['audio'] = self.create_audio_tab()
        self.tabs['art_style'] = self.create_art_style_tab()
        self.tabs['technical'] = self.create_technical_tab()
    
    def create_introduction_tab(self):
        """Create the introduction/overview tab - this one uses static fields"""
        frame = self.add_tab('introduction', 'Introduction')
        
        # Main scrollable area
        canvas = tk.Canvas(frame, bg='#f8f8f8')
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#f8f8f8')
        
        # Set up scrolling behavior
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # All the introduction fields
        intro_fields = [
            ("Project Name", "project_name", "entry"),
            ("Game Title", "game_title", "entry"),
            ("Genre", "genre", "entry"),
            ("Target Platform(s)", "platform", "entry"),
            ("Target Audience", "target_audience", "entry"),
            ("Game Overview", "game_overview", "text"),
            ("Unique Selling Points", "unique_selling_points", "text"),
            ("Inspiration/References", "inspiration", "text")
        ]
        
        # Build the fields
        self.introduction_widgets = {}
        for i, (label, key, widget_type) in enumerate(intro_fields):
            # Field label
            tk.Label(scrollable_frame, text=label, **self.title_style).grid(
                row=i*2, column=0, sticky='w', pady=(10, 5), padx=10)
            
            # Field input
            if widget_type == "entry":
                widget = tk.Entry(scrollable_frame, width=80, **self.entry_style)
                widget.grid(row=i*2+1, column=0, sticky='ew', padx=10, pady=(0, 5))
            elif widget_type == "text":
                widget = scrolledtext.ScrolledText(scrollable_frame, height=6, width=80, 
                                                 bg='white', relief=tk.SUNKEN, bd=1, wrap=tk.WORD)
                widget.grid(row=i*2+1, column=0, sticky='ew', padx=10, pady=(0, 5))
            
            self.introduction_widgets[key] = widget
            
            # Add callback for project name and game title to update project display
            if key in ['project_name', 'game_title'] and widget_type == "entry":
                widget.bind('<KeyRelease>', self.on_project_info_changed)
        
        scrollable_frame.columnconfigure(0, weight=1)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def on_project_info_changed(self, event=None):
        """Update project display when project name or game title changes"""
        # Get current values directly from the widgets for live updating
        project_name = ""
        game_title = ""
        
        if 'project_name' in self.introduction_widgets:
            project_name = self.introduction_widgets['project_name'].get().strip()
        
        if 'game_title' in self.introduction_widgets:
            game_title = self.introduction_widgets['game_title'].get().strip()
        
        # Determine what to display
        display_name = project_name
        if not display_name:
            display_name = game_title
        if not display_name:
            display_name = "Untitled Project"
        
        # Update the display immediately
        self.project_label.config(text=display_name)
    
    def create_design_pillars_tab(self):
        """Create the design pillars tab for core design principles"""
        frame = self.add_tab('design_pillars', 'Design Pillars')
        
        # Container for everything
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # Header with controls
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Design Pillars", **self.title_style).pack(side=tk.LEFT)
        
        # Control buttons on the right side
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add New Pillar", 
                 command=lambda: self.add_design_pillar_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('design_pillars'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('design_pillars'), **self.button_style).pack(side=tk.LEFT)
        
        # Scrollable area for all the design pillar entries
        self.design_pillars_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.design_pillars_scrollbar = tk.Scrollbar(main_container, orient="vertical", 
                                              command=self.design_pillars_canvas.yview)
        self.design_pillars_frame = tk.Frame(self.design_pillars_canvas, bg='#f8f8f8')
        
        # Connect the scrolling
        self.design_pillars_frame.bind(
            "<Configure>",
            lambda e: self.design_pillars_canvas.configure(scrollregion=self.design_pillars_canvas.bbox("all"))
        )
        
        self.design_pillars_canvas.create_window((0, 0), window=self.design_pillars_frame, anchor="nw")
        self.design_pillars_canvas.configure(yscrollcommand=self.design_pillars_scrollbar.set)
        
        # Pack everything
        self.design_pillars_canvas.pack(side="left", fill="both", expand=True)
        self.design_pillars_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def create_mechanics_tab(self):
        """Create the mechanics tab where you define all your game systems"""
        frame = self.add_tab('mechanics', 'Mechanics')
        
        # Container for everything
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # Header with controls
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Game Mechanics", **self.title_style).pack(side=tk.LEFT)
        
        # Control buttons on the right side
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add New Mechanic", 
                 command=lambda: self.add_mechanic_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('mechanics'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('mechanics'), **self.button_style).pack(side=tk.LEFT)
        
        # Scrollable area for all the mechanics entries
        self.mechanics_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.mechanics_scrollbar = tk.Scrollbar(main_container, orient="vertical", 
                                              command=self.mechanics_canvas.yview)
        self.mechanics_frame = tk.Frame(self.mechanics_canvas, bg='#f8f8f8')
        
        # Connect the scrolling
        self.mechanics_frame.bind(
            "<Configure>",
            lambda e: self.mechanics_canvas.configure(scrollregion=self.mechanics_canvas.bbox("all"))
        )
        
        self.mechanics_canvas.create_window((0, 0), window=self.mechanics_frame, anchor="nw")
        self.mechanics_canvas.configure(yscrollcommand=self.mechanics_scrollbar.set)
        
        # Pack everything
        self.mechanics_canvas.pack(side="left", fill="both", expand=True)
        self.mechanics_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def add_mechanic_entry(self, entry_data=None):
        """Add a new game mechanic entry"""
        # Define what fields a mechanic should have
        fields = [
            ("Mechanic Name", "name", "entry"),
            ("Description", "description", "text"),
            ("How It Works", "implementation", "text"),
            ("Impact on Gameplay", "impact", "text"),
            ("Balance Notes", "balance", "text")
        ]
        
        # Figure out a title for this entry
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Mechanic #{len(self.collapsible_entries['mechanics']) + 1}"
        
        # Create the collapsible entry
        entry = CollapsibleEntry(
            parent=self.mechanics_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('mechanics', e),
            fields=fields
        )
        
        # Keep track of it
        self.collapsible_entries['mechanics'].append(entry)
        
        # Update the scroll area
        self.mechanics_frame.update_idletasks()
        self.mechanics_canvas.configure(scrollregion=self.mechanics_canvas.bbox("all"))
    
    def add_design_pillar_entry(self, entry_data=None):
        """Add a new design pillar entry"""
        # Define what fields a design pillar should have
        fields = [
            ("Pillar Name", "name", "entry"),
            ("Core Principle", "principle", "text"),
            ("Why It Matters", "importance", "text"),
            ("Implementation Guidelines", "implementation", "text"),
            ("Examples in Game", "examples", "text")
        ]
        
        # Figure out a title for this entry
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Pillar #{len(self.collapsible_entries['design_pillars']) + 1}"
        
        # Create the collapsible entry
        entry = CollapsibleEntry(
            parent=self.design_pillars_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('design_pillars', e),
            fields=fields
        )
        
        # Keep track of it
        self.collapsible_entries['design_pillars'].append(entry)
        
        # Update the scroll area
        self.design_pillars_frame.update_idletasks()
        self.design_pillars_canvas.configure(scrollregion=self.design_pillars_canvas.bbox("all"))
    
    def delete_collapsible_entry(self, section, entry):
        """Remove an entry from the specified section"""
        if entry in self.collapsible_entries[section]:
            self.collapsible_entries[section].remove(entry)
        
        # Update the appropriate scroll area
        canvas_map = {
            'design_pillars': self.design_pillars_canvas,
            'mechanics': self.mechanics_canvas,
            'levels': self.levels_canvas,
            'characters': self.characters_canvas,
            'items': self.items_canvas,
            'audio': self.audio_canvas,
            'art_style': self.art_canvas,
            'technical': self.technical_canvas
        }
        
        frame_map = {
            'design_pillars': self.design_pillars_frame,
            'mechanics': self.mechanics_frame,
            'levels': self.levels_frame,
            'characters': self.characters_frame,
            'items': self.items_frame,
            'audio': self.audio_frame,
            'art_style': self.art_frame,
            'technical': self.technical_frame
        }
        
        if section in frame_map:
            frame_map[section].update_idletasks()
            canvas_map[section].configure(scrollregion=canvas_map[section].bbox("all"))
    
    def expand_all_entries(self, section):
        """Expand every entry in the given section - useful for editing"""
        for entry in self.collapsible_entries[section]:
            entry.expand()
    
    def collapse_all_entries(self, section):
        """Collapse everything in the section - good for getting an overview"""
        for entry in self.collapsible_entries[section]:
            entry.collapse()
    
    def create_levels_tab(self):
        """Create the levels tab for all your level designs"""
        frame = self.add_tab('levels', 'Levels')
        
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # Header with controls
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Level Design", **self.title_style).pack(side=tk.LEFT)
        
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add New Level", 
                 command=lambda: self.add_level_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('levels'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('levels'), **self.button_style).pack(side=tk.LEFT)
        
        # Scrollable list
        self.levels_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.levels_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.levels_canvas.yview)
        self.levels_frame = tk.Frame(self.levels_canvas, bg='#f8f8f8')
        
        self.levels_frame.bind(
            "<Configure>",
            lambda e: self.levels_canvas.configure(scrollregion=self.levels_canvas.bbox("all"))
        )
        
        self.levels_canvas.create_window((0, 0), window=self.levels_frame, anchor="nw")
        self.levels_canvas.configure(yscrollcommand=self.levels_scrollbar.set)
        
        self.levels_canvas.pack(side="left", fill="both", expand=True)
        self.levels_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def add_level_entry(self, entry_data=None):
        """Add a new level design entry"""
        fields = [
            ("Level Name", "name", "entry"),
            ("Level Type", "level_type", "entry"),
            ("Difficulty", "difficulty", "entry"),
            ("Estimated Play Time", "play_time", "entry"),
            ("Level Objectives", "objectives", "text"),
            ("Environment Description", "environment", "text"),
            ("Enemies and Challenges", "enemies", "text"),
            ("Rewards and Collectibles", "rewards", "text"),
            ("Design Notes", "notes", "text")
        ]
        
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Level #{len(self.collapsible_entries['levels']) + 1}"
        
        entry = CollapsibleEntry(
            parent=self.levels_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('levels', e),
            fields=fields
        )
        
        self.collapsible_entries['levels'].append(entry)
        
        self.levels_frame.update_idletasks()
        self.levels_canvas.configure(scrollregion=self.levels_canvas.bbox("all"))
    
    def create_characters_tab(self):
        """Character design section - stats, abilities, backstory, etc."""
        frame = self.add_tab('characters', 'Characters')
        
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Character Design", **self.title_style).pack(side=tk.LEFT)
        
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add New Character", 
                 command=lambda: self.add_character_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('characters'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('characters'), **self.button_style).pack(side=tk.LEFT)
        
        self.characters_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.characters_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.characters_canvas.yview)
        self.characters_frame = tk.Frame(self.characters_canvas, bg='#f8f8f8')
        
        self.characters_frame.bind(
            "<Configure>",
            lambda e: self.characters_canvas.configure(scrollregion=self.characters_canvas.bbox("all"))
        )
        
        self.characters_canvas.create_window((0, 0), window=self.characters_frame, anchor="nw")
        self.characters_canvas.configure(yscrollcommand=self.characters_scrollbar.set)
        
        self.characters_canvas.pack(side="left", fill="both", expand=True)
        self.characters_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def add_character_entry(self, entry_data=None):
        """Add a new character design entry"""
        fields = [
            ("Character Name", "name", "entry"),
            ("Role/Class", "role", "entry"),
            ("Health Points", "health", "entry"),
            ("Movement Speed", "speed", "entry"),
            ("Attack Damage", "damage", "entry"),
            ("Attack Speed", "attack_speed", "entry"),
            ("Special Abilities", "abilities", "text"),
            ("Background Story", "background", "text"),
            ("Visual Description", "visual", "text"),
            ("AI Behavior", "behavior", "text"),
            ("Strengths", "strengths", "text"),
            ("Weaknesses", "weaknesses", "text")
        ]
        
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Character #{len(self.collapsible_entries['characters']) + 1}"
        
        entry = CollapsibleEntry(
            parent=self.characters_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('characters', e),
            fields=fields
        )
        
        self.collapsible_entries['characters'].append(entry)
        
        self.characters_frame.update_idletasks()
        self.characters_canvas.configure(scrollregion=self.characters_canvas.bbox("all"))
    
    def create_items_tab(self):
        """Items and equipment design tab"""
        frame = self.add_tab('items', 'Items')
        
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Item Design", **self.title_style).pack(side=tk.LEFT)
        
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add New Item", 
                 command=lambda: self.add_item_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('items'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('items'), **self.button_style).pack(side=tk.LEFT)
        
        self.items_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.items_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.items_canvas.yview)
        self.items_frame = tk.Frame(self.items_canvas, bg='#f8f8f8')
        
        self.items_frame.bind(
            "<Configure>",
            lambda e: self.items_canvas.configure(scrollregion=self.items_canvas.bbox("all"))
        )
        
        self.items_canvas.create_window((0, 0), window=self.items_frame, anchor="nw")
        self.items_canvas.configure(yscrollcommand=self.items_scrollbar.set)
        
        self.items_canvas.pack(side="left", fill="both", expand=True)
        self.items_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def create_audio_tab(self):
        """Audio and sound design tab"""
        frame = self.add_tab('audio', 'Audio')
        
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Audio Design", **self.title_style).pack(side=tk.LEFT)
        
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add Audio Element", 
                 command=lambda: self.add_audio_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('audio'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('audio'), **self.button_style).pack(side=tk.LEFT)
        
        self.audio_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.audio_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.audio_canvas.yview)
        self.audio_frame = tk.Frame(self.audio_canvas, bg='#f8f8f8')
        
        self.audio_frame.bind(
            "<Configure>",
            lambda e: self.audio_canvas.configure(scrollregion=self.audio_canvas.bbox("all"))
        )
        
        self.audio_canvas.create_window((0, 0), window=self.audio_frame, anchor="nw")
        self.audio_canvas.configure(yscrollcommand=self.audio_scrollbar.set)
        
        self.audio_canvas.pack(side="left", fill="both", expand=True)
        self.audio_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def create_art_style_tab(self):
        """Art style and visual design tab"""
        frame = self.add_tab('art_style', 'Art Style')
        
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Art Style & Visuals", **self.title_style).pack(side=tk.LEFT)
        
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add Art Element", 
                 command=lambda: self.add_art_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('art_style'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('art_style'), **self.button_style).pack(side=tk.LEFT)
        
        self.art_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.art_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.art_canvas.yview)
        self.art_frame = tk.Frame(self.art_canvas, bg='#f8f8f8')
        
        self.art_frame.bind(
            "<Configure>",
            lambda e: self.art_canvas.configure(scrollregion=self.art_canvas.bbox("all"))
        )
        
        self.art_canvas.create_window((0, 0), window=self.art_frame, anchor="nw")
        self.art_canvas.configure(yscrollcommand=self.art_scrollbar.set)
        
        self.art_canvas.pack(side="left", fill="both", expand=True)
        self.art_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    def create_technical_tab(self):
        """Technical specifications tab"""
        frame = self.add_tab('technical', 'Technical')
        
        main_container = tk.Frame(frame, bg='#f8f8f8')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        header_frame = tk.Frame(main_container, bg='#f8f8f8')
        header_frame.pack(fill=tk.X, pady=(0, 8))
        
        tk.Label(header_frame, text="Technical Specifications", **self.title_style).pack(side=tk.LEFT)
        
        controls_frame = tk.Frame(header_frame, bg='#f8f8f8')
        controls_frame.pack(side=tk.RIGHT)
        
        tk.Button(controls_frame, text="Add Technical Element", 
                 command=lambda: self.add_technical_entry(), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Expand All", 
                 command=lambda: self.expand_all_entries('technical'), **self.button_style).pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(controls_frame, text="Collapse All", 
                 command=lambda: self.collapse_all_entries('technical'), **self.button_style).pack(side=tk.LEFT)
        
        self.technical_canvas = tk.Canvas(main_container, bg='#f8f8f8')
        self.technical_scrollbar = tk.Scrollbar(main_container, orient="vertical", command=self.technical_canvas.yview)
        self.technical_frame = tk.Frame(self.technical_canvas, bg='#f8f8f8')
        
        self.technical_frame.bind(
            "<Configure>",
            lambda e: self.technical_canvas.configure(scrollregion=self.technical_canvas.bbox("all"))
        )
        
        self.technical_canvas.create_window((0, 0), window=self.technical_frame, anchor="nw")
        self.technical_canvas.configure(yscrollcommand=self.technical_scrollbar.set)
        
        self.technical_canvas.pack(side="left", fill="both", expand=True)
        self.technical_scrollbar.pack(side="right", fill="y")
        
        return frame
    
    # Now implement the entry creation methods for each section
    def add_item_entry(self, entry_data=None):
        """Add a new item design entry"""
        fields = [
            ("Item Name", "name", "entry"),
            ("Item Type", "item_type", "entry"),
            ("Rarity", "rarity", "entry"),
            ("Damage/Effect Value", "damage", "entry"),
            ("Use Time", "use_time", "entry"),
            ("Cooldown", "cooldown", "entry"),
            ("Category", "category", "entry"),
            ("Description", "description", "text"),
            ("Function/Purpose", "purpose", "text"),
            ("How it Works", "mechanism", "text"),
            ("Visual Description", "visual", "text"),
            ("Acquisition Method", "acquisition", "text"),
            ("Balance Notes", "balance", "text")
        ]
        
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Item #{len(self.collapsible_entries['items']) + 1}"
        
        entry = CollapsibleEntry(
            parent=self.items_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('items', e),
            fields=fields
        )
        
        self.collapsible_entries['items'].append(entry)
        
        self.items_frame.update_idletasks()
        self.items_canvas.configure(scrollregion=self.items_canvas.bbox("all"))
    
    def add_audio_entry(self, entry_data=None):
        """Add a new audio design entry"""
        fields = [
            ("Audio Name", "name", "entry"),
            ("Type", "audio_type", "entry"),
            ("Context/Trigger", "context", "entry"),
            ("Duration", "duration", "entry"),
            ("Volume Level", "volume", "entry"),
            ("Description", "description", "text"),
            ("Mood/Feeling", "mood", "text"),
            ("Implementation Notes", "implementation", "text")
        ]
        
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Audio #{len(self.collapsible_entries['audio']) + 1}"
        
        entry = CollapsibleEntry(
            parent=self.audio_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('audio', e),
            fields=fields
        )
        
        self.collapsible_entries['audio'].append(entry)
        
        self.audio_frame.update_idletasks()
        self.audio_canvas.configure(scrollregion=self.audio_canvas.bbox("all"))
    
    def add_art_entry(self, entry_data=None):
        """Add a new art style entry"""
        fields = [
            ("Element Name", "name", "entry"),
            ("Art Style", "style", "entry"),
            ("Color Palette", "colors", "entry"),
            ("Resolution/Size", "resolution", "entry"),
            ("Description", "description", "text"),
            ("Mood/Atmosphere", "mood", "text"),
            ("Technical Requirements", "technical", "text"),
            ("Reference Images/Inspiration", "references", "text")
        ]
        
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Art Element #{len(self.collapsible_entries['art_style']) + 1}"
        
        entry = CollapsibleEntry(
            parent=self.art_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('art_style', e),
            fields=fields
        )
        
        self.collapsible_entries['art_style'].append(entry)
        
        self.art_frame.update_idletasks()
        self.art_canvas.configure(scrollregion=self.art_canvas.bbox("all"))
    
    def add_technical_entry(self, entry_data=None):
        """Add a new technical specification entry"""
        fields = [
            ("Component Name", "name", "entry"),
            ("Technology/Engine", "technology", "entry"),
            ("Performance Target", "performance", "entry"),
            ("Platform Requirements", "requirements", "entry"),
            ("Description", "description", "text"),
            ("Implementation Details", "implementation", "text"),
            ("Challenges/Risks", "challenges", "text"),
            ("Testing Requirements", "testing", "text")
        ]
        
        if entry_data and entry_data.get('name'):
            title = entry_data['name']
        else:
            title = f"Technical #{len(self.collapsible_entries['technical']) + 1}"
        
        entry = CollapsibleEntry(
            parent=self.technical_frame,
            title=title,
            entry_data=entry_data,
            delete_callback=lambda e: self.delete_collapsible_entry('technical', e),
            fields=fields
        )
        
        self.collapsible_entries['technical'].append(entry)
        
        self.technical_frame.update_idletasks()
        self.technical_canvas.configure(scrollregion=self.technical_canvas.bbox("all"))
    
    def get_footer_credits(self):
        """Get the standard footer credits for all exported files"""
        return ("\n" + "="*60 + "\n"
                "Generated by Solar - Open Source GDD Builder\n"
                "Created by Dead Orbit Studios\n"
                "Built by Mikey LaBrecque\n"
                "Available at: https://github.com/mikeybowman/Solar\n"
                "="*60)
    
    # File operations with updated branding
    def new_project(self):
        """Create a brand new project"""
        if messagebox.askyesno("New Project", 
                             "Create a new project?\n\nAny unsaved changes will be lost."):
            self.project_data = self.create_empty_project()
            self.current_project_path = None
            self.update_ui_with_data()
            self.update_project_display()
    
    def set_project_folder(self):
        """Let the user pick where to save their project files"""
        folder_path = filedialog.askdirectory(title="Select Project Folder")
        if folder_path:
            self.current_project_path = folder_path
            self.update_project_display()
    
    def save_project(self):
        """Save the current project to disk"""
        if not self.current_project_path:
            self.save_project_as()
            return
        
        # Collect all the data from the UI
        self.collect_all_data()
        
        # Figure out the filename
        project_name = self.project_data['metadata']['project_name'] or 'untitled'
        # Clean up the filename - remove characters that don't play nice with filesystems
        clean_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        clean_name = clean_name.replace(' ', '_')
        
        project_file = os.path.join(self.current_project_path, f"{clean_name}_gdd.json")
        
        try:
            with open(project_file, 'w', encoding='utf-8') as f:
                json.dump(self.project_data, f, indent=2)
            
            self.project_data['metadata']['last_modified'] = datetime.now().isoformat()
            messagebox.showinfo("Project Saved", f"Project saved successfully!\n\n{project_file}")
            
        except Exception as e:
            messagebox.showerror("Save Error", f"Couldn't save the project:\n\n{str(e)}")
    
    def save_project_as(self):
        """Save the project with a new name/location"""
        file_path = filedialog.asksaveasfilename(
            title="Save Project As",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            self.current_project_path = os.path.dirname(file_path)
            self.collect_all_data()
            
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.project_data, f, indent=2)
                
                self.project_data['metadata']['last_modified'] = datetime.now().isoformat()
                messagebox.showinfo("Project Saved", f"Project saved successfully!\n\n{file_path}")
                self.update_project_display()
                
            except Exception as e:
                messagebox.showerror("Save Error", f"Couldn't save the project:\n\n{str(e)}")
    
    def load_project(self):
        """Load an existing project from disk"""
        file_path = filedialog.askopenfilename(
            title="Load Project",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.project_data = json.load(f)
                
                self.current_project_path = os.path.dirname(file_path)
                self.update_ui_with_data()
                self.update_project_display()
                messagebox.showinfo("Project Loaded", "Project loaded successfully!")
                
            except Exception as e:
                messagebox.showerror("Load Error", f"Couldn't load the project:\n\n{str(e)}")
    
    def collect_all_data(self):
        """Gather all the data from the UI and put it in our data structure"""
        # Update metadata
        self.project_data['metadata']['last_modified'] = datetime.now().isoformat()
        
        # Update studio and author from current settings
        self.project_data['metadata']['studio'] = self.settings.get('default_studio', 'Your Game Studio')
        self.project_data['metadata']['author'] = self.settings.get('default_author', 'Game Designer')
        
        # Introduction data from the static fields
        for key, widget in self.introduction_widgets.items():
            if isinstance(widget, (tk.Entry, ttk.Entry)):
                self.project_data['introduction'][key] = widget.get()
            elif isinstance(widget, scrolledtext.ScrolledText):
                self.project_data['introduction'][key] = widget.get('1.0', tk.END).rstrip('\n')
        
        # Update metadata project name from introduction project_name field, fallback to game_title
        project_name = self.project_data['introduction'].get('project_name')
        if not project_name:
            project_name = self.project_data['introduction'].get('game_title')
        if project_name:
            self.project_data['metadata']['project_name'] = project_name
        
        # Dynamic sections data using the collapsible entries
        for section in self.collapsible_entries:
            self.project_data[section] = []
            for entry in self.collapsible_entries[section]:
                entry_data = entry.get_data()
                self.project_data[section].append(entry_data)
        
        # Update the project display
        self.update_project_display()
    
    def export_file_structure(self):
        """Export the GDD as an organized folder structure"""
        # Get the latest data
        self.collect_all_data()
        
        # Ask the user where they want to create the folder structure
        base_directory = filedialog.askdirectory(title="Choose Directory for GDD File Structure")
        if not base_directory:
            return
        
        try:
            # Create the main project folder
            project_name = self.project_data['introduction'].get('game_title', 'Untitled_Game')
            safe_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_name = safe_name.replace(' ', '_')
            
            project_folder = os.path.join(base_directory, f"{safe_name}_GDD")
            os.makedirs(project_folder, exist_ok=True)
            
            # Create the main README file
            readme_path = os.path.join(project_folder, "README.txt")
            with open(readme_path, 'w', encoding='utf-8') as f:
                f.write(f"Game Design Document: {project_name}\n")
                f.write("="*50 + "\n\n")
                f.write(f"Studio: {self.project_data['metadata']['studio']}\n")
                f.write(f"Author: {self.project_data['metadata']['author']}\n")
                f.write(f"Version: {self.project_data['metadata']['version']}\n")
                f.write(f"Created: {self.project_data['metadata']['created_date'][:10]}\n")
                f.write(f"Last Modified: {self.project_data['metadata']['last_modified'][:10]}\n\n")
                
                f.write("FOLDER STRUCTURE:\n")
                f.write("- 01_Introduction/: Game overview and basic information\n")
                f.write("- 02_Mechanics/: Game mechanics and systems\n")
                f.write("- 03_Levels/: Level design documents\n")
                f.write("- 04_Characters/: Character designs and specs\n")
                f.write("- 05_Items/: Item designs and balance info\n")
                f.write("- 06_Audio/: Audio design specifications\n")
                f.write("- 07_Art_Style/: Visual design documents\n")
                f.write("- 08_Technical/: Technical requirements and specs\n\n")
                f.write("Each folder contains individual .txt files for each entry.\n")
            
            # Create introduction folder and file
            intro_folder = os.path.join(project_folder, "01_Introduction")
            os.makedirs(intro_folder, exist_ok=True)
            
            intro_file = os.path.join(intro_folder, "game_overview.txt")
            with open(intro_file, 'w', encoding='utf-8') as f:
                intro = self.project_data['introduction']
                f.write(f"GAME TITLE: {intro.get('game_title', 'Untitled Game')}\n")
                f.write("="*40 + "\n\n")
                
                for field_name, field_key in [
                    ("Genre", "genre"),
                    ("Target Platform(s)", "platform"), 
                    ("Target Audience", "target_audience"),
                    ("Game Overview", "game_overview"),
                    ("Unique Selling Points", "unique_selling_points"),
                    ("Inspiration/References", "inspiration")
                ]:
                    if intro.get(field_key):
                        f.write(f"{field_name}:\n{intro[field_key]}\n\n")
            
            # Create folders for dynamic content sections
            section_info = {
                'mechanics': ('02_Mechanics', 'Game Mechanics'),
                'levels': ('03_Levels', 'Level Design'),
                'characters': ('04_Characters', 'Characters'),
                'items': ('05_Items', 'Items'),
                'audio': ('06_Audio', 'Audio'),
                'art_style': ('07_Art_Style', 'Art Style'),
                'technical': ('08_Technical', 'Technical')
            }
            
            for section_key, (folder_name, display_name) in section_info.items():
                if self.project_data.get(section_key):
                    section_folder = os.path.join(project_folder, folder_name)
                    os.makedirs(section_folder, exist_ok=True)
                    
                    for i, entry_data in enumerate(self.project_data[section_key], 1):
                        # Create a safe filename from the entry name
                        entry_name = entry_data.get('name', f'{display_name}_{i:03d}')
                        safe_filename = "".join(c for c in entry_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        safe_filename = safe_filename.replace(' ', '_')
                        if not safe_filename:
                            safe_filename = f"{display_name}_{i:03d}"
                        
                        entry_file = os.path.join(section_folder, f"{safe_filename}.txt")
                        
                        with open(entry_file, 'w', encoding='utf-8') as f:
                            f.write(f"{display_name.upper()}: {entry_data.get('name', f'Entry {i}')}\n")
                            f.write("="*50 + "\n\n")
                            
                            # Write all the fields for this entry
                            for key, value in entry_data.items():
                                if value and str(value).strip():
                                    # Make the field name look nice
                                    formatted_key = key.replace('_', ' ').title()
                                    f.write(f"{formatted_key}:\n{value}\n\n")
            
            messagebox.showinfo("Export Complete!", 
                              f"GDD file structure created successfully!\n\n"
                              f"Location: {project_folder}\n\n"
                              f"The folder contains organized .txt files with all your content.")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to create file structure:\n\n{str(e)}")
    
    def export_to_text(self):
        """Export the GDD to a simple text file"""
        self.collect_all_data()
        
        file_path = filedialog.asksaveasfilename(
            title="Export to Text File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    # Header
                    f.write("GAME DESIGN DOCUMENT\n")
                    f.write("="*60 + "\n\n")
                    
                    game_title = self.project_data['introduction'].get('game_title', 'Untitled Game')
                    f.write(f"{game_title}\n")
                    f.write("-" * len(game_title) + "\n\n")
                    
                    # Metadata
                    f.write(f"Studio: {self.project_data['metadata']['studio']}\n")
                    f.write(f"Author: {self.project_data['metadata']['author']}\n")
                    f.write(f"Version: {self.project_data['metadata']['version']}\n")
                    f.write(f"Created: {self.project_data['metadata']['created_date'][:10]}\n")
                    f.write(f"Last Modified: {self.project_data['metadata']['last_modified'][:10]}\n\n")
                    
                    # Table of contents
                    f.write("TABLE OF CONTENTS\n")
                    f.write("="*20 + "\n\n")
                    f.write("1. Introduction\n")
                    f.write("2. Design Pillars\n")
                    f.write("3. Game Mechanics\n")
                    f.write("4. Level Design\n")
                    f.write("5. Character Design\n")
                    f.write("6. Item Design\n")
                    f.write("7. Audio Design\n")
                    f.write("8. Art Style\n")
                    f.write("9. Technical Specifications\n\n")
                    
                    # Introduction section
                    f.write("1. INTRODUCTION\n")
                    f.write("="*15 + "\n\n")
                    intro = self.project_data['introduction']
                    
                    for field_name, field_key in [
                        ("Project Name", "project_name"),
                        ("Game Title", "game_title"),
                        ("Game Overview", "game_overview"),
                        ("Genre", "genre"),
                        ("Target Platform(s)", "platform"),
                        ("Target Audience", "target_audience"),
                        ("Unique Selling Points", "unique_selling_points"),
                        ("Inspiration/References", "inspiration"),
                        ("Development Timeline", "development_timeline")
                    ]:
                        if intro.get(field_key):
                            f.write(f"{field_name}:\n{intro[field_key]}\n\n")
                    
                    # Dynamic sections
                    section_numbers = {
                        'design_pillars': '2. DESIGN PILLARS',
                        'mechanics': '3. GAME MECHANICS',
                        'levels': '4. LEVEL DESIGN',
                        'characters': '5. CHARACTER DESIGN',
                        'items': '6. ITEM DESIGN',
                        'audio': '7. AUDIO DESIGN',
                        'art_style': '8. ART STYLE',
                        'technical': '9. TECHNICAL SPECIFICATIONS'
                    }
                    
                    for section_key, section_title in section_numbers.items():
                        if self.project_data.get(section_key):
                            f.write(f"{section_title}\n")
                            f.write("="*len(section_title) + "\n\n")
                            
                            for i, entry_data in enumerate(self.project_data[section_key], 1):
                                entry_name = entry_data.get('name', f'Entry {i}')
                                f.write(f"{i}. {entry_name}\n")
                                f.write("-"*len(f"{i}. {entry_name}") + "\n")
                                
                                for key, value in entry_data.items():
                                    if key != 'name' and value and str(value).strip():
                                        formatted_key = key.replace('_', ' ').title()
                                        f.write(f"{formatted_key}: {value}\n")
                                f.write("\n")
                
                    # Success message outside the with block to avoid I/O error  
                    messagebox.showinfo("Export Complete!", f"GDD exported successfully!\n\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export to text:\n\n{str(e)}")
    
    # File operations
    def new_project(self):
        """Create a new project"""
        if messagebox.askyesno("New Project", "Create a new project? Unsaved changes will be lost."):
            self.project_data = self.create_empty_project()
            self.current_project_path = None
            self.update_ui_with_data()
            self.update_project_display()
    
    def set_project_folder(self):
        """Set the project folder path"""
        folder_path = filedialog.askdirectory(title="Select Project Folder")
        if folder_path:
            self.current_project_path = folder_path
            self.update_project_display()
    
    def export_file_structure(self):
        """Export GDD as organized file structure"""
        # Collect current data
        self.collect_all_data()
        
        # Choose directory
        base_directory = filedialog.askdirectory(title="Choose Directory for GDD File Structure")
        if not base_directory:
            return
        
        try:
            # Create main project folder
            project_name = self.project_data['introduction'].get('game_title', 'Untitled_Game').replace(' ', '_')
            project_folder = os.path.join(base_directory, f"{project_name}_GDD")
            os.makedirs(project_folder, exist_ok=True)
            
            # Create README with project overview
            readme_path = os.path.join(project_folder, "README.txt")
            with open(readme_path, 'w', encoding='utf-8') as f:
                f.write(f"Game Design Document: {self.project_data['introduction'].get('game_title', 'Untitled Game')}\n")
                f.write("=" * 60 + "\n\n")
                f.write(f"Studio: {self.project_data['metadata']['studio']}\n")
                f.write(f"Author: {self.project_data['metadata']['author']}\n")
                f.write(f"Version: {self.project_data['metadata']['version']}\n")
                f.write(f"Created: {self.project_data['metadata']['created_date'][:10]}\n")
                f.write(f"Last Modified: {self.project_data['metadata']['last_modified'][:10]}\n\n")
                f.write("FOLDER STRUCTURE:\n")
                f.write("- 01_Introduction/: Game overview and basic information\n")
                f.write("- 02_Design_Pillars/: Core design principles and guidelines\n")
                f.write("- 03_Mechanics/: Game mechanics and systems\n")
                f.write("- 04_Levels/: Level design documents\n")
                f.write("- 05_Characters/: Character designs and specifications\n")
                f.write("- 06_Items/: Item designs and balance information\n")
                f.write("- 07_Audio/: Audio design specifications\n")
                f.write("- 08_Art_Style/: Visual design documents\n")
                f.write("- 09_Technical/: Technical requirements and specifications\n\n")
                f.write("Each folder contains individual .txt files for each entry.\n")
            
            # Create Introduction folder and file
            intro_folder = os.path.join(project_folder, "01_Introduction")
            os.makedirs(intro_folder, exist_ok=True)
            
            intro_file = os.path.join(intro_folder, "game_overview.txt")
            with open(intro_file, 'w', encoding='utf-8') as f:
                intro = self.project_data['introduction']
                f.write(f"GAME TITLE: {intro.get('game_title', '')}\n")
                f.write("=" * 40 + "\n\n")
                
                # Include all introduction fields
                intro_fields = [
                    ("Project Name", "project_name"),
                    ("Genre", "genre"),
                    ("Target Platform(s)", "platform"),
                    ("Target Audience", "target_audience"),
                    ("Game Overview", "game_overview"),
                    ("Unique Selling Points", "unique_selling_points"),
                    ("Inspiration/References", "inspiration"),
                    ("Development Timeline", "development_timeline")
                ]
                
                for field_name, field_key in intro_fields:
                    if intro.get(field_key):
                        f.write(f"{field_name}:\n{intro[field_key]}\n\n")
            
            # Create dynamic section folders and files
            section_info = {
                'design_pillars': ('02_Design_Pillars', 'Design_Pillars'),
                'mechanics': ('03_Mechanics', 'Mechanics'),
                'levels': ('04_Levels', 'Levels'),
                'characters': ('05_Characters', 'Characters'),
                'items': ('06_Items', 'Items'),
                'audio': ('07_Audio', 'Audio'),
                'art_style': ('08_Art_Style', 'Art_Style'),
                'technical': ('09_Technical', 'Technical')
            }
            
            for section_key, (folder_name, display_name) in section_info.items():
                if self.project_data.get(section_key):
                    section_folder = os.path.join(project_folder, folder_name)
                    os.makedirs(section_folder, exist_ok=True)
                    
                    for i, entry_data in enumerate(self.project_data[section_key], 1):
                        # Create filename from entry name or use index
                        entry_name = entry_data.get('name', f'{display_name}_{i:03d}')
                        # Clean filename
                        safe_filename = "".join(c for c in entry_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        safe_filename = safe_filename.replace(' ', '_')
                        
                        entry_file = os.path.join(section_folder, f"{safe_filename}.txt")
                        
                        with open(entry_file, 'w', encoding='utf-8') as f:
                            f.write(f"{display_name.upper()}: {entry_data.get('name', f'Entry {i}')}\n")
                            f.write("=" * 50 + "\n\n")
                            
                            # Write all fields for this entry
                            for key, value in entry_data.items():
                                if value and value.strip():
                                    formatted_key = key.replace('_', ' ').title()
                                    f.write(f"{formatted_key}:\n{value}\n\n")
            
            messagebox.showinfo("Export Successful", 
                              f"GDD file structure created at:\n{project_folder}\n\n"
                              f"Contains organized folders with individual .txt files for each entry.")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to create file structure: {str(e)}")
    
    
    def update_ui_with_data(self):
        """Update all UI widgets with loaded data"""
        # Introduction data
        for key, widget in self.introduction_widgets.items():
            value = self.project_data['introduction'].get(key, '')
            if isinstance(widget, (tk.Entry, ttk.Entry)):
                widget.delete(0, tk.END)
                widget.insert(0, value)
            elif isinstance(widget, scrolledtext.ScrolledText):
                widget.delete('1.0', tk.END)
                widget.insert('1.0', value)
        
        # Clear and rebuild dynamic sections
        self.clear_dynamic_sections()
        
        # Rebuild all dynamic sections using collapsible entries
        section_add_methods = {
            'design_pillars': self.add_design_pillar_entry,
            'mechanics': self.add_mechanic_entry,
            'levels': self.add_level_entry,
            'characters': self.add_character_entry,
            'items': self.add_item_entry,
            'audio': self.add_audio_entry,
            'art_style': self.add_art_entry,
            'technical': self.add_technical_entry
        }
        
        for section, add_method in section_add_methods.items():
            for entry_data in self.project_data.get(section, []):
                add_method(entry_data)
    
    def clear_dynamic_sections(self):
        """Clear all dynamic sections"""
        # Clear all collapsible entries
        for section in self.collapsible_entries:
            self.collapsible_entries[section].clear()
        
        # Clear all frames
        frame_map = {
            'design_pillars': self.design_pillars_frame,
            'mechanics': self.mechanics_frame,
            'levels': self.levels_frame,
            'characters': self.characters_frame,
            'items': self.items_frame,
            'audio': self.audio_frame,
            'art_style': self.art_frame,
            'technical': self.technical_frame
        }
        
        for section, frame in frame_map.items():
            for child in frame.winfo_children():
                child.destroy()
    
    def update_project_display(self):
        """Update the project info display at the top"""
        # Get project name from project_name field first, then game title, then fallback to metadata
        project_name = self.project_data['introduction'].get('project_name')
        if not project_name:
            project_name = self.project_data['introduction'].get('game_title')
        if not project_name:
            project_name = self.project_data['metadata'].get('project_name')
        if not project_name:
            project_name = "Untitled Project"
        
        self.project_label.config(text=project_name)
        
        if self.current_project_path:
            # Shorten the path if it's too long for the display
            path_text = self.current_project_path
            if len(path_text) > 50:
                path_text = "..." + path_text[-47:]
            self.path_label.config(text=path_text)
        else:
            self.path_label.config(text="No folder set")
    
    def export_to_word(self):
        """Export GDD to Word document"""
        if not DOCX_AVAILABLE:
            messagebox.showerror("Export Error", "python-docx is required for Word export.\nInstall with: pip install python-docx")
            return
        
        # Collect current data
        self.collect_all_data()
        
        file_path = filedialog.asksaveasfilename(
            title="Export to Word",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                doc = Document()
                
                # Title
                title = doc.add_heading(f"Game Design Document", 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                game_title = self.project_data['introduction'].get('game_title', 'Untitled Game')
                subtitle = doc.add_heading(game_title, 1)
                subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Metadata
                doc.add_paragraph(f"Version: {self.project_data['metadata']['version']}")
                doc.add_paragraph(f"Studio: {self.project_data['metadata']['studio']}")
                doc.add_paragraph(f"Author: {self.project_data['metadata']['author']}")
                doc.add_paragraph(f"Created: {self.project_data['metadata']['created_date'][:10]}")
                doc.add_paragraph(f"Last Modified: {self.project_data['metadata']['last_modified'][:10]}")
                
                doc.add_page_break()
                
                # Table of Contents
                doc.add_heading("Table of Contents", 1)
                toc_items = [
                    "1. Introduction",
                    "2. Design Pillars",
                    "3. Game Mechanics",
                    "4. Level Design",
                    "5. Character Design",
                    "6. Item Design",
                    "7. Audio Design",
                    "8. Art Style & Visuals",
                    "9. Technical Specifications"
                ]
                for item in toc_items:
                    doc.add_paragraph(item, style='List Number')
                
                doc.add_page_break()
                
                # Introduction
                doc.add_heading("1. Introduction", 1)
                intro = self.project_data['introduction']
                
                # Define all introduction fields with their display names
                intro_fields = [
                    ("Project Name", "project_name"),
                    ("Game Title", "game_title"),
                    ("Game Overview", "game_overview"),
                    ("Genre", "genre"),
                    ("Target Platform(s)", "platform"),
                    ("Target Audience", "target_audience"),
                    ("Unique Selling Points", "unique_selling_points"),
                    ("Inspiration/References", "inspiration"),
                    ("Development Timeline", "development_timeline")
                ]
                
                for field_name, field_key in intro_fields:
                    if intro.get(field_key):
                        doc.add_heading(field_name, 2)
                        doc.add_paragraph(intro[field_key])
                
                # Design Pillars
                if self.project_data.get('design_pillars'):
                    doc.add_heading("2. Design Pillars", 1)
                    
                    # Define all design pillar fields with their display names
                    pillar_fields = [
                        ("Core Principle", "principle"),
                        ("Why It Matters", "importance"),
                        ("Implementation Guidelines", "implementation"),
                        ("Examples in Game", "examples")
                    ]
                    
                    for i, pillar in enumerate(self.project_data['design_pillars'], 1):
                        if pillar.get('name'):
                            doc.add_heading(f"2.{i} {pillar['name']}", 2)
                            
                            for field_name, field_key in pillar_fields:
                                if pillar.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {pillar[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Mechanics
                if self.project_data.get('mechanics'):
                    doc.add_heading("3. Game Mechanics", 1)
                    
                    # Define all mechanic fields with their display names
                    mechanic_fields = [
                        ("Description", "description"),
                        ("How It Works", "implementation"),
                        ("Impact on Gameplay", "impact"),
                        ("Balance Notes", "balance")
                    ]
                    
                    for i, mechanic in enumerate(self.project_data['mechanics'], 1):
                        if mechanic.get('name'):
                            doc.add_heading(f"3.{i} {mechanic['name']}", 2)
                            
                            for field_name, field_key in mechanic_fields:
                                if mechanic.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {mechanic[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Levels
                if self.project_data.get('levels'):
                    doc.add_heading("4. Level Design", 1)
                    
                    # Define all level fields with their display names
                    level_fields = [
                        ("Level Type", "level_type"),
                        ("Difficulty", "difficulty"),
                        ("Estimated Play Time", "play_time"),
                        ("Level Objectives", "objectives"),
                        ("Environment Description", "environment"),
                        ("Enemies and Challenges", "enemies"),
                        ("Rewards and Collectibles", "rewards"),
                        ("Design Notes", "notes")
                    ]
                    
                    for i, level in enumerate(self.project_data['levels'], 1):
                        if level.get('name'):
                            doc.add_heading(f"4.{i} {level['name']}", 2)
                            
                            for field_name, field_key in level_fields:
                                if level.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {level[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Characters
                if self.project_data.get('characters'):
                    doc.add_heading("5. Character Design", 1)
                    
                    # Define all character fields with their display names
                    character_fields = [
                        ("Role/Class", "role"),
                        ("Health Points", "health"),
                        ("Movement Speed", "speed"),
                        ("Attack Damage", "damage"),
                        ("Attack Speed", "attack_speed"),
                        ("Special Abilities", "abilities"),
                        ("Background Story", "background"),
                        ("Visual Description", "visual"),
                        ("AI Behavior", "behavior"),
                        ("Strengths", "strengths"),
                        ("Weaknesses", "weaknesses")
                    ]
                    
                    for i, character in enumerate(self.project_data['characters'], 1):
                        if character.get('name'):
                            doc.add_heading(f"5.{i} {character['name']}", 2)
                            
                            for field_name, field_key in character_fields:
                                if character.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {character[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Items
                if self.project_data.get('items'):
                    doc.add_heading("6. Item Design", 1)
                    
                    # Define all item fields with their display names
                    item_fields = [
                        ("Item Type", "item_type"),
                        ("Rarity", "rarity"),
                        ("Damage/Effect Value", "damage"),
                        ("Use Time", "use_time"),
                        ("Cooldown", "cooldown"),
                        ("Category", "category"),
                        ("Description", "description"),
                        ("Function/Purpose", "purpose"),
                        ("How it Works", "mechanism"),
                        ("Visual Description", "visual"),
                        ("Acquisition Method", "acquisition"),
                        ("Balance Notes", "balance")
                    ]
                    
                    for i, item in enumerate(self.project_data['items'], 1):
                        if item.get('name'):
                            doc.add_heading(f"6.{i} {item['name']}", 2)
                            
                            for field_name, field_key in item_fields:
                                if item.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {item[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Audio
                if self.project_data.get('audio'):
                    doc.add_heading("7. Audio Design", 1)
                    
                    # Define all audio fields with their display names
                    audio_fields = [
                        ("Type", "audio_type"),
                        ("Context/Trigger", "context"),
                        ("Duration", "duration"),
                        ("Volume Level", "volume"),
                        ("Description", "description"),
                        ("Mood/Feeling", "mood"),
                        ("Implementation Notes", "implementation")
                    ]
                    
                    for i, audio in enumerate(self.project_data['audio'], 1):
                        if audio.get('name'):
                            doc.add_heading(f"7.{i} {audio['name']}", 2)
                            
                            for field_name, field_key in audio_fields:
                                if audio.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {audio[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Art Style
                if self.project_data.get('art_style'):
                    doc.add_heading("8. Art Style & Visuals", 1)
                    
                    # Define all art style fields with their display names
                    art_fields = [
                        ("Art Style", "style"),
                        ("Color Palette", "colors"),
                        ("Resolution/Size", "resolution"),
                        ("Description", "description"),
                        ("Mood/Atmosphere", "mood"),
                        ("Technical Requirements", "technical"),
                        ("Reference Images/Inspiration", "references")
                    ]
                    
                    for i, art in enumerate(self.project_data['art_style'], 1):
                        if art.get('name'):
                            doc.add_heading(f"8.{i} {art['name']}", 2)
                            
                            for field_name, field_key in art_fields:
                                if art.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {art[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                # Technical
                if self.project_data.get('technical'):
                    doc.add_heading("9. Technical Specifications", 1)
                    
                    # Define all technical fields with their display names
                    technical_fields = [
                        ("Technology/Engine", "technology"),
                        ("Performance Target", "performance"),
                        ("Platform Requirements", "requirements"),
                        ("Description", "description"),
                        ("Implementation Details", "implementation"),
                        ("Challenges/Risks", "challenges"),
                        ("Testing Requirements", "testing")
                    ]
                    
                    for i, tech in enumerate(self.project_data['technical'], 1):
                        if tech.get('name'):
                            doc.add_heading(f"9.{i} {tech['name']}", 2)
                            
                            for field_name, field_key in technical_fields:
                                if tech.get(field_key):
                                    doc.add_paragraph(f"{field_name}: {tech[field_key]}")
                            doc.add_paragraph("")  # Add spacing between entries
                
                doc.save(file_path)
                messagebox.showinfo("Export Successful", f"GDD exported to {file_path}")
                
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export to Word: {str(e)}")
    
    def show_about(self):
        """Show information about Solar"""
        about_text = '''Solar - Open Source GDD Builder

A free, open-source tool for creating professional 
game design documents.

Built with love for the indie game development community!

Originally created by: Mikey LaBrecque
Studio: Dead Orbit Studios
Version: v1.1.0 - Open Source Edition

Features:
• Classic early 2000s interface design
• Collapsible entry system for easy organization
• Export to Word, Text, and File Structure
• Configurable studio and author information
• Cross-platform compatibility (Windows, Mac, Linux)
• Complete source code available for modification

This software is open source - feel free to modify, 
improve, and share with other developers!

For more information and source code:
https://github.com/mikeybowman/Solar

Happy game development!'''
        
        messagebox.showinfo("About Solar", about_text)

def main():
    """Main entry point for the application"""
    # Create the main window
    root = tk.Tk()
    
    # Set up our application
    app = SolarGDDBuilder(root)
    
    # Start the GUI event loop
    root.mainloop()

if __name__ == "__main__":
    main()
