import tkinter as tk
from tkinter import ttk, messagebox
import os
import json
from datetime import datetime
from .main_window import MainWindow

class HomeWindow(tk.Tk):
    """Home page for VMP Tool - manages projects and provides navigation."""
    
    def __init__(self):
        super().__init__()
        self.title("VMP Tool - Home")
        self.geometry("800x600")
        
        # Create projects directory if it doesn't exist
        self.projects_dir = os.path.join(os.getcwd(), "VMP-Projects")
        if not os.path.exists(self.projects_dir):
            os.makedirs(self.projects_dir)
        
        self.setup_ui()
        self.refresh_project_list()
    
    def setup_ui(self):
        """Create the home page UI."""
        # Header
        header_frame = tk.Frame(self, bg='#2c3e50', height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="Visual Manufacturing Procedures", 
                              font=("Arial", 24, "bold"), bg='#2c3e50', fg='white')
        title_label.pack(expand=True)
        
        # Main content area
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Buttons frame
        buttons_frame = tk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Create New VMP button
        new_vmp_btn = tk.Button(buttons_frame, text="Create New VMP", 
                               command=self.create_new_vmp,
                               font=("Arial", 12, "bold"),
                               bg='#3498db', fg='white', 
                               padx=20, pady=10)
        new_vmp_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Refresh button
        refresh_btn = tk.Button(buttons_frame, text="Refresh", 
                               command=self.refresh_project_list,
                               font=("Arial", 10),
                               bg='#95a5a6', fg='white',
                               padx=15, pady=8)
        refresh_btn.pack(side=tk.LEFT)
        
        # Projects list frame
        list_frame = tk.LabelFrame(main_frame, text="Your VMPs", font=("Arial", 12, "bold"))
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollable list
        self.create_project_list(list_frame)
    
    def create_project_list(self, parent):
        """Create a scrollable list of projects."""
        # Create frame for the scrollable list
        list_container = tk.Frame(parent)
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas for scrolling
        self.canvas = tk.Canvas(list_container, yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.canvas.yview)
        
        # Frame inside canvas for project items
        self.projects_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.projects_frame, anchor="nw")
        
        # Bind canvas resize
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        self.projects_frame.bind('<Configure>', self.on_frame_configure)
    
    def on_canvas_configure(self, event):
        """Handle canvas resize."""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def on_frame_configure(self, event):
        """Handle frame resize to update scroll region."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def refresh_project_list(self):
        """Refresh the list of projects."""
        # Clear existing items
        for widget in self.projects_frame.winfo_children():
            widget.destroy()
        
        # Get all project files
        project_files = []
        if os.path.exists(self.projects_dir):
            for file in os.listdir(self.projects_dir):
                if file.endswith('.json'):
                    project_files.append(file)
        
        project_files.sort(reverse=True)  # Most recent first
        
        if not project_files:
            # Show empty state
            empty_label = tk.Label(self.projects_frame, 
                                  text="No VMPs created yet.\nClick 'Create New VMP' to get started!",
                                  font=("Arial", 12), fg='gray')
            empty_label.pack(pady=50)
        else:
            # Show project items
            for i, filename in enumerate(project_files):
                self.create_project_item(filename, i)
    
    def create_project_item(self, filename, index):
        """Create a single project item in the list."""
        # Load project metadata
        project_path = os.path.join(self.projects_dir, filename)
        try:
            with open(project_path, 'r') as f:
                project_data = json.load(f)
        except:
            return  # Skip corrupted files
        
        # Project item frame
        item_frame = tk.Frame(self.projects_frame, relief=tk.RAISED, borderwidth=1)
        item_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Project info
        info_frame = tk.Frame(item_frame)
        info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Project name
        name = project_data.get('name', filename.replace('.json', ''))
        name_label = tk.Label(info_frame, text=name, font=("Arial", 14, "bold"))
        name_label.pack(anchor=tk.W)
        
        # Creation date
        created = project_data.get('created', 'Unknown')
        date_label = tk.Label(info_frame, text=f"Created: {created}", font=("Arial", 10), fg='gray')
        date_label.pack(anchor=tk.W)
        
        # Page count
        pages = project_data.get('pages', [])
        pages_label = tk.Label(info_frame, text=f"Pages: {len(pages)}", font=("Arial", 10), fg='gray')
        pages_label.pack(anchor=tk.W)
        
        # Buttons frame
        buttons_frame = tk.Frame(item_frame)
        buttons_frame.pack(side=tk.RIGHT, padx=10, pady=10)
        
        # Open button
        open_btn = tk.Button(buttons_frame, text="Open", 
                            command=lambda f=filename: self.open_project(f),
                            bg='#2ecc71', fg='white', padx=15, pady=5)
        open_btn.pack(pady=2)
        
        # Export PDF button
        export_btn = tk.Button(buttons_frame, text="Export PDF", 
                              command=lambda f=filename: self.export_project_pdf(f),
                              bg='#e74c3c', fg='white', padx=15, pady=5)
        export_btn.pack(pady=2)
    
    def create_new_vmp(self):
        """Create a new VMP project."""
        # Hide home window and open main window
        self.withdraw()
        main_window = MainWindow(home_window=self)
        main_window.protocol("WM_DELETE_WINDOW", self.on_main_window_close)
    
    def open_project(self, filename):
        """Open an existing project."""
        project_path = os.path.join(self.projects_dir, filename)
        try:
            # Hide home window and open main window with project
            self.withdraw()
            main_window = MainWindow(home_window=self, project_file=project_path)
            main_window.protocol("WM_DELETE_WINDOW", self.on_main_window_close)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open project: {str(e)}")
    
    def export_project_pdf(self, filename):
        """Export a project directly to PDF."""
        project_path = os.path.join(self.projects_dir, filename)
        try:
            # Create a temporary main window to handle the export
            temp_window = MainWindow(project_file=project_path, export_only=True)
            temp_window.withdraw()  # Hide it
            temp_window.export_to_pdf()
            temp_window.destroy()
            messagebox.showinfo("Success", "PDF exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export PDF: {str(e)}")
    
    def on_main_window_close(self):
        """Handle main window closing - return to home."""
        self.deiconify()  # Show home window again
        self.refresh_project_list()  # Refresh in case new projects were saved
    
    def save_project(self, project_data, name=None):
        """Save a project to the projects directory."""
        if not name:
            name = f"VMP_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # Add metadata
        project_data['name'] = name
        project_data['created'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Save to file
        filename = f"{name}.json"
        filepath = os.path.join(self.projects_dir, filename)
        
        with open(filepath, 'w') as f:
            json.dump(project_data, f, indent=2)
        
        return filepath
