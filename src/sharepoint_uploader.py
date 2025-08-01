import os
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import requests
import json
import urllib.parse
import base64

class SharePointUploader:
    """Handles direct SharePoint upload using network access and sharing links."""
    
    def __init__(self):
        # Default SharePoint folder URL provided by user
        self.default_sharepoint_url = "https://bobrick.sharepoint.com/:f:/s/ManufacturingEngineering/ElsQrBO3x1BFid4ZQGzflwgBsltOVvplh7-wW4QjWMQ6Zw?e=WDzeie"
        self.site_url = "https://bobrick.sharepoint.com/sites/ManufacturingEngineering"
        self.folder_path = "Shared Documents"  # Default document library
        
    def test_network_access(self):
        """Test if we can access SharePoint via network without authentication."""
        try:
            # Try to access the SharePoint site to see if we have network access
            response = requests.get(self.site_url, timeout=5)
            return response.status_code in [200, 302, 401, 403]  # Any response means we can reach it
        except:
            return False
    
    def upload_via_network_path(self, file_path, custom_folder=None):
        """Try to upload via mapped network drive or UNC path."""
        try:
            # Common SharePoint network mappings
            possible_paths = [
                r"\\bobrick.sharepoint.com\sites\ManufacturingEngineering\Shared Documents",
                r"\\bobrick.sharepoint.com@SSL\sites\ManufacturingEngineering\Shared Documents",
                r"\\bobrick-my.sharepoint.com\sites\ManufacturingEngineering\Shared Documents",
            ]
            
            if custom_folder:
                possible_paths = [os.path.join(path, custom_folder) for path in possible_paths]
            
            filename = os.path.basename(file_path)
            
            for network_path in possible_paths:
                try:
                    if os.path.exists(network_path):
                        destination = os.path.join(network_path, filename)
                        import shutil
                        shutil.copy2(file_path, destination)
                        return True, f"File uploaded successfully to: {destination}"
                except Exception as e:
                    continue
            
            return False, "Could not access SharePoint via network path"
            
        except Exception as e:
            return False, f"Network upload failed: {str(e)}"
    
    def upload_via_browser_automation(self, file_path):
        """Open SharePoint in browser and provide instructions for manual upload."""
        try:
            import webbrowser
            webbrowser.open(self.default_sharepoint_url)
            
            filename = os.path.basename(file_path)
            instructions = f"""SharePoint folder opened in your browser.

To upload your file:
1. Click 'Upload' in the SharePoint toolbar
2. Select 'Files' from the dropdown
3. Choose the file: {filename}
4. Click 'Open' to upload

File location: {file_path}"""
            
            messagebox.showinfo("Manual Upload Instructions", instructions)
            return True, "SharePoint opened for manual upload"
            
        except Exception as e:
            return False, f"Could not open SharePoint: {str(e)}"
    
    def upload_file(self, file_path, custom_folder=None):
        """Upload a file to SharePoint using available methods."""
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File not found: {file_path}")
            return False
        
        filename = os.path.basename(file_path)
        
        # Method 1: Try network path upload
        success, message = self.upload_via_network_path(file_path, custom_folder)
        if success:
            messagebox.showinfo("Upload Successful", message)
            return True
        
        # Method 2: Open browser for manual upload
        success, message = self.upload_via_browser_automation(file_path)
        if success:
            return True
        
        # If all methods fail
        messagebox.showerror("Upload Failed", 
                           f"Could not upload file automatically.\n\n" +
                           f"Please manually upload the file:\n{file_path}\n\n" +
                           f"To SharePoint folder:\n{self.default_sharepoint_url}")
        return False

class SharePointUploadDialog:
    """Simplified dialog for SharePoint upload."""
    
    def __init__(self, parent, file_path):
        self.parent = parent
        self.file_path = file_path
        self.uploader = SharePointUploader()
        self.result = False
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Upload to SharePoint")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)  # Prevent resizing to maintain layout
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the simplified upload dialog UI."""
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Upload to SharePoint", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # File info
        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(file_frame, text="File to upload:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        tk.Label(file_frame, text=os.path.basename(self.file_path), font=("Arial", 10)).pack(anchor=tk.W)
        
        # Destination info
        dest_frame = tk.Frame(main_frame)
        dest_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(dest_frame, text="Destination:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        tk.Label(dest_frame, text="Manufacturing Engineering SharePoint", font=("Arial", 10)).pack(anchor=tk.W)
        
        # Optional subfolder
        folder_frame = tk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(folder_frame, text="Subfolder (optional):", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        self.folder_entry = tk.Entry(folder_frame)
        self.folder_entry.pack(fill=tk.X, pady=(5, 0))
        self.folder_entry.insert(0, "VMP-Files")  # Default subfolder
        
        # Instructions
        info_frame = tk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = "This will attempt to upload your file directly to the Manufacturing Engineering SharePoint folder. If automatic upload fails, SharePoint will open in your browser for manual upload."
        info_label = tk.Label(info_frame, text=info_text, font=("Arial", 9), 
                             wraplength=400, justify=tk.LEFT, fg='#666666')
        info_label.pack(anchor=tk.W)
        
        # Buttons - ensure they're always visible
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Add some spacing
        spacer = tk.Frame(button_frame)
        spacer.pack(side=tk.LEFT, expand=True)
        
        cancel_btn = tk.Button(button_frame, text="Cancel", command=self.cancel,
                              bg='#6c757d', fg='white', font=("Arial", 11, "bold"),
                              padx=25, pady=10, width=12)
        cancel_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.upload_btn = tk.Button(button_frame, text="Upload to SharePoint", command=self.upload_file,
                                   bg='#0078d4', fg='white', font=("Arial", 11, "bold"),
                                   padx=25, pady=10, width=18)
        self.upload_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
    def upload_file(self):
        """Handle file upload."""
        subfolder = self.folder_entry.get().strip()
        
        if self.uploader.upload_file(self.file_path, subfolder if subfolder else None):
            self.result = True
            self.dialog.destroy()
        
    def cancel(self):
        """Cancel the upload."""
        self.dialog.destroy()

def upload_to_sharepoint(parent, file_path):
    """Convenience function to upload a file to SharePoint."""
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File not found: {file_path}")
        return False
        
    dialog = SharePointUploadDialog(parent, file_path)
    parent.wait_window(dialog.dialog)
    return dialog.result
