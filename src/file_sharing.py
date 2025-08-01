import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser

class FileSharing:
    """Alternative file sharing options when SharePoint Graph API is blocked."""
    
    @staticmethod
    def copy_to_network_location(parent, file_path):
        """Copy file to a network location or shared drive."""
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File not found: {file_path}")
            return False
        
        # Ask user to select destination folder
        destination_folder = filedialog.askdirectory(
            title="Select Network Location or Shared Drive",
            initialdir="//",  # Start with network locations on Windows
        )
        
        if not destination_folder:
            return False
        
        try:
            filename = os.path.basename(file_path)
            destination_path = os.path.join(destination_folder, filename)
            
            # Copy the file
            shutil.copy2(file_path, destination_path)
            
            messagebox.showinfo("Success", 
                               f"File copied successfully to:\n{destination_path}\n\n" +
                               "You can now share this location with your team.")
            return True
            
        except Exception as e:
            messagebox.showerror("Copy Failed", f"Failed to copy file: {str(e)}")
            return False
    
    @staticmethod
    def open_sharepoint_in_browser(parent):
        """Open SharePoint in browser for manual upload."""
        dialog = SharePointBrowserDialog(parent)
        parent.wait_window(dialog.dialog)
        return dialog.result
    
    @staticmethod
    def create_email_with_attachment(parent, file_path):
        """Create an email with the VMP file as attachment."""
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File not found: {file_path}")
            return False
        
        try:
            # Create mailto URL with attachment (works with Outlook)
            filename = os.path.basename(file_path)
            subject = f"VMP File: {filename}"
            body = f"Please find the attached Visual Manufacturing Procedure file: {filename}"
            
            # For Windows, try to open with default email client
            import subprocess
            subprocess.run([
                'rundll32.exe', 'shell32.dll,ShellExec_RunDLL',
                'mailto:?subject=' + subject + '&body=' + body
            ])
            
            messagebox.showinfo("Email Created", 
                               f"An email has been created with subject '{subject}'.\n\n" +
                               f"Please manually attach the file:\n{file_path}")
            return True
            
        except Exception as e:
            messagebox.showwarning("Email Failed", 
                                 f"Could not create email automatically: {str(e)}\n\n" +
                                 f"Please manually create an email and attach:\n{file_path}")
            return False

class SharePointBrowserDialog:
    """Dialog to help user navigate to SharePoint for manual upload."""
    
    def __init__(self, parent):
        self.parent = parent
        self.result = False
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("SharePoint Manual Upload")
        self.dialog.geometry("500x350")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the dialog UI."""
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="SharePoint Manual Upload", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions
        instructions = tk.Text(main_frame, height=12, wrap=tk.WORD, 
                              font=("Arial", 10), bg="#f5f5f5")
        instructions.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        instructions_text = """Since automatic SharePoint upload is blocked by your organization's security policies, you can upload manually:

1. Click "Open SharePoint" below to open your organization's SharePoint site
2. Navigate to the document library where you want to store VMP files
3. Click "Upload" or drag and drop your VMP file
4. The file location will be copied to your clipboard for easy access

Alternative options:
• Use "Copy to Network Drive" to save to a shared network location
• Use "Email File" to send the VMP file via email
• Save to OneDrive and share the link with your team

Your VMP file is ready to share once you complete any of these steps."""
        
        instructions.insert("1.0", instructions_text)
        instructions.config(state=tk.DISABLED)
        
        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        sharepoint_btn = tk.Button(button_frame, text="Open SharePoint", 
                                  command=self.open_sharepoint,
                                  bg='#0078d4', fg='white', font=("Arial", 10, "bold"),
                                  padx=20, pady=5)
        sharepoint_btn.pack(side=tk.LEFT)
        
        close_btn = tk.Button(button_frame, text="Close", command=self.close_dialog,
                             bg='#6c757d', fg='white', font=("Arial", 10, "bold"),
                             padx=20, pady=5)
        close_btn.pack(side=tk.RIGHT)
    
    def open_sharepoint(self):
        """Open SharePoint in default browser."""
        try:
            # Open generic SharePoint URL - user will need to navigate to their org
            webbrowser.open("https://bobrick.sharepoint.com")
            self.result = True
            messagebox.showinfo("SharePoint Opened", 
                               "SharePoint has been opened in your browser.\n" +
                               "Navigate to your document library and upload your VMP file.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open SharePoint: {str(e)}")
    
    def close_dialog(self):
        """Close the dialog."""
        self.dialog.destroy()

class FileSharingDialog:
    """Main dialog for file sharing options."""
    
    def __init__(self, parent, file_path):
        self.parent = parent
        self.file_path = file_path
        self.result = False
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Share VMP File")
        self.dialog.geometry("450x300")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the sharing options dialog."""
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Share VMP File", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # File info
        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(file_frame, text="File to share:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        tk.Label(file_frame, text=os.path.basename(self.file_path), 
                font=("Arial", 10)).pack(anchor=tk.W)
        
        # Sharing options
        options_frame = tk.Frame(main_frame)
        options_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(options_frame, text="Choose sharing method:", 
                font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # SharePoint browser option
        sharepoint_btn = tk.Button(options_frame, text="📁 Open SharePoint in Browser", 
                                  command=self.open_sharepoint_browser,
                                  bg='#0078d4', fg='white', font=("Arial", 10, "bold"),
                                  padx=20, pady=8, width=30)
        sharepoint_btn.pack(pady=5, fill=tk.X)
        
        # Network location option
        network_btn = tk.Button(options_frame, text="💾 Copy to Network Drive", 
                               command=self.copy_to_network,
                               bg='#28a745', fg='white', font=("Arial", 10, "bold"),
                               padx=20, pady=8, width=30)
        network_btn.pack(pady=5, fill=tk.X)
        
        # Email option
        email_btn = tk.Button(options_frame, text="📧 Send via Email", 
                             command=self.send_via_email,
                             bg='#17a2b8', fg='white', font=("Arial", 10, "bold"),
                             padx=20, pady=8, width=30)
        email_btn.pack(pady=5, fill=tk.X)
        
        # Cancel button
        cancel_btn = tk.Button(options_frame, text="Cancel", command=self.cancel,
                              bg='#6c757d', fg='white', font=("Arial", 10, "bold"),
                              padx=20, pady=8, width=30)
        cancel_btn.pack(pady=(15, 0), fill=tk.X)
    
    def open_sharepoint_browser(self):
        """Open SharePoint in browser for manual upload."""
        if FileSharing.open_sharepoint_in_browser(self.parent):
            self.result = True
            self.dialog.destroy()
    
    def copy_to_network(self):
        """Copy file to network location."""
        if FileSharing.copy_to_network_location(self.parent, self.file_path):
            self.result = True
            self.dialog.destroy()
    
    def send_via_email(self):
        """Send file via email."""
        if FileSharing.create_email_with_attachment(self.parent, self.file_path):
            self.result = True
            self.dialog.destroy()
    
    def cancel(self):
        """Cancel sharing."""
        self.dialog.destroy()

def share_file(parent, file_path):
    """Main function to share a file using available methods."""
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File not found: {file_path}")
        return False
    
    dialog = FileSharingDialog(parent, file_path)
    parent.wait_window(dialog.dialog)
    return dialog.result
