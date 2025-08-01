import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import json
from .sharepoint_uploader import upload_to_sharepoint

class Page:
    def __init__(self, page_type='standard', title="", bullets=None, image_path1=None, image_path2=None, full_image_path=None, 
                 created_by="", date="", version="", approved_by="", approval_date="",
                 safety_warning=False, quality_check=False):
        self.page_type = page_type
        self.title = title
        self.bullets = bullets if bullets is not None else ["", "", ""]
        self.image_path1 = image_path1
        self.image_path2 = image_path2
        self.full_image_path = full_image_path
        # New title page fields
        self.created_by = created_by
        self.date = date
        self.version = version
        self.approved_by = approved_by
        self.approval_date = approval_date
        # New warning/check fields
        self.safety_warning = safety_warning
        self.quality_check = quality_check

    def to_dict(self):
        return {
            'page_type': self.page_type,
            'title': self.title,
            'bullets': self.bullets,
            'image_path1': self.image_path1,
            'image_path2': self.image_path2,
            'full_image_path': self.full_image_path,
            'created_by': self.created_by,
            'date': self.date,
            'version': self.version,
            'approved_by': self.approved_by,
            'approval_date': self.approval_date,
            'safety_warning': self.safety_warning,
            'quality_check': self.quality_check
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            page_type=data.get('page_type', 'standard'),
            title=data.get('title', ''),
            bullets=data.get('bullets', ["", "", ""]),
            image_path1=data.get('image_path1'),
            image_path2=data.get('image_path2'),
            full_image_path=data.get('full_image_path'),
            created_by=data.get('created_by', ''),
            date=data.get('date', ''),
            version=data.get('version', ''),
            approved_by=data.get('approved_by', ''),
            approval_date=data.get('approval_date', ''),
            safety_warning=data.get('safety_warning', False),
            quality_check=data.get('quality_check', False)
        )

class EditorPage(tk.Frame):
    """Editor page for VMP Tool - handles editing and page navigation."""

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.images_dir = os.path.join(os.getcwd(), "VMP-Images")
        self.project_file = None
        # Create the initial page and apply a workaround for initialization issues
        page = Page('title')
        if not hasattr(page, 'created_by'):
            page.created_by = ""
            page.date = ""
            page.version = ""
            page.approved_by = ""
            page.approval_date = ""
        self.pages = [page]
        self.current_page_index = 0
        self.selected_gallery_image_path = None

        # Initialize UI elements to None
        self.title_text = None
        self.created_by_entry = None
        self.date_entry = None
        self.version_entry = None
        self.approved_by_entry = None
        self.approval_date_entry = None

        self.safety_var = tk.BooleanVar()
        self.quality_var = tk.BooleanVar()

        self.setup_ui()
        self.show_page() # Show initial blank page

    def load_data(self, project_file=None):
        """Loads a project or resets to a new one."""
        if project_file and os.path.exists(project_file):
            self.load_project(project_file)
        else:
            self.project_file = None
            self.pages = [Page('title')]  # Start with a title page
            self.current_page_index = 0
            self.show_page()
        self.load_gallery_images()

    def setup_ui(self):
        """Set up the user interface."""
        main_paned_window = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, bg='#ecf0f1')
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        editor_pane = tk.Frame(main_paned_window, bg='#ecf0f1')
        main_paned_window.add(editor_pane, stretch="always")

        gallery_pane = tk.Frame(main_paned_window, bg='#bdc3c7', width=250)
        main_paned_window.add(gallery_pane, stretch="never")

        editor_pane.grid_rowconfigure(0, weight=1)
        editor_pane.grid_columnconfigure(0, weight=1)

        main_frame = tk.Frame(editor_pane, bg='#ecf0f1')
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        self.warnings_container = tk.Frame(main_frame, bg='#ecf0f1')
        self.warnings_container.pack(fill=tk.X, padx=10, pady=(10, 0))

        self.content_container = tk.Frame(main_frame, bg='#ecf0f1')
        self.content_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        nav_frame = tk.Frame(editor_pane, bg='#bdc3c7')
        nav_frame.grid(row=1, column=0, sticky='ew')
        nav_frame.grid_columnconfigure(1, weight=1)

        left_buttons = tk.Frame(nav_frame, bg='#bdc3c7')
        left_buttons.grid(row=0, column=0, padx=10, pady=5)

        self.home_btn = tk.Button(left_buttons, text="Home", command=lambda: self.controller.show_frame("HomePage"))
        self.home_btn.pack(side=tk.LEFT, padx=5)

        center_buttons = tk.Frame(nav_frame, bg='#bdc3c7')
        center_buttons.grid(row=0, column=1, pady=5)

        self.prev_btn = tk.Button(center_buttons, text="< Prev", command=self.prev_page)
        self.prev_btn.pack(side=tk.LEFT, padx=5)

        self.page_label = tk.Label(center_buttons, text="Page 1/1", bg='#bdc3c7')
        self.page_label.pack(side=tk.LEFT, padx=5)

        self.next_btn = tk.Button(center_buttons, text="Next >", command=self.next_page)
        self.next_btn.pack(side=tk.LEFT, padx=5)

        # Add checkboxes for warnings
        warnings_frame = tk.Frame(center_buttons, bg='#bdc3c7')
        warnings_frame.pack(side=tk.LEFT, padx=20)
        self.safety_check = tk.Checkbutton(warnings_frame, text="Safety Warning", variable=self.safety_var, bg='#bdc3c7', command=self.update_warnings)
        self.safety_check.pack(side=tk.LEFT)
        self.quality_check = tk.Checkbutton(warnings_frame, text="Quality Check", variable=self.quality_var, bg='#bdc3c7', command=self.update_warnings)
        self.quality_check.pack(side=tk.LEFT, padx=10)

        right_buttons = tk.Frame(nav_frame, bg='#bdc3c7')
        right_buttons.grid(row=0, column=2, padx=10, pady=5)

        # Add page type selection
        add_page_frame = tk.Frame(right_buttons, bg='#bdc3c7')
        add_page_frame.pack(side=tk.LEFT, padx=5)
        
        tk.Label(add_page_frame, text="Add:", bg='#bdc3c7', font=("Arial", 8)).pack(side=tk.TOP)
        
        from tkinter import ttk
        self.page_type_var = tk.StringVar(value="Standard Page")
        self.page_type_combo = ttk.Combobox(add_page_frame, textvariable=self.page_type_var, 
                                           values=["Title Page", "Standard Page", "Full Image Page"],
                                           state="readonly", width=12, font=("Arial", 8))
        self.page_type_combo.pack(side=tk.TOP, pady=(0, 2))
        
        self.add_page_btn = tk.Button(add_page_frame, text="Add", command=self.add_page, font=("Arial", 8))
        self.add_page_btn.pack(side=tk.TOP)

        self.delete_page_btn = tk.Button(right_buttons, text="Delete Page", command=self.delete_page)
        self.delete_page_btn.pack(side=tk.LEFT, padx=5)

        self.save_btn = tk.Button(right_buttons, text="Save Project", command=self.save_project)
        self.save_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = tk.Button(right_buttons, text="Export to PDF", command=self.export_to_pdf)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        self.sharepoint_btn = tk.Button(right_buttons, text="Upload to SharePoint", command=self.upload_to_sharepoint,
                                       bg='#0078d4', fg='white')
        self.sharepoint_btn.pack(side=tk.LEFT, padx=5)

        self.setup_image_gallery(gallery_pane)
        self.update_navigation()

    def setup_image_gallery(self, parent_frame):
        """Creates a scrollable image gallery."""
        gallery_label = tk.Label(parent_frame, text="Image Gallery", font=("Arial", 12, "bold"), bg='#bdc3c7')
        gallery_label.pack(pady=10)

        canvas = tk.Canvas(parent_frame, bg='#ecf0f1')
        scrollbar = tk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg='#ecf0f1')

        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))) 
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.load_gallery_images()

    def load_gallery_images(self):
        """Loads images from the VMP-Images directory into the gallery."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not os.path.exists(self.images_dir):
            os.makedirs(self.images_dir)

        image_files = [f for f in os.listdir(self.images_dir) if f.lower().endswith(('png', 'jpg', 'jpeg', 'gif'))]
        
        for image_name in image_files:
            image_path = os.path.join(self.images_dir, image_name)
            try:
                img = Image.open(image_path)
                img.thumbnail((150, 150))
                photo = ImageTk.PhotoImage(img)

                label = tk.Label(self.scrollable_frame, image=photo, bg='#ecf0f1', relief=tk.RAISED, borderwidth=2)
                label.image = photo
                label.pack(pady=5, padx=5)
                label.bind("<Button-1>", lambda e, p=image_path, l=label: self.select_gallery_image(p, l))
            except Exception as e:
                print(f"Error loading gallery image {image_name}: {e}")

    def select_gallery_image(self, image_path, clicked_label):
        """Highlights the selected image in the gallery."""
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, tk.Label):
                widget.config(bg='#ecf0f1', relief=tk.RAISED)
        
        clicked_label.config(bg='#3498db', relief=tk.SUNKEN)
        self.selected_gallery_image_path = image_path

    def assign_image_to_placeholder(self, image_index):
        """Assigns the selected gallery image to a placeholder."""
        if not self.selected_gallery_image_path:
            messagebox.showwarning("No Image Selected", "Please select an image from the gallery first.")
            return

        page = self.pages[self.current_page_index]
        if image_index == 0:
            page.image_path1 = self.selected_gallery_image_path
        elif image_index == 1:
            page.image_path2 = self.selected_gallery_image_path
        elif image_index == 'full':
            page.full_image_path = self.selected_gallery_image_path
        
        self.show_page()

    def show_page(self):
        """Displays the current page's content based on page type."""
        # Clear previous page's content and warnings first
        for widget in self.content_container.winfo_children():
            widget.destroy()
        for widget in self.warnings_container.winfo_children():
            widget.destroy()

        self.update_navigation()

        page = self.pages[self.current_page_index]
        
        # Reset UI references
        self.bullet_texts = []
        self.title_text = None
        self.full_image_label = None
        self.created_by_entry = None
        self.date_entry = None
        self.version_entry = None
        self.approved_by_entry = None
        self.approval_date_entry = None

        self.safety_var.set(page.safety_warning)
        self.quality_var.set(page.quality_check)
        self.show_warning_indicators(page)

        if page.page_type == 'title':
            self.show_title_page(page)
        elif page.page_type == 'standard':
            self.show_standard_page(page)
        elif page.page_type == 'full_image':
            self.show_full_image_page(page)
    
    def show_title_page(self, page):
        """Display title page layout with metadata fields."""
        container = tk.Frame(self.content_container, bg='#ecf0f1')
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        container.grid_columnconfigure(0, weight=1)

        # --- Title --- 
        title_label = tk.Label(container, text="Procedure Title", font=("Arial", 14, "bold"), bg='#ecf0f1')
        title_label.grid(row=0, column=0, sticky='w')
        self.title_text = tk.Text(container, height=3, wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1, font=("Arial", 18, "bold"))
        self.title_text.insert(tk.END, page.title)
        self.title_text.grid(row=1, column=0, sticky='ew', pady=(0, 20))

        # --- Metadata Frame ---
        metadata_frame = tk.Frame(container, bg='#ecf0f1')
        metadata_frame.grid(row=2, column=0, sticky='ew')
        metadata_frame.grid_columnconfigure(1, weight=1)
        metadata_frame.grid_columnconfigure(3, weight=1)

        fields = {
            "Created by:": (0, 0, 'created_by_entry', page.created_by),
            "Date:": (1, 0, 'date_entry', page.date),
            "Version:": (2, 0, 'version_entry', page.version),
            "Approved by:": (0, 2, 'approved_by_entry', page.approved_by),
            "Approval Date:": (1, 2, 'approval_date_entry', page.approval_date)
        }

        for label_text, (r, c, entry_attr, value) in fields.items():
            label = tk.Label(metadata_frame, text=label_text, font=("Arial", 10), bg='#ecf0f1')
            label.grid(row=r, column=c, sticky='w', padx=5, pady=5)
            entry = tk.Entry(metadata_frame, font=("Arial", 10))
            entry.insert(tk.END, value)
            entry.grid(row=r, column=c + 1, sticky='ew', padx=5, pady=5)
            setattr(self, entry_attr, entry)
    
    def show_standard_page(self, page):
        """Display standard page layout (3 bullets + 2 images)."""
        left_frame = tk.Frame(self.content_container, bg='#ecf0f1')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        right_frame = tk.Frame(self.content_container, bg='#ecf0f1')
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.bullet_texts = []
        for i in range(3):
            text_widget = tk.Text(left_frame, height=8, width=50, wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1)
            text_widget.insert(tk.END, page.bullets[i])
            text_widget.pack(pady=5, fill=tk.BOTH, expand=True)
            self.bullet_texts.append(text_widget)

        self.display_image(right_frame, page.image_path1, 0)
        self.display_image(right_frame, page.image_path2, 1)
    
    def show_full_image_page(self, page):
        """Display full image page layout."""
        image_frame = tk.Frame(self.content_container, bg='#ecf0f1')
        image_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.display_full_image(image_frame, page.full_image_path)

    def display_image(self, parent, image_path, index):
        """Displays an image in a given frame or a placeholder if no image."""
        container = tk.Frame(parent, bg='#ecf0f1')
        container.pack(fill=tk.BOTH, expand=True, pady=5)

        try:
            if image_path and os.path.exists(image_path):
                img = Image.open(image_path)
                img.thumbnail((400, 300))
                photo = ImageTk.PhotoImage(img)

                label = tk.Label(container, image=photo, bg="#ffffff")
                label.image = photo
            else:
                label = tk.Label(container, text="Click to assign image", bg="#cccccc", height=15, width=40, relief=tk.GROOVE)
        except Exception as e:
            label = tk.Label(container, text="Invalid Image", bg="#ffcccc", height=15, width=40)

        label.pack(fill=tk.BOTH, expand=True)
        label.bind("<Button-1>", lambda e, i=index: self.assign_image_to_placeholder(i))
    
    def display_full_image(self, parent, image_path):
        """Displays a full-page image or placeholder."""
        try:
            if image_path and os.path.exists(image_path):
                img = Image.open(image_path)
                # Scale to fit the container while maintaining aspect ratio
                img.thumbnail((800, 600))
                photo = ImageTk.PhotoImage(img)
                label = tk.Label(parent, image=photo, bg="#ffffff")
                label.image = photo
            else:
                label = tk.Label(parent, text="Click to assign full page image", bg="#cccccc", 
                                height=30, width=80, relief=tk.GROOVE, font=("Arial", 14))
        except Exception as e:
            label = tk.Label(parent, text="Invalid Image", bg="#ffcccc", 
                           height=30, width=80, font=("Arial", 14))

        label.pack(fill=tk.BOTH, expand=True)
        label.bind("<Button-1>", lambda e: self.assign_image_to_placeholder('full'))
        self.full_image_label = label

    def show_warning_indicators(self, page):
        """Displays colored indicators based on page flags."""
        if page.safety_warning:
            safety_indicator = tk.Label(self.warnings_container, text="SAFETY WARNING", bg="#f1c40f", fg="#000000", font=("Arial", 10, "bold"))
            safety_indicator.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        if page.quality_check:
            quality_indicator = tk.Label(self.warnings_container, text="QUALITY CHECK", bg="#3498db", fg="#ffffff", font=("Arial", 10, "bold"))
            quality_indicator.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

    def update_warnings(self):
        """Called when a warning checkbox is clicked."""
        self.save_current_page_data()
        self.show_page()

    def save_current_page_data(self):
        """Saves the current page's data based on page type."""
        if not hasattr(self, 'pages') or not self.pages:
            return
            
        page = self.pages[self.current_page_index]
        
        # Save title page data
        if page.page_type == 'title':
            if hasattr(self, 'title_text') and self.title_text:
                page.title = self.title_text.get("1.0", tk.END).strip()
            if hasattr(self, 'created_by_entry') and self.created_by_entry:
                page.created_by = self.created_by_entry.get().strip()
            if hasattr(self, 'date_entry') and self.date_entry:
                page.date = self.date_entry.get().strip()
            if hasattr(self, 'version_entry') and self.version_entry:
                page.version = self.version_entry.get().strip()
            if hasattr(self, 'approved_by_entry') and self.approved_by_entry:
                page.approved_by = self.approved_by_entry.get().strip()
            if hasattr(self, 'approval_date_entry') and self.approval_date_entry:
                page.approval_date = self.approval_date_entry.get().strip()
        
        # Save warning/check data for all page types
        page.safety_warning = self.safety_var.get()
        page.quality_check = self.quality_var.get()
        
        # Save standard page data
        if hasattr(self, 'bullet_texts') and self.bullet_texts:
            for i in range(3):
                page.bullets[i] = self.bullet_texts[i].get("1.0", tk.END).strip()
    


    def update_navigation(self):
        """Updates the navigation buttons and page label."""
        self.page_label.config(text=f"Page {self.current_page_index + 1}/{len(self.pages)}")
        self.prev_btn.config(state=tk.NORMAL if self.current_page_index > 0 else tk.DISABLED)
        self.next_btn.config(state=tk.NORMAL if self.current_page_index < len(self.pages) - 1 else tk.DISABLED)
        self.delete_page_btn.config(state=tk.NORMAL if len(self.pages) > 1 else tk.DISABLED)

    def next_page(self):
        if self.current_page_index < len(self.pages) - 1:
            self.save_current_page_data() # Save current page before moving
            self.current_page_index += 1
            self.show_page() # Then show the new page

    def prev_page(self):
        if self.current_page_index > 0:
            self.save_current_page_data() # Save current page before moving
            self.current_page_index -= 1
            self.show_page() # Then show the new page

    def add_page(self):
        self.save_current_page_data() # Save current page before adding a new one
        
        # Determine page type from dropdown
        page_type_map = {
            "Title Page": "title",
            "Standard Page": "standard", 
            "Full Image Page": "full_image"
        }
        selected_type = page_type_map.get(self.page_type_var.get(), "standard")
        
        self.pages.insert(self.current_page_index + 1, Page(selected_type))
        self.current_page_index += 1
        self.show_page()

    def delete_page(self):
        if len(self.pages) > 1:
            if messagebox.askyesno("Delete Page", "Are you sure you want to delete this page?"):
                del self.pages[self.current_page_index]
                if self.current_page_index >= len(self.pages):
                    self.current_page_index = len(self.pages) - 1
                self.show_page()

    def save_project(self):
        """Saves the entire project to a .vmp file."""
        self.save_current_page_data() # Ensure current page data is saved before serializing
        if not self.project_file:
            self.project_file = filedialog.asksaveasfilename(defaultextension=".vmp", filetypes=[("VMP Files", "*.vmp")], initialdir=os.path.join(os.getcwd(), "VMP-Projects"), title="Save Project As")
        if self.project_file:
            try:
                data = {'pages': [page.to_dict() for page in self.pages]}
                with open(self.project_file, 'w') as f:
                    json.dump(data, f, indent=4)
                messagebox.showinfo("Success", "Project saved successfully.")
                self.controller.frames["HomePage"].refresh_project_list()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save project: {e}")

    def load_project(self, project_file):
        """Loads a project from a .vmp file."""
        try:
            with open(project_file, 'r') as f:
                data = json.load(f)
            self.pages = []
            for page_data in data['pages']:
                # Defensive loading: ensure new fields exist, providing defaults if not.
                page_data.setdefault('created_by', '')
                page_data.setdefault('date', '')
                page_data.setdefault('version', '')
                page_data.setdefault('approved_by', '')
                page_data.setdefault('approval_date', '')
                self.pages.append(Page.from_dict(page_data))
            self.project_file = project_file
            self.current_page_index = 0
            self.show_page()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project: {e}")

    def export_warning_indicators_to_pdf(self, pdf, page):
        """Draws colored warning indicators at the top of a PDF page."""
        initial_y = pdf.get_y()
        indicator_height = 8
        
        if page.safety_warning:
            pdf.set_fill_color(241, 196, 15) # Yellow
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, indicator_height, "SAFETY WARNING", 1, 1, 'C', fill=True)
            pdf.ln(2)
        
        if page.quality_check:
            pdf.set_fill_color(52, 152, 219) # Blue
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, indicator_height, "QUALITY CHECK", 1, 1, 'C', fill=True)
            pdf.ln(2)

        # Reset colors and font
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(0, 0, 0)
        # pdf.set_y(initial_y + (12 * (int(page.safety_warning) + int(page.quality_check))))

    def export_to_pdf(self):
        """Exports the current project to a PDF file."""
        self.save_current_page_data()
        if not self.project_file:
            messagebox.showwarning("Save Required", "Please save the project before exporting.")
            self.save_project()
            if not self.project_file: # User cancelled save
                return

        pdf_filename = os.path.splitext(os.path.basename(self.project_file))[0] + '.pdf'
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Documents", "*.pdf")], initialfile=pdf_filename, title="Export to PDF")
        if not save_path:
            return

        from fpdf import FPDF
        pdf = FPDF(orientation='L', unit='mm', format='A4')

        pdf.set_auto_page_break(auto=True, margin=15)

        for i, page in enumerate(self.pages):
            pdf.add_page()
            self.export_warning_indicators_to_pdf(pdf, page)
            
            if page.page_type == 'title':
                self.export_title_page(pdf, page, i + 1)
            elif page.page_type == 'standard':
                self.export_standard_page(pdf, page, i + 1)
            elif page.page_type == 'full_image':
                self.export_full_image_page(pdf, page, i + 1)

        try:
            pdf.output(save_path)
            messagebox.showinfo("Success", f"PDF exported to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")
    
    def export_title_page(self, pdf, page, page_num):
        """Export a title page with metadata to PDF."""
        # --- Title ---
        pdf.set_font("Arial", 'B', 28)
        pdf.set_y(80)  # Position title to make space for metadata
        if page.title.strip():
            pdf.multi_cell(0, 15, page.title.strip(), 0, 'C')
        else:
            pdf.multi_cell(0, 15, "Untitled Procedure", 0, 'C')
        pdf.ln(20)

        # --- Metadata Table ---
        pdf.set_font("Arial", '', 12)
        
        # Define column widths and positions
        col_width = (pdf.w - 2 * 15) / 2  # Two columns
        left_col_x = pdf.l_margin
        right_col_x = left_col_x + col_width

        # Store initial Y position to align columns
        initial_y = pdf.get_y()
        
        # --- Left Column ---
        pdf.set_x(left_col_x)
        pdf.cell(30, 10, "Created by:", 0, 0)
        pdf.cell(col_width - 30, 10, page.created_by, 0, 1)
        
        pdf.set_x(left_col_x)
        pdf.cell(30, 10, "Date:", 0, 0)
        pdf.cell(col_width - 30, 10, page.date, 0, 1)

        pdf.set_x(left_col_x)
        pdf.cell(30, 10, "Version:", 0, 0)
        pdf.cell(col_width - 30, 10, page.version, 0, 1)

        # --- Right Column ---
        pdf.set_y(initial_y)  # Reset Y to align with the top of the left column
        
        pdf.set_x(right_col_x)
        pdf.cell(35, 10, "Approved by:", 0, 0)
        pdf.cell(col_width - 35, 10, page.approved_by, 0, 1)

        pdf.set_x(right_col_x)
        pdf.cell(35, 10, "Approval Date:", 0, 0)
        pdf.cell(col_width - 35, 10, page.approval_date, 0, 1)
    
    def export_standard_page(self, pdf, page, page_num):
        """Export a standard page (3 bullets + 2 images) to PDF."""
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"Page {page_num}", 0, 1, 'C')
        pdf.ln(5)

        col_width = (pdf.w - 2 * 15 - 10) / 2
        text_col_x = 15
        img_col_x = 15 + col_width + 10
        
        current_y = pdf.get_y()

        # Calculate vertical positions for the three bullet points
        content_h = pdf.h - current_y - 15 # 15 is bottom margin
        bullet_y_positions = [current_y, current_y + (content_h / 3), current_y + 2 * (content_h / 3)]

        pdf.set_font("Arial", '', 12)
        max_y_after_text = current_y
        for i, bullet in enumerate(page.bullets):
            if bullet.strip():
                pdf.set_y(bullet_y_positions[i])
                pdf.set_x(text_col_x)
                pdf.multi_cell(col_width, 10, f'* {bullet.strip()}')
                max_y_after_text = max(max_y_after_text, pdf.get_y())

        pdf.set_y(current_y)

        img_paths = [page.image_path1, page.image_path2]
        img_y_positions = [current_y, current_y + ((pdf.h - 2 * 15 - 20) / 2) + 5]

        for idx, img_path in enumerate(img_paths):
            if img_path and os.path.exists(img_path):
                try:
                    img = Image.open(img_path)
                    w, h = img.size
                    aspect_ratio = w / h
                    display_w = col_width
                    display_h = display_w / aspect_ratio
                    max_h = (pdf.h - 2 * 15 - 20) / 2 - 5
                    if display_h > max_h:
                        display_h = max_h
                        display_w = display_h * aspect_ratio
                    
                    pdf.image(img_path, x=img_col_x + (col_width - display_w) / 2, y=img_y_positions[idx], w=display_w, h=display_h)
                except Exception as e:
                    print(f"Could not add image {img_path} to PDF. Error: {e}")
        
        if max_y_after_text > pdf.get_y():
            pdf.set_y(max_y_after_text)
    
    def export_full_image_page(self, pdf, page, page_num):
        """Export a full image page to PDF."""
        if page.full_image_path and os.path.exists(page.full_image_path):
            try:
                img = Image.open(page.full_image_path)
                w, h = img.size
                aspect_ratio = w / h
                
                # Calculate dimensions to fit the page with margins
                page_w = pdf.w - 30  # 15mm margin on each side
                page_h = pdf.h - 30  # 15mm margin on top and bottom
                
                if aspect_ratio > (page_w / page_h):
                    # Image is wider, fit to width
                    display_w = page_w
                    display_h = page_w / aspect_ratio
                else:
                    # Image is taller, fit to height
                    display_h = page_h
                    display_w = page_h * aspect_ratio
                
                # Center the image
                x = (pdf.w - display_w) / 2
                y = (pdf.h - display_h) / 2
                
                pdf.image(page.full_image_path, x=x, y=y, w=display_w, h=display_h)
            except Exception as e:
                print(f"Could not add full image {page.full_image_path} to PDF. Error: {e}")
                # Show placeholder text if image fails
                pdf.set_font("Arial", '', 16)
                pdf.set_y(pdf.h / 2)
                pdf.cell(0, 10, "Image could not be loaded", 0, 1, 'C')
        else:
            # No image assigned, show placeholder
            pdf.set_font("Arial", '', 16)
            pdf.set_y(pdf.h / 2)
            pdf.cell(0, 10, "No image assigned", 0, 1, 'C')
    
    def upload_to_sharepoint(self):
        """Upload the current project to SharePoint."""
        # Save the project first if it hasn't been saved
        if not self.project_file:
            messagebox.showwarning("Save Required", "Please save the project before uploading to SharePoint.")
            self.save_project()
            if not self.project_file:  # User cancelled save
                return
        else:
            # Save current changes
            self.save_current_page_data()
            try:
                data = {'pages': [page.to_dict() for page in self.pages]}
                with open(self.project_file, 'w') as f:
                    json.dump(data, f, indent=4)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save project: {e}")
                return
        
        try:
            upload_to_sharepoint(self, self.project_file)
        except Exception as e:
            messagebox.showerror("Upload Error", f"Failed to upload to SharePoint: {str(e)}")
