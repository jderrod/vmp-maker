import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import json
from .page import Page

class EditorPage(tk.Frame):
    """Editor page for VMP Tool - handles editing and page navigation."""

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.images_dir = os.path.join(os.getcwd(), "VMP-Images")
        self.project_file = None
        self.pages = [Page()]
        self.current_page_index = 0
        self.selected_gallery_image_path = None

        self.setup_ui()
        self.show_page() # Show initial blank page

    def load_data(self, project_file=None):
        """Loads a project or resets to a new one."""
        if project_file and os.path.exists(project_file):
            self.load_project(project_file)
        else:
            self.project_file = None
            self.pages = [Page()]
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

        right_buttons = tk.Frame(nav_frame, bg='#bdc3c7')
        right_buttons.grid(row=0, column=2, padx=10, pady=5)

        self.add_page_btn = tk.Button(right_buttons, text="Add Page", command=self.add_page)
        self.add_page_btn.pack(side=tk.LEFT, padx=5)

        self.delete_page_btn = tk.Button(right_buttons, text="Delete Page", command=self.delete_page)
        self.delete_page_btn.pack(side=tk.LEFT, padx=5)

        self.save_btn = tk.Button(right_buttons, text="Save Project", command=self.save_project)
        self.save_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = tk.Button(right_buttons, text="Export to PDF", command=self.export_to_pdf)
        self.export_btn.pack(side=tk.LEFT, padx=5)

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
        
        self.show_page()

    def show_page(self):
        """Displays the current page's content."""
        self.save_current_page_bullets()

        for widget in self.content_container.winfo_children():
            widget.destroy()

        page = self.pages[self.current_page_index]

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

        self.update_navigation()

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

    def save_current_page_bullets(self):
        """Saves the content of the bullet point text boxes."""
        if hasattr(self, 'bullet_texts') and self.bullet_texts:
            page = self.pages[self.current_page_index]
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
            self.save_current_page_bullets()
            self.current_page_index += 1
            self.show_page()

    def prev_page(self):
        if self.current_page_index > 0:
            self.save_current_page_bullets()
            self.current_page_index -= 1
            self.show_page()

    def add_page(self):
        self.save_current_page_bullets()
        self.pages.insert(self.current_page_index + 1, Page())
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
        self.save_current_page_bullets()
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
            self.pages = [Page.from_dict(pd) for pd in data['pages']]
            self.project_file = project_file
            self.current_page_index = 0
            self.show_page()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project: {e}")

    def export_to_pdf(self):
        """Exports the current project to a PDF file."""
        self.save_current_page_bullets()
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
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 10, f"Page {i + 1}", 0, 1, 'C')
            pdf.ln(10)

            col_width = (pdf.w - 2 * 15 - 10) / 2
            text_col_x = 15
            img_col_x = 15 + col_width + 10
            
            current_y = pdf.get_y()

            pdf.set_font("Arial", '', 12)
            for bullet in page.bullets:
                if bullet.strip():
                    pdf.set_x(text_col_x)
                    pdf.multi_cell(col_width, 10, f'* {bullet.strip()}')
                    pdf.ln(2)
            
            y_after_text = pdf.get_y()
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
            
            if y_after_text > pdf.get_y():
                pdf.set_y(y_after_text)

        try:
            pdf.output(save_path)
            messagebox.showinfo("Success", f"PDF exported to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")
