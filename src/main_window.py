import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from .page import Page

# Import reportlab components
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT

class MainWindow(tk.Tk):
    """Main window for the VMP application."""
    def __init__(self):
        super().__init__()
        self.title("Visual Manufacturing Procedures")
        self.geometry("900x700") # Increased window size for larger content

        self.pages = [Page()]
        self.current_page_index = 0
        self.bullet_entries = []

        # Main container frame
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # This frame will hold all the content rows and be centered
        self.content_container = tk.Frame(main_frame)
        self.content_container.pack(expand=True)

        # Controls Frame
        controls_frame = tk.Frame(self)
        controls_frame.pack(fill=tk.X, padx=10, pady=5)

        self.prev_button = tk.Button(controls_frame, text="Previous Page", command=self.prev_page)
        self.prev_button.pack(side=tk.LEFT)

        self.page_label = tk.Label(controls_frame, text="")
        self.page_label.pack(side=tk.LEFT, expand=True)

        self.next_button = tk.Button(controls_frame, text="Next Page", command=self.next_page)
        self.next_button.pack(side=tk.RIGHT)

        add_page_button = tk.Button(controls_frame, text="Add Page", command=self.add_page)
        add_page_button.pack(side=tk.RIGHT, padx=5)

        export_button = tk.Button(controls_frame, text="Export to PDF", command=self.export_to_pdf)
        export_button.pack(side=tk.RIGHT)

        self.show_page()

    def show_page(self):
        """Displays the current page's content."""
        self.save_current_page_bullets()

        # Clear previous content
        for widget in self.content_container.winfo_children():
            widget.destroy()

        self.bullet_entries = []
        self.image_placeholders = []
        self.select_image_buttons = []

        page = self.pages[self.current_page_index]

        # Create a two-column layout
        left_frame = tk.Frame(self.content_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        right_frame = tk.Frame(self.content_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Populate Left Frame with 3 Text Boxes ---
        for i in range(3):
            text_widget = tk.Text(left_frame, width=45, height=4, wrap=tk.WORD, font=("Helvetica", 10))
            text_widget.insert("1.0", page.bullets[i])
            text_widget.pack(pady=5, fill=tk.X, expand=True)
            text_widget.bind("<KeyRelease>", self._on_key_release)
            self.bullet_entries.append(text_widget)

        # --- Populate Right Frame with 2 Image Placeholders ---
        for i in range(2):
            # Create a container for the image and its button
            image_container = tk.Frame(right_frame)
            image_container.pack(pady=5, fill=tk.BOTH, expand=True)

            # Use a Label as the placeholder for easier image handling and centering
            placeholder = tk.Label(image_container, bg='lightgrey', relief=tk.RAISED, borderwidth=1)
            placeholder.pack(fill=tk.BOTH, expand=True)

            if page.images[i]:
                self.show_image(placeholder, page.images[i])

            select_button = tk.Button(image_container, text="Select Image", command=lambda i=i: self._select_image(i))
            select_button.pack(side=tk.BOTTOM, pady=5)

            self.image_placeholders.append(placeholder)
            self.select_image_buttons.append(select_button)

        self.update_navigation()

    def show_image(self, container, image_path):
        """Loads and displays a dynamically resized image in a given container."""
        # This event is triggered when the container (placeholder) is resized
        def on_configure(event):
            try:
                # Get container size
                w, h = event.width, event.height

                if w <= 1 or h <= 1: return # Avoid processing tiny sizes

                img = Image.open(image_path)
                img.thumbnail((w - 10, h - 40), Image.Resampling.LANCZOS) # Resize to fit, leaving space for button
                photo_img = ImageTk.PhotoImage(img)

                container.config(image=photo_img)
                container.image = photo_img # Keep a reference!
            except Exception as e:
                container.config(text=f"Invalid Image\n{e}", bg='red', fg='white')

        container.bind('<Configure>', on_configure)

    def _select_image(self, index):
        """Opens a file dialog to select an image."""
        file_path = filedialog.askopenfilename(
            title="Select an Image",
            filetypes=(("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*"))
        )
        if file_path:
            page = self.pages[self.current_page_index]
            page.images[index] = file_path

            # Directly update the image without redrawing the whole page
            placeholder = self.image_placeholders[index]
            self.show_image(placeholder, file_path)

    def _on_key_release(self, event):
        """Handle key release to adjust widget height and enforce char limit."""
        text_widget = event.widget
        # Enforce 200 character limit
        content = text_widget.get("1.0", "end-1c") # -1c to exclude trailing newline
        if len(content) > 200:
            text_widget.delete("1.200", tk.END)
        
        self._adjust_text_widget_height(text_widget)

    def _adjust_text_widget_height(self, text_widget):
        """Adjust the height of a text widget to fit its content."""
        # Get the number of lines in the widget
        lines = text_widget.index(tk.END).split('.')[0]
        # The height is the number of lines, but at least 1
        new_height = max(1, int(lines) - 1)
        text_widget.config(height=new_height)

    def save_current_page_bullets(self):
        """Saves the content of the bullet point Text widgets."""
        if self.bullet_entries:
            page = self.pages[self.current_page_index]
            for i, text_widget in enumerate(self.bullet_entries):
                page.bullets[i] = text_widget.get("1.0", "end-1c")

    def update_navigation(self):
        """Updates the page navigation controls."""
        self.page_label.config(text=f"Page {self.current_page_index + 1} of {len(self.pages)}")
        self.prev_button.config(state=tk.NORMAL if self.current_page_index > 0 else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if self.current_page_index < len(self.pages) - 1 else tk.DISABLED)

    def next_page(self):
        """Goes to the next page."""
        if self.current_page_index < len(self.pages) - 1:
            self.save_current_page_bullets()
            self.current_page_index += 1
            self.show_page()

    def prev_page(self):
        """Goes to the previous page."""
        if self.current_page_index > 0:
            self.save_current_page_bullets()
            self.current_page_index -= 1
            self.show_page()

    def add_page(self):
        """Adds a new page."""
        self.save_current_page_bullets()
        self.pages.append(Page())
        self.current_page_index = len(self.pages) - 1
        self.show_page()

    def export_to_pdf(self):
        """Exports the VMP to a PDF file."""
        self.save_current_page_bullets()

        file_path = filedialog.asksaveasfilename(
            title="Save PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")]
        )

        if not file_path:
            return

        try:
            c = canvas.Canvas(file_path, pagesize=landscape(letter))
            width, height = landscape(letter)

            for i, page_data in enumerate(self.pages):
                # --- Draw Page Number ---
                c.drawString(0.5 * inch, height - 0.5 * inch, f"Page {i + 1} of {len(self.pages)}")

                styles = getSampleStyleSheet()
                bullet_style = ParagraphStyle(
                    'bulletStyle', parent=styles['BodyText'], fontName='Helvetica', fontSize=12, leading=14,
                    leftIndent=20, firstLineIndent=-10, alignment=TA_LEFT
                )

                # --- Define Layout --- 
                left_margin = 0.5 * inch
                right_margin = 0.5 * inch
                top_margin = 0.5 * inch
                bottom_margin = 0.5 * inch
                gutter = 0.5 * inch

                available_width = width - left_margin - right_margin - gutter
                text_area_width = available_width / 2
                image_area_width = available_width / 2

                available_height = height - top_margin - bottom_margin

                # --- Draw Left Column (3 Text Boxes) ---
                text_col_x = left_margin
                text_row_height = available_height / 3.0

                for j, bullet_text in enumerate(page_data.bullets):
                    text_y_center = (height - top_margin) - (j * text_row_height) - (text_row_height / 2)
                    p = Paragraph(f"\u2022<font name=Helvetica>{bullet_text or ''}</font>", bullet_style)
                    p_w, p_h = p.wrapOn(c, text_area_width, text_row_height)
                    p.drawOn(c, text_col_x, text_y_center - p_h / 2)

                # --- Draw Right Column (2 Images) ---
                image_col_x = left_margin + text_area_width + gutter
                image_row_height = available_height / 2.0 # Two images share the height

                for j, img_path in enumerate(page_data.images):
                    image_y_center = (height - top_margin) - (j * image_row_height) - (image_row_height / 2)
                    
                    img_x = image_col_x
                    img_y = image_y_center - (image_row_height / 2) + (0.1 * inch)
                    img_w = image_area_width
                    img_h = image_row_height - (0.2 * inch)

                    c.rect(img_x, img_y, img_w, img_h, stroke=1, fill=0)
                    if img_path:
                        try:
                            c.drawImage(ImageReader(img_path), img_x, img_y, width=img_w, height=img_h, preserveAspectRatio=True, anchor='c')
                        except Exception:
                            c.drawString(img_x + 0.1 * inch, image_y_center, "Invalid Image")

                c.showPage()
            
            c.save()
            messagebox.showinfo("Success", f"Successfully exported PDF to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")
