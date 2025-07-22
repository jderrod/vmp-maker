#!/usr/bin/env python3
"""
Main entry point for the Visual Manufacturing Procedures application.
This file creates the main App controller that manages all frames.
"""

import tkinter as tk
from src.home_page import HomePage
from src.main_window import EditorPage

class App(tk.Tk):
    """Main application controller."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Visual Manufacturing Procedures Tool")
        self.geometry("1280x720")

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (HomePage, EditorPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("HomePage")

    def show_frame(self, page_name, project_file=None):
        """Shows a frame for the given page name."""
        frame = self.frames[page_name]
        if page_name == "EditorPage":
            # Use a dedicated method in EditorPage to load data
            frame.load_data(project_file)
        elif page_name == "HomePage":
            # Refresh the project list every time we show the home page
            frame.refresh_project_list()
            
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()
