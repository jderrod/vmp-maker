#!/usr/bin/env python3
"""
Main entry point for the Visual Manufacturing Procedures application.
"""

from src.home_window import HomeWindow

if __name__ == "__main__":
    """Main entry point for the application."""
    # Before running, ensure you have installed the required libraries:
    # pip install tkinterdnd2-universal Pillow
    app = HomeWindow()
    app.mainloop()
