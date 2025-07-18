class Page:
    """
    Represents a single page in the Visual Manufacturing Procedure.

    Each page contains a list of bullet points (as strings) and a list of image paths.
    """
    def __init__(self):
        """Initializes a new, empty page."""
        self.bullets = ["", "", ""] # Three bullet points
        self.images = [None, None]    # Two images
