class Page:
    """
    Represents a single page in the Visual Manufacturing Procedure.

    Each page contains three bullet points and paths for two images.
    """
    def __init__(self):
        """Initializes a new, empty page."""
        self.bullets = ["", "", ""]  # Three bullet points
        self.image_path1 = None
        self.image_path2 = None

    def to_dict(self):
        """Converts the page object to a dictionary for JSON serialization."""
        return {
            'bullets': self.bullets,
            'image_path1': self.image_path1,
            'image_path2': self.image_path2
        }

    @classmethod
    def from_dict(cls, data):
        """Creates a page object from a dictionary."""
        page = cls()
        page.bullets = data.get('bullets', ["", "", ""])
        page.image_path1 = data.get('image_path1')
        page.image_path2 = data.get('image_path2')
        return page

