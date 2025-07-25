class Page:
    """
    Represents a single page in the Visual Manufacturing Procedure.

    Supports three page types:
    - 'title': Title page with centered title text
    - 'standard': Standard page with 3 bullet points and 2 images
    - 'full_image': Full page with single large image
    """
    def __init__(self, page_type='title'):
        """Initializes a new page of the specified type."""
        self.page_type = page_type
        
        # Common attributes
        self.title = ""  # Used for title pages
        
        # Standard page attributes
        self.bullets = ["", "", ""]  # Three bullet points
        self.image_path1 = None
        self.image_path2 = None
        
        # Full image page attributes
        self.full_image_path = None

    def to_dict(self):
        """Converts the page object to a dictionary for JSON serialization."""
        return {
            'page_type': self.page_type,
            'title': self.title,
            'bullets': self.bullets,
            'image_path1': self.image_path1,
            'image_path2': self.image_path2,
            'full_image_path': self.full_image_path
        }

    @classmethod
    def from_dict(cls, data):
        """Creates a page object from a dictionary."""
        page_type = data.get('page_type', 'title')
        page = cls(page_type)
        page.title = data.get('title', "")
        page.bullets = data.get('bullets', ["", "", ""])
        page.image_path1 = data.get('image_path1')
        page.image_path2 = data.get('image_path2')
        page.full_image_path = data.get('full_image_path')
        return page

