"""Image insertion tools."""


def register_image_tools(mcp, engine):
    @mcp.tool()
    def insert_image(image_path: str, width_inches: float = 0, height_inches: float = 0, path: str = "") -> dict:
        """Insert an image into the document.

        Args:
            image_path: Full path to the image file (PNG, JPG, BMP, etc.).
            width_inches: Optional image width in inches. 0 = auto.
            height_inches: Optional image height in inches. 0 = auto.
            path: Optional document path.
        """
        try:
            return {
                "success": True,
                **engine.insert_image(
                    image_path,
                    width_inches if width_inches > 0 else None,
                    height_inches if height_inches > 0 else None,
                    path if path else None,
                ),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
