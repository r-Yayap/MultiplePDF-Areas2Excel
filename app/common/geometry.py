def adjust_coordinates_for_rotation(coordinates, rotation, pdf_height, pdf_width):
    """
    Adjust area coordinates [x0,y0,x1,y1] for page rotation (0/90/180/270).
    """
    if rotation == 0:
        return coordinates
    x0, y0, x1, y1 = coordinates
    if rotation == 90:
        return [y0, pdf_width - x1, y1, pdf_width - x0]
    if rotation == 180:
        return [pdf_width - x1, pdf_height - y1, pdf_width - x0, pdf_height - y0]
    if rotation == 270:
        return [pdf_height - y1, x0, pdf_height - y0, x1]
    raise ValueError("Invalid rotation angle. Must be 0, 90, 180, or 270 degrees.")
