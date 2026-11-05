import io
import os
import shutil
import tempfile

from visio_mcp.addressing import find_page
from visio_mcp.com_runtime import current_session, find_document
from visio_mcp.live_tool import live_tool

VIS_BBOX_UPRIGHT_WH = 1
VIS_RASTER_USE_CUSTOM = 3
VIS_RASTER_PIXEL_PER_INCH = 0
MAX_DPI = 192.0
MIN_DPI = 24.0
BADGE_FILL = (20, 60, 150)
BADGE_TEXT = (255, 255, 255)
BADGE_MARGIN = 12


def _export_png(page, dpi):
    """Export through a temp file. Export writes without changing focus or the Saved flag.
    The resolution is an app-wide setting the user owns, so it's always put back."""
    settings = page.Application.Settings
    previous = settings.GetRasterExportResolution()
    folder = tempfile.mkdtemp(prefix="visio-mcp-")
    path = os.path.join(folder, "page.png")
    try:
        if dpi is not None:
            settings.SetRasterExportResolution(VIS_RASTER_USE_CUSTOM, dpi, dpi, VIS_RASTER_PIXEL_PER_INCH)
        page.Export(path)
        with open(path, "rb") as exported:
            return exported.read()
    finally:
        try:
            settings.SetRasterExportResolution(*previous)
        finally:
            shutil.rmtree(folder, ignore_errors=True)


def _annotate(data, page, left, top, scale_x, scale_y):
    """Badge each shape's top-left corner with its ID, leaving its text readable. A margin
    keeps badges on shapes at the crop edge from being cut off."""
    from PIL import Image, ImageDraw, ImageFont, ImageOps

    with Image.open(io.BytesIO(data)) as source:
        image = ImageOps.expand(source.convert("RGB"), border=BADGE_MARGIN, fill=(255, 255, 255))
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default(size=13)
    shapes = page.Shapes
    for index in range(1, shapes.Count + 1):
        shape = shapes(index)
        if shape.OneD:
            continue
        shape_left, _, _, shape_top = shape.BoundingBox(VIS_BBOX_UPRIGHT_WH)
        x = BADGE_MARGIN + (shape_left - left) * scale_x
        y = BADGE_MARGIN + (top - shape_top) * scale_y
        label = str(shape.ID)
        box = draw.textbbox((x, y), label, font=font, anchor="mm")
        draw.rounded_rectangle((box[0] - 3, box[1] - 2, box[2] + 3, box[3] + 2), radius=3, fill=BADGE_FILL)
        draw.text((x, y), label, font=font, fill=BADGE_TEXT, anchor="mm")
    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def render_page(document, page, max_px=1400, annotate_ids=False):
    """Returns (metadata, PNG bytes). Export crops to the shapes' bounding box."""
    from PIL import Image

    if max_px < 64:
        raise ValueError("max_px must be at least 64.")
    left, bottom, right, top = page.BoundingBox(VIS_BBOX_UPRIGHT_WH)
    longest = max(right - left, top - bottom)
    dpi = None
    if longest > 0:
        dpi = max(MIN_DPI, min(MAX_DPI, max_px / longest))
    data = _export_png(page, dpi)
    with Image.open(io.BytesIO(data)) as image:
        width, height = image.size
    if annotate_ids and longest > 0:
        data = _annotate(data, page, left, top, width / (right - left), height / (top - bottom))
        width += 2 * BADGE_MARGIN
        height += 2 * BADGE_MARGIN
    metadata = {
        "drawing": document.Name,
        "page": page.Name,
        "width_px": width,
        "height_px": height,
        "shapes": page.Shapes.Count,
    }
    if longest > 0:
        metadata["crop_inches"] = {"left": round(left, 3), "bottom": round(bottom, 3), "right": round(right, 3), "top": round(top, 3)}
    return metadata, data


@live_tool(mutates=False)
def visio_get_page_image(
    max_px: int = 1400,
    annotate_ids: bool = False,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Get a PNG of a page, to check how a diagram looks. Doesn't change focus or the drawing.

    The image is cropped to the shapes, and its longest side is at most max_px, unless a very
    large page needs more to stay legible. annotate_ids badges each shape with its ID.
    """
    from fastmcp.utilities.types import Image

    document = find_document(current_session().app, drawing)
    metadata, data = render_page(document, find_page(document, page), max_px, annotate_ids)
    return [metadata, Image(data=data, format="png")]
