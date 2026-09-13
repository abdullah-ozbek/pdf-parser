
from flask import Flask, request, jsonify, send_file

import fitz
import json
import re
import statistics
import os
import gzip
import time
import uuid
from difflib import SequenceMatcher

from io import BytesIO

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


app = Flask(__name__)

# ============================================================
# CONFIG
# ============================================================

CHUNK_SIZE = 40

# Extracted PDF structures are kept temporarily on the Render instance so
# Make.com only needs to pass a small structure_id to /create-docx.
STRUCTURE_STORE_DIR = os.environ.get(
    "STRUCTURE_STORE_DIR",
    "/tmp/pdf_translator_structures",
)
STRUCTURE_TTL_SECONDS = int(os.environ.get(
    "STRUCTURE_TTL_SECONDS",
    str(6 * 60 * 60),
))

DEFAULT_FONT_NAME = "Arial"
DEFAULT_FONT_SIZE = 10
TABLE_FONT_SIZE = 9

MIN_VERTICAL_GAP_PT = 0
MAX_VERTICAL_GAP_PT = 40

ENABLE_TABLE_DETECTION = True
ENABLE_DRAWING_DETECTION = True
ENABLE_LIST_DETECTION = True
ENABLE_STYLE_DETECTION = True
ENABLE_COLUMN_DETECTION = True
ENABLE_FORM_DETECTION = True

# Spatial duplicate suppression. PDF generators sometimes expose the same
# visible text more than once (for example as overlapping text objects).
# Suppress those duplicates BEFORE pdf_N IDs are assigned.
ENABLE_SPATIAL_DUPLICATE_DETECTION = True
SPATIAL_DUPLICATE_IOU = 0.72
SPATIAL_DUPLICATE_CONTAINMENT = 0.86
SPATIAL_DUPLICATE_TEXT_SIMILARITY = 0.90
SPATIAL_DUPLICATE_CENTER_TOLERANCE_PT = 5.0
ENABLE_FORM_DOCX_REBUILD = True

HEADING_SIZE_RATIO = 1.18
HEADING_LARGE_RATIO = 1.45

COLUMN_MIN_BLOCKS = 4
COLUMN_GAP_RATIO = 0.08

LINE_TOLERANCE = 1.5

THIN_RECT_MAX_THICKNESS = 3.0
MIN_LINE_LENGTH = 8.0

FORM_MIN_LINE_RATIO = 0.08
FORM_MAX_LINE_RATIO = 0.80
FORM_LABEL_MAX_ABOVE_PT = 30.0
FORM_LABEL_MAX_BELOW_PT = 10.0
FORM_ROW_Y_TOLERANCE = 8.0
FORM_COLUMN_X_TOLERANCE = 20.0
FORM_MIN_FIELD_WIDTH = 35.0
FORM_LABEL_X_TOLERANCE = 12.0
FORM_VERTICAL_BORDER_TOLERANCE = 4.0

FORM_PAGE_MIN_FIELDS = 8
FORM_PAGE_MIN_ROWS = 4
FORM_DOCX_FONT_PT = 7.5
FORM_DOCX_HEADING_FONT_PT = 9.0
FORM_DOCX_ROW_GAP_MAX_PT = 10.0
FORM_DOCX_LEFT_RIGHT_SAFETY_PT = 3.0
FORM_DOCX_MIN_CELL_WIDTH_PT = 24.0

# Generic form-layout preservation. PDF forms differ: some place the field
# label below the writing line, others above it. Preserve the detected visual
# relationship instead of imposing one fixed layout.
FORM_LABEL_ON_LINE_TOLERANCE_PT = 2.5
FORM_LABEL_COLORED_MAX_DISTANCE_PT = 4.0
FORM_FIELD_BLANK_HEIGHT_PT = 13.0

# Colored + emphasized text around form geometry is usually a section heading
# rather than a fillable-field label. Keep this generic: no page numbers,
# document-specific words, or fixed coordinates are used.
FORM_SECTION_HEADING_MIN_FONT_PT = 8.5

# Footer detection
ENABLE_FOOTER_DETECTION = True
FOOTER_SEARCH_REGION_START_RATIO = 0.60
FOOTER_REGION_START_RATIO = 0.60
FOOTER_HARD_REGION_START_RATIO = 0.93
FOOTER_REPEAT_MIN_PAGES = 2
FOOTER_POSITION_TOLERANCE_RATIO = 0.08
FOOTER_PAGE_MARKER_REGEX = re.compile(
    r"\b(seite|sayfa|page|stand|version|sürüm|tarih|datum)\b",
    re.IGNORECASE,
)

# Header / letterhead detection
ENABLE_HEADER_DETECTION = True
HEADER_REGION_END_RATIO = 0.36
HEADER_REPEAT_MIN_PAGES = 2
HEADER_POSITION_TOLERANCE_RATIO = 0.08
# Unique first-page letterheads should be compact. Large upper-page body
# blocks that happen to contain an URL / e-mail must never become Word
# headers, otherwise Word repeats that body text on every overflow page.
HEADER_UNIQUE_MAX_HEIGHT_RATIO = 0.14
HEADER_UNIQUE_MAX_TEXT_LENGTH = 420
HEADER_UNIQUE_MAX_LINES = 10
HEADER_CONTACT_REGEX = re.compile(
    r"(gmbh|ag\b|kg\b|straße|strasse|telefon|tel\.?\s|telefax|fax\b|"
    r"e-?mail|@|www\.|https?://|iban|bic|bank|sparkasse|\b\d{5}\s+[A-ZÄÖÜ])",
    re.IGNORECASE,
)

# Generic form labels can contain words such as "Straße" or "E-Mail".
# Those are not letterheads and must never be promoted into a Word header.
HEADER_FORM_LABEL_REGEX = re.compile(
    r"(straße\s*,?\s*hausnr|strasse\s*,?\s*hausnr|plz\s*,?\s*ort|"
    r"geburtsdatum|vorname|akademischer\s+grad|kundennummer|förderstelle|"
    r"foerderstelle|teilnehmer/?-?in|anmeldung)",
    re.IGNORECASE,
)

# ============================================================
# ROOT
# ============================================================

@app.route("/", methods=["GET"])
def home():
    return {
        "status": "ok",
        "service": "PDF Structure Parser"
    }

# ============================================================
# GENERAL HELPERS
# ============================================================

def clamp(value, minimum, maximum):
    return max(minimum, min(maximum, value))


def round_num(value):
    try:
        return round(float(value), 2)
    except Exception:
        return 0.0


def bbox_dict(rect):
    return {
        "x0": round_num(rect[0]),
        "y0": round_num(rect[1]),
        "x1": round_num(rect[2]),
        "y1": round_num(rect[3]),
    }


def clean_text(value):
    if value is None:
        return ""
    value = str(value).replace("\u00ad", "").replace("\u00a0", " ")
    return value.strip()


def rect_center_inside(inner_rect, outer_rect):
    cx = (inner_rect[0] + inner_rect[2]) / 2
    cy = (inner_rect[1] + inner_rect[3]) / 2
    return (
        outer_rect[0] <= cx <= outer_rect[2]
        and outer_rect[1] <= cy <= outer_rect[3]
    )


def horizontal_overlap(x0_a, x1_a, x0_b, x1_b):
    return max(0.0, min(x1_a, x1_b) - max(x0_a, x0_b))


def bbox_overlap_ratio(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix = max(0.0, min(ax1, bx1) - max(ax0, bx0))
    iy = max(0.0, min(ay1, by1) - max(ay0, by0))
    inter = ix * iy
    area = max(1.0, (ax1 - ax0) * (ay1 - ay0))
    return inter / area


def _rect_metrics(a, b):
    ax0, ay0, ax1, ay1 = map(float, a)
    bx0, by0, bx1, by1 = map(float, b)
    iw = max(0.0, min(ax1, bx1) - max(ax0, bx0))
    ih = max(0.0, min(ay1, by1) - max(ay0, by0))
    inter = iw * ih
    area_a = max(1.0, (ax1 - ax0) * (ay1 - ay0))
    area_b = max(1.0, (bx1 - bx0) * (by1 - by0))
    union = max(1.0, area_a + area_b - inter)
    return inter / union, inter / min(area_a, area_b)


def _normalized_spatial_text(text):
    text = clean_text(text).lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\wäöüßçğıöşü]+", "", text, flags=re.UNICODE)
    return text


def _block_rect(block):
    b = block.get("bbox", {})
    return (
        float(b.get("x0", 0)), float(b.get("y0", 0)),
        float(b.get("x1", 0)), float(b.get("y1", 0)),
    )


def _block_quality(block):
    # Prefer the richer extraction when two objects describe the same region.
    text = clean_text(block.get("text", ""))
    spans = block.get("spans", []) or []
    lines = block.get("lines", []) or []
    return (len(text), len(spans), len(lines))


def blocks_are_spatial_duplicates(a, b):
    ta = _normalized_spatial_text(a.get("text", ""))
    tb = _normalized_spatial_text(b.get("text", ""))
    if not ta or not tb:
        return False

    ra, rb = _block_rect(a), _block_rect(b)
    iou, containment = _rect_metrics(ra, rb)
    acx, acy = (ra[0] + ra[2]) / 2, (ra[1] + ra[3]) / 2
    bcx, bcy = (rb[0] + rb[2]) / 2, (rb[1] + rb[3]) / 2
    centers_close = (
        abs(acx - bcx) <= SPATIAL_DUPLICATE_CENTER_TOLERANCE_PT
        and abs(acy - bcy) <= SPATIAL_DUPLICATE_CENTER_TOLERANCE_PT
    )

    # Exact normalized text needs only strong geometric evidence.
    if ta == tb:
        return (
            iou >= SPATIAL_DUPLICATE_IOU
            or containment >= SPATIAL_DUPLICATE_CONTAINMENT
            or centers_close
        )

    # Near-identical text must also substantially occupy the same region.
    similarity = SequenceMatcher(None, ta, tb).ratio()
    return (
        similarity >= SPATIAL_DUPLICATE_TEXT_SIMILARITY
        and (
            iou >= SPATIAL_DUPLICATE_IOU
            or containment >= SPATIAL_DUPLICATE_CONTAINMENT
        )
    )


def suppress_spatial_duplicate_blocks(blocks):
    if not ENABLE_SPATIAL_DUPLICATE_DETECTION:
        return blocks, []

    kept = []
    suppressed = []
    # Stable visual order makes the result deterministic across /extract-pdf
    # and /create-docx, which is essential because translations use pdf_N IDs.
    ordered = sorted(
        blocks,
        key=lambda b: (
            round(float(b.get("bbox", {}).get("y0", 0)), 2),
            round(float(b.get("bbox", {}).get("x0", 0)), 2),
            -len(clean_text(b.get("text", ""))),
        ),
    )

    for block in ordered:
        duplicate_index = None
        for i, existing in enumerate(kept):
            if blocks_are_spatial_duplicates(block, existing):
                duplicate_index = i
                break

        if duplicate_index is None:
            kept.append(block)
            continue

        existing = kept[duplicate_index]
        if _block_quality(block) > _block_quality(existing):
            kept[duplicate_index] = block
            suppressed.append(existing)
        else:
            suppressed.append(block)

    kept.sort(key=lambda b: (
        float(b.get("bbox", {}).get("y0", 0)),
        float(b.get("bbox", {}).get("x0", 0)),
    ))
    return kept, suppressed

# ============================================================
# FONT / STYLE
# ============================================================

def rgb_from_int(color_value):
    try:
        color_value = int(color_value)
        r = (color_value >> 16) & 255
        g = (color_value >> 8) & 255
        b = color_value & 255
        return {
            "r": r,
            "g": g,
            "b": b,
            "hex": f"#{r:02X}{g:02X}{b:02X}",
        }
    except Exception:
        return {"r": 0, "g": 0, "b": 0, "hex": "#000000"}


def span_is_bold(span):
    flags = int(span.get("flags", 0))
    font = str(span.get("font", "")).lower()
    return (
        bool(flags & 16)
        or "bold" in font
        or "black" in font
        or "heavy" in font
        or "semibold" in font
        or "demi" in font
    )


def span_is_italic(span):
    flags = int(span.get("flags", 0))
    font = str(span.get("font", "")).lower()
    return bool(flags & 2) or "italic" in font or "oblique" in font


def span_is_monospace(span):
    return bool(int(span.get("flags", 0)) & 8)


def span_is_serif(span):
    return bool(int(span.get("flags", 0)) & 4)


def normalize_span(span):
    bbox = span.get("bbox", (0, 0, 0, 0))
    return {
        "text": clean_text(span.get("text", "")),
        "font": str(span.get("font", "")),
        "size": round_num(span.get("size", 0)),
        "bold": span_is_bold(span),
        "italic": span_is_italic(span),
        "monospace": span_is_monospace(span),
        "serif": span_is_serif(span),
        "color": rgb_from_int(span.get("color", 0)),
        "bbox": bbox_dict(bbox),
    }


def get_block_style(spans):
    valid = [s for s in spans if s.get("text")]
    if not valid:
        return {
            "font": "",
            "size": 0,
            "max_size": 0,
            "bold": False,
            "italic": False,
            "color": "#000000",
        }

    sizes = [
        float(s.get("size", 0))
        for s in valid
        if float(s.get("size", 0)) > 0
    ]
    fonts = [s.get("font", "") for s in valid if s.get("font")]
    colors = [s.get("color", {}).get("hex", "#000000") for s in valid]

    try:
        font = statistics.mode(fonts) if fonts else ""
    except Exception:
        font = fonts[0] if fonts else ""

    try:
        color = statistics.mode(colors) if colors else "#000000"
    except Exception:
        color = colors[0] if colors else "#000000"

    return {
        "font": font,
        "size": round_num(statistics.median(sizes)) if sizes else 0,
        "max_size": round_num(max(sizes)) if sizes else 0,
        "bold": any(s.get("bold") for s in valid),
        "italic": any(s.get("italic") for s in valid),
        "color": color,
    }

# ============================================================
# LIST DETECTION
# ============================================================

BULLET_REGEX = re.compile(
    r"^\s*([•●▪◦‣⁃·]|[\uF0B7]|[-–—])\s+"
)

NUMBER_LIST_REGEX = re.compile(
    r"^\s*(\d+[\.\)]|\(\d+\)|[a-zA-Z][\.\)]|\([a-zA-Z]\))\s+"
)


def detect_list_info(text):
    if not ENABLE_LIST_DETECTION or not text:
        return None

    first_line = text.splitlines()[0].strip()

    match = BULLET_REGEX.match(first_line)
    if match:
        return {
            "is_list": True,
            "list_type": "bullet",
            "marker": match.group(1),
        }

    match = NUMBER_LIST_REGEX.match(first_line)
    if match:
        return {
            "is_list": True,
            "list_type": "numbered",
            "marker": match.group(1),
        }

    return None

# ============================================================
# DRAWINGS
# ============================================================

def point_xy(point):
    try:
        return float(point.x), float(point.y)
    except Exception:
        try:
            return float(point[0]), float(point[1])
        except Exception:
            return 0.0, 0.0


def normalize_drawings(page):
    if not ENABLE_DRAWING_DETECTION:
        return [], [], []

    drawings_out = []
    lines_out = []
    rectangles_out = []

    try:
        drawings = page.get_drawings()
    except Exception as e:
        print("get_drawings error:", str(e))
        return [], [], []

    drawing_counter = 1
    line_counter = 1
    rect_counter = 1

    for drawing in drawings:
        drawing_bbox = drawing.get("rect")

        drawing_data = {
            "id": f"drawing_{drawing_counter}",
            "type": "drawing",
            "bbox": bbox_dict(drawing_bbox) if drawing_bbox is not None else None,
            "fill": str(drawing.get("fill")),
            "color": str(drawing.get("color")),
            "width": round_num(drawing.get("width", 0)),
        }
        drawing_counter += 1

        drawing_items = []

        for item in drawing.get("items", []):
            if not item:
                continue

            command = item[0]

            if command == "l" and len(item) >= 3:
                x0, y0 = point_xy(item[1])
                x1, y1 = point_xy(item[2])

                horizontal = abs(y1 - y0) <= LINE_TOLERANCE
                vertical = abs(x1 - x0) <= LINE_TOLERANCE

                orientation = "diagonal"
                if horizontal:
                    orientation = "horizontal"
                elif vertical:
                    orientation = "vertical"

                length = (((x1 - x0) ** 2) + ((y1 - y0) ** 2)) ** 0.5

                line = {
                    "id": f"line_{line_counter}",
                    "type": "line",
                    "source": "line",
                    "x0": round_num(x0),
                    "y0": round_num(y0),
                    "x1": round_num(x1),
                    "y1": round_num(y1),
                    "orientation": orientation,
                    "length": round_num(length),
                    "thickness": round_num(drawing.get("width", 0)),
                }

                line_counter += 1
                lines_out.append(line)
                drawing_items.append(line)

            elif command == "re" and len(item) >= 2:
                rect = item[1]

                try:
                    x0 = float(rect.x0)
                    y0 = float(rect.y0)
                    x1 = float(rect.x1)
                    y1 = float(rect.y1)

                    rect_width = abs(x1 - x0)
                    rect_height = abs(y1 - y0)

                    if (
                        rect_height <= THIN_RECT_MAX_THICKNESS
                        and rect_width >= MIN_LINE_LENGTH
                    ):
                        center_y = (y0 + y1) / 2

                        line = {
                            "id": f"line_{line_counter}",
                            "type": "line",
                            "source": "thin_rectangle",
                            "x0": round_num(min(x0, x1)),
                            "y0": round_num(center_y),
                            "x1": round_num(max(x0, x1)),
                            "y1": round_num(center_y),
                            "orientation": "horizontal",
                            "length": round_num(rect_width),
                            "thickness": round_num(rect_height),
                        }

                        line_counter += 1
                        lines_out.append(line)
                        drawing_items.append(line)
                        continue

                    if (
                        rect_width <= THIN_RECT_MAX_THICKNESS
                        and rect_height >= MIN_LINE_LENGTH
                    ):
                        center_x = (x0 + x1) / 2

                        line = {
                            "id": f"line_{line_counter}",
                            "type": "line",
                            "source": "thin_rectangle",
                            "x0": round_num(center_x),
                            "y0": round_num(min(y0, y1)),
                            "x1": round_num(center_x),
                            "y1": round_num(max(y0, y1)),
                            "orientation": "vertical",
                            "length": round_num(rect_height),
                            "thickness": round_num(rect_width),
                        }

                        line_counter += 1
                        lines_out.append(line)
                        drawing_items.append(line)
                        continue

                    rect_data = {
                        "id": f"rect_{rect_counter}",
                        "type": "rectangle",
                        "bbox": bbox_dict((x0, y0, x1, y1)),
                        "width": round_num(rect_width),
                        "height": round_num(rect_height),
                    }

                    rect_counter += 1
                    rectangles_out.append(rect_data)
                    drawing_items.append(rect_data)

                except Exception as e:
                    print("Rectangle parse error:", str(e))

        drawing_data["items"] = drawing_items
        if drawing_items:
            drawings_out.append(drawing_data)

    return drawings_out, lines_out, rectangles_out

# ============================================================
# FORM DETECTION
# ============================================================

def detect_form_lines(lines, page_width):
    result = []
    minimum = page_width * FORM_MIN_LINE_RATIO
    maximum = page_width * FORM_MAX_LINE_RATIO

    for line in lines:
        if line.get("orientation") != "horizontal":
            continue

        length = float(line.get("length", 0))
        if minimum <= length <= maximum:
            result.append(line)

    return result


def find_vertical_boundaries(field_line, vertical_lines):
    x0 = float(field_line.get("x0", 0))
    x1 = float(field_line.get("x1", 0))
    y = float(field_line.get("y0", 0))

    left_border = False
    right_border = False
    internal_borders = []

    for line in vertical_lines:
        vx = float(line.get("x0", 0))
        vy0 = min(float(line.get("y0", 0)), float(line.get("y1", 0)))
        vy1 = max(float(line.get("y0", 0)), float(line.get("y1", 0)))

        if not (
            vy0 - FORM_VERTICAL_BORDER_TOLERANCE
            <= y
            <= vy1 + FORM_VERTICAL_BORDER_TOLERANCE
        ):
            continue

        if abs(vx - x0) <= FORM_VERTICAL_BORDER_TOLERANCE:
            left_border = True
        elif abs(vx - x1) <= FORM_VERTICAL_BORDER_TOLERANCE:
            right_border = True
        elif x0 < vx < x1:
            internal_borders.append(round_num(vx))

    return {
        "left_border": left_border,
        "right_border": right_border,
        "internal_borders": sorted(list(set(internal_borders))),
    }


def find_form_label(field_line, raw_blocks):
    """Find the nearest *text line* above a form underline.

    Earlier versions used the whole PDF text block as a label. One block can
    contain many form labels, so the same German block was then reused in
    several reconstructed cells. Working at line granularity keeps every
    inferred field tied to one visual label line while still translating the
    original block only once.
    """
    fx0 = float(field_line.get("x0", 0))
    fx1 = float(field_line.get("x1", 0))
    fy = float(field_line.get("y0", 0))

    candidates = []

    for block in raw_blocks:
        block_lines = block.get("lines", []) or []

        # Fall back to the whole block only when PyMuPDF did not expose lines.
        if not block_lines:
            block_lines = [{
                "text": clean_text(block.get("text", "")),
                "bbox": block.get("bbox", {}),
                "_line_index": 0,
            }]

        for line_index, line_data in enumerate(block_lines):
            text = clean_text(line_data.get("text", ""))
            if not text:
                # normalize_raw_text_blocks stores spans on lines; rebuild text
                # defensively if a line-level text value is missing.
                text = clean_text("".join(
                    span.get("text", "")
                    for span in (line_data.get("spans", []) or [])
                ))
            if not text:
                continue

            bbox = line_data.get("bbox") or block.get("bbox", {})
            bx0 = float(bbox.get("x0", 0))
            by0 = float(bbox.get("y0", 0))
            bx1 = float(bbox.get("x1", 0))
            by1 = float(bbox.get("y1", 0))

            vertical_distance = fy - by1

            if (
                vertical_distance < -FORM_LABEL_MAX_BELOW_PT
                or vertical_distance > FORM_LABEL_MAX_ABOVE_PT
            ):
                continue

            overlap = horizontal_overlap(
                fx0 - FORM_LABEL_X_TOLERANCE,
                fx1 + FORM_LABEL_X_TOLERANCE,
                bx0,
                bx1,
            )

            if overlap <= 0:
                continue

            line_width = max(1.0, bx1 - bx0)
            overlap_ratio = overlap / line_width

            center_field = (fx0 + fx1) / 2
            center_text = (bx0 + bx1) / 2
            center_distance = abs(center_field - center_text)

            line_color = line_dominant_color(line_data)

            # A visually emphasized colored line is structural form text, not
            # a fillable-field label. Some PDFs draw a decorative underline
            # immediately beneath such headings; without this guard that line
            # was mistaken for an input field and the heading disappeared.
            if line_is_form_section_heading(line_data):
                continue

            # Other colored text is also less likely to be a field label when
            # it is not sitting directly on the writing line.
            if (
                not color_is_near_black(line_color)
                and abs(vertical_distance) > FORM_LABEL_COLORED_MAX_DISTANCE_PT
            ):
                continue

            # Form labels are usually aligned to the left edge of their
            # writing field.  Center-distance heavily penalises short labels
            # such as "Ort" / "City" and can make the matcher incorrectly
            # reuse a longer label from the previous row.  Prefer vertical
            # proximity + left-edge alignment, with only a small centre term.
            left_distance = abs(fx0 - bx0)
            score = (
                abs(vertical_distance) * 6
                + left_distance * 0.8
                + center_distance * 0.10
                - overlap_ratio * 20
            )

            candidates.append({
                "score": score,
                "text": text,
                "bbox": bbox,
                "line_index": line_index,
                "source_block_bbox": block.get("bbox", {}),
                "color": line_color,
            })

    if not candidates:
        return {
            "text": "",
            "bbox": None,
            "line_index": None,
            "source_block_bbox": None,
            "color": "#000000",
        }

    candidates.sort(key=lambda item: item["score"])
    best = candidates[0]

    return {
        "text": best["text"],
        "bbox": best["bbox"],
        "line_index": best["line_index"],
        "source_block_bbox": best["source_block_bbox"],
        "color": best.get("color", "#000000"),
    }


def line_dominant_color(line_data):
    colors = []
    for span in (line_data.get("spans", []) or []):
        color = (span.get("color") or {}).get("hex", "#000000")
        if color:
            colors.append(str(color).upper())
    if not colors:
        return "#000000"
    try:
        return statistics.mode(colors)
    except Exception:
        return colors[0]


def color_is_near_black(hex_color):
    try:
        value = str(hex_color or "#000000").lstrip("#")
        if len(value) != 6:
            return True
        r, g, b = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
        return max(r, g, b) <= 70
    except Exception:
        return True


def line_is_form_section_heading(line_data):
    """Detect an emphasized colored form section heading.

    This deliberately uses visual properties only. A colored bold/semibold
    line with a normal heading-sized font is treated as structural text and
    must not be consumed as the label of a nearby horizontal form line.
    """
    spans = [
        span for span in (line_data.get("spans", []) or [])
        if clean_text(span.get("text", ""))
    ]
    if not spans:
        return False

    color = line_dominant_color(line_data)
    if color_is_near_black(color):
        return False

    sizes = [float(span.get("size", 0) or 0) for span in spans]
    max_size = max(sizes) if sizes else 0.0
    emphasized = any(
        bool(span.get("bold"))
        or "bold" in str(span.get("font", "")).lower()
        or "semibold" in str(span.get("font", "")).lower()
        or "demi" in str(span.get("font", "")).lower()
        for span in spans
    )

    return emphasized and max_size >= FORM_SECTION_HEADING_MIN_FONT_PT


def classify_label_position(label_bbox, field_y):
    if not label_bbox:
        return "unknown"
    y0 = float(label_bbox.get("y0", field_y))
    y1 = float(label_bbox.get("y1", field_y))
    if y0 >= field_y - FORM_LABEL_ON_LINE_TOLERANCE_PT:
        return "below"
    if y1 <= field_y + FORM_LABEL_ON_LINE_TOLERANCE_PT:
        return "above"
    return "overlap"


def set_paragraph_bottom_border(paragraph, size=5, color="000000"):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = pBdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pBdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), str(color).replace("#", ""))


def apply_hex_font_color(run, hex_color):
    try:
        value = str(hex_color or "").replace("#", "").upper()
        if len(value) == 6:
            run.font.color.rgb = RGBColor.from_string(value)
    except Exception:
        pass


def detect_field_type(label, line, vertical_info):
    label_lower = label.lower()

    if ("ja" in label_lower and "nein" in label_lower) or (
        "yes" in label_lower and "no" in label_lower
    ):
        return "choice"

    length = float(line.get("length", 0))

    if vertical_info.get("left_border") and vertical_info.get("right_border"):
        return "boxed_field"

    if length < 80:
        return "short_text_field"

    return "text_field"


def cluster_form_rows(fields):
    if not fields:
        return []

    ordered = sorted(fields, key=lambda item: (item["y"], item["x0"]))
    groups = []

    for field in ordered:
        placed = False

        for group in groups:
            if abs(field["y"] - group["average_y"]) <= FORM_ROW_Y_TOLERANCE:
                group["fields"].append(field)
                group["average_y"] = (
                    sum(item["y"] for item in group["fields"])
                    / len(group["fields"])
                )
                placed = True
                break

        if not placed:
            groups.append({
                "average_y": field["y"],
                "fields": [field],
            })

    result = []

    for index, group in enumerate(groups, start=1):
        group_fields = sorted(group["fields"], key=lambda item: item["x0"])

        result.append({
            "row_id": f"form_row_{index}",
            "y": round_num(group["average_y"]),
            "field_count": len(group_fields),
            "field_ids": [field["field_id"] for field in group_fields],
        })

    return result


def cluster_form_columns(fields):
    if not fields:
        return []

    groups = []

    for field in sorted(fields, key=lambda item: item["x0"]):
        placed = False

        for group in groups:
            if abs(field["x0"] - group["average_x"]) <= FORM_COLUMN_X_TOLERANCE:
                group["fields"].append(field)
                group["average_x"] = (
                    sum(item["x0"] for item in group["fields"])
                    / len(group["fields"])
                )
                placed = True
                break

        if not placed:
            groups.append({
                "average_x": field["x0"],
                "fields": [field],
            })

    result = []

    for index, group in enumerate(groups, start=1):
        result.append({
            "column_id": f"form_column_{index}",
            "x": round_num(group["average_x"]),
            "field_count": len(group["fields"]),
            "field_ids": [field["field_id"] for field in group["fields"]],
        })

    return result


def detect_form_structure(form_lines, vertical_lines, raw_blocks):
    if not ENABLE_FORM_DETECTION:
        return {"fields": [], "rows": [], "columns": []}

    fields = []
    field_counter = 1

    ordered_lines = sorted(
        form_lines,
        key=lambda line: (
            float(line.get("y0", 0)),
            float(line.get("x0", 0)),
        ),
    )

    for line in ordered_lines:
        x0 = min(float(line.get("x0", 0)), float(line.get("x1", 0)))
        x1 = max(float(line.get("x0", 0)), float(line.get("x1", 0)))
        y = float(line.get("y0", 0))
        width = x1 - x0

        if width < FORM_MIN_FIELD_WIDTH:
            continue

        label_info = find_form_label(line, raw_blocks)
        vertical_info = find_vertical_boundaries(line, vertical_lines)

        field = {
            "field_id": f"form_{field_counter}",
            "type": detect_field_type(
                label_info["text"],
                line,
                vertical_info,
            ),
            "x0": round_num(x0),
            "y": round_num(y),
            "x1": round_num(x1),
            "width": round_num(width),
            "line_id": line.get("id"),
            "line_source": line.get("source"),
            "label": label_info["text"],
            "label_bbox": label_info["bbox"],
            "label_line_index": label_info.get("line_index"),
            "label_source_block_bbox": label_info.get("source_block_bbox"),
            "label_color": label_info.get("color", "#000000"),
            "label_position": classify_label_position(label_info.get("bbox"), y),
            "left_border": vertical_info["left_border"],
            "right_border": vertical_info["right_border"],
            "internal_borders": vertical_info["internal_borders"],
        }

        fields.append(field)
        field_counter += 1

    rows = cluster_form_rows(fields)
    columns = cluster_form_columns(fields)

    field_lookup = {field["field_id"]: field for field in fields}

    for row in rows:
        for field_id in row["field_ids"]:
            if field_id in field_lookup:
                field_lookup[field_id]["row_id"] = row["row_id"]

    for column in columns:
        for field_id in column["field_ids"]:
            if field_id in field_lookup:
                field_lookup[field_id]["column_id"] = column["column_id"]

    return {
        "fields": fields,
        "rows": rows,
        "columns": columns,
    }

# ============================================================
# TABLE DETECTION
# ============================================================

def find_page_tables(page):
    if not ENABLE_TABLE_DETECTION:
        return []

    try:
        finder = page.find_tables()
        if finder is None:
            return []
        return list(finder.tables)
    except Exception as e:
        print("Table detection error:", str(e))
        return []


def clean_table_matrix(matrix):
    if not matrix:
        return []

    cleaned = []
    max_columns = 0

    for row in matrix:
        if row is None:
            row = []

        clean_row = [clean_text(value) for value in row]
        max_columns = max(max_columns, len(clean_row))
        cleaned.append(clean_row)

    if max_columns == 0:
        return []

    for row in cleaned:
        while len(row) < max_columns:
            row.append("")

    return [
        row
        for row in cleaned
        if any(value.strip() for value in row)
    ]


def extract_table_cell_bboxes(detected_table):
    result = []

    try:
        for row in detected_table.rows:
            row_result = []

            for cell in row.cells:
                if cell is None:
                    row_result.append(None)
                else:
                    row_result.append(bbox_dict(cell))

            result.append(row_result)

    except Exception:
        pass

    return result

# ============================================================
# TEXT EXTRACTION
# ============================================================

def extract_raw_text_blocks(page):
    try:
        text_dict = page.get_text(
            "dict",
            flags=fitz.TEXT_PRESERVE_WHITESPACE,
        )
    except Exception:
        text_dict = page.get_text("dict")

    blocks_out = []

    for raw_block in text_dict.get("blocks", []):
        if raw_block.get("type") != 0:
            continue

        block_bbox = raw_block.get("bbox", (0, 0, 0, 0))
        lines_out = []
        all_spans = []
        text_lines = []

        for raw_line in raw_block.get("lines", []):
            line_spans = []

            for raw_span in raw_line.get("spans", []):
                span = normalize_span(raw_span)

                if not span["text"]:
                    continue

                line_spans.append(span)
                all_spans.append(span)

            if not line_spans:
                continue

            line_text = "".join(
                span["text"] for span in line_spans
            ).strip()

            if not line_text:
                continue

            line_bbox = raw_line.get("bbox", block_bbox)

            lines_out.append({
                "text": line_text,
                "bbox": bbox_dict(line_bbox),
                "spans": line_spans,
            })

            text_lines.append(line_text)

        text = "\n".join(text_lines).strip()

        if not text:
            continue

        blocks_out.append({
            "text": text,
            "bbox": bbox_dict(block_bbox),
            "lines": lines_out,
            "spans": all_spans,
            "style": get_block_style(all_spans),
        })

    return blocks_out

# ============================================================
# FONT STATS / HEADINGS
# ============================================================

def get_page_font_statistics(blocks):
    sizes = []

    for block in blocks:
        for span in block.get("spans", []):
            size = float(span.get("size", 0))
            if size > 0:
                sizes.append(size)

    if not sizes:
        return {
            "median_size": DEFAULT_FONT_SIZE,
            "max_size": DEFAULT_FONT_SIZE,
        }

    return {
        "median_size": round_num(statistics.median(sizes)),
        "max_size": round_num(max(sizes)),
    }


def classify_text_role(block, font_stats):
    text = clean_text(block.get("text", ""))
    style = block.get("style", {})

    median_size = max(
        1,
        float(
            font_stats.get(
                "median_size",
                DEFAULT_FONT_SIZE,
            )
        ),
    )

    max_size = float(style.get("max_size", 0))
    bold = bool(style.get("bold", False))
    line_count = max(1, len(block.get("lines", [])))
    short = len(text) <= 120

    if short and max_size >= median_size * HEADING_LARGE_RATIO:
        return "heading_1"

    if (
        short
        and bold
        and max_size >= median_size * HEADING_SIZE_RATIO
    ):
        return "heading_2"

    if short and bold and line_count <= 2:
        return "subheading"

    return "paragraph"

# ============================================================
# COLUMN DETECTION
# ============================================================

def detect_columns(blocks, page_width):
    if not ENABLE_COLUMN_DETECTION:
        return {"column_count": 1, "divider_x": None}

    candidates = []

    for block in blocks:
        bbox = block.get("bbox", {})
        x0 = float(bbox.get("x0", 0))
        x1 = float(bbox.get("x1", 0))
        width = x1 - x0

        if 10 < width < page_width * 0.70:
            candidates.append((x0, x1))

    if len(candidates) < COLUMN_MIN_BLOCKS * 2:
        return {"column_count": 1, "divider_x": None}

    center = page_width / 2

    left = [item for item in candidates if item[0] < center]
    right = [
        item
        for item in candidates
        if item[0] >= center * 0.85
    ]

    if (
        len(left) < COLUMN_MIN_BLOCKS
        or len(right) < COLUMN_MIN_BLOCKS
    ):
        return {"column_count": 1, "divider_x": None}

    left_edge = statistics.median(item[1] for item in left)
    right_edge = statistics.median(item[0] for item in right)
    gap = right_edge - left_edge

    if gap >= page_width * COLUMN_GAP_RATIO:
        return {
            "column_count": 2,
            "divider_x": round_num(
                (left_edge + right_edge) / 2
            ),
        }

    return {"column_count": 1, "divider_x": None}

# ============================================================
# UNDERLINE DETECTION
# ============================================================

def detect_span_underlines(spans, horizontal_lines):
    for span in spans:
        span["underline"] = False

        bbox = span.get("bbox", {})
        x0 = float(bbox.get("x0", 0))
        y1 = float(bbox.get("y1", 0))
        x1 = float(bbox.get("x1", 0))
        text_width = max(1.0, x1 - x0)

        for line in horizontal_lines:
            ly = float(line.get("y0", 0))
            lx0 = min(
                float(line.get("x0", 0)),
                float(line.get("x1", 0)),
            )
            lx1 = max(
                float(line.get("x0", 0)),
                float(line.get("x1", 0)),
            )

            vertical_distance = abs(ly - y1)
            overlap = max(
                0.0,
                min(x1, lx1) - max(x0, lx0),
            )

            if (
                vertical_distance <= 3.0
                and (overlap / text_width) >= 0.50
            ):
                span["underline"] = True
                break


# ============================================================
# FOOTER DETECTION
# ============================================================

def normalize_footer_signature(text):
    text = clean_text(text).lower()
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_footer_relaxed_signature(text):
    text = normalize_footer_signature(text)

    # Page/date style footer text can differ only by page number/date.
    if FOOTER_PAGE_MARKER_REGEX.search(text):
        text = re.sub(r"\d+", "#", text)

    return text


def build_footer_signature_sets(document):
    """
    Find repeated lower-page blocks across the PDF.

    The first implementation only scanned the bottom ~18% of a page.
    Many PDFs place company address / phone / bank information noticeably
    higher, so we now scan from 60% of page height downward and also keep
    the typical relative Y position for every repeated signature.
    """
    if not ENABLE_FOOTER_DETECTION:
        return {}, {}

    exact_occurrences = {}
    relaxed_occurrences = {}

    for page_index in range(len(document)):
        page = document[page_index]
        page_height = max(1.0, float(page.rect.height))
        raw_blocks = extract_raw_text_blocks(page)

        for block in raw_blocks:
            bbox = block.get("bbox", {})
            y0 = float(bbox.get("y0", 0))
            rel_y = y0 / page_height

            # Only consider the lower part of the page as a footer candidate.
            if rel_y < FOOTER_SEARCH_REGION_START_RATIO:
                continue

            text = clean_text(block.get("text", ""))
            if not text:
                continue

            exact = normalize_footer_signature(text)
            relaxed = normalize_footer_relaxed_signature(text)

            if exact:
                exact_occurrences.setdefault(exact, []).append(
                    (page_index, rel_y)
                )

            if relaxed:
                relaxed_occurrences.setdefault(relaxed, []).append(
                    (page_index, rel_y)
                )

    def build_repeated_map(occurrences):
        repeated = {}

        for signature, items in occurrences.items():
            page_ids = {page_index for page_index, _ in items}

            if len(page_ids) < FOOTER_REPEAT_MIN_PAGES:
                continue

            positions = sorted(rel_y for _, rel_y in items)
            typical_y = statistics.median(positions)

            repeated[signature] = {
                "page_count": len(page_ids),
                "typical_y": typical_y,
            }

        return repeated

    return (
        build_repeated_map(exact_occurrences),
        build_repeated_map(relaxed_occurrences),
    )


def block_is_footer(
    block,
    page_height,
    exact_footer_signatures,
    relaxed_footer_signatures,
):
    if not ENABLE_FOOTER_DETECTION:
        return False

    bbox = block.get("bbox", {})
    y0 = float(bbox.get("y0", 0))
    y1 = float(bbox.get("y1", y0))
    page_height = max(1.0, float(page_height))
    rel_y = y0 / page_height

    # Keep true body content out of the footer classifier.
    if rel_y < FOOTER_REGION_START_RATIO:
        return False

    text = clean_text(block.get("text", ""))
    if not text:
        return False

    exact = normalize_footer_signature(text)
    relaxed = normalize_footer_relaxed_signature(text)

    def repeated_at_similar_position(signature, repeated_map):
        info = repeated_map.get(signature)
        if not info:
            return False

        typical_y = float(info.get("typical_y", rel_y))
        return abs(rel_y - typical_y) <= FOOTER_POSITION_TOLERANCE_RATIO

    # Repeated company/contact/footer blocks are accepted even when they are
    # not extremely close to the physical bottom of the page.
    if repeated_at_similar_position(exact, exact_footer_signatures):
        return True

    if repeated_at_similar_position(relaxed, relaxed_footer_signatures):
        return True

    # Very bottom page markers remain footers even if unique.
    hard_footer_zone = (
        y0 >= page_height * FOOTER_HARD_REGION_START_RATIO
        or y1 >= page_height * 0.97
    )

    if hard_footer_zone and FOOTER_PAGE_MARKER_REGEX.search(text):
        return True

    return False

# ============================================================
# HEADER / LETTERHEAD DETECTION
# ============================================================

def normalize_header_signature(text):
    text = clean_text(text).lower()
    text = re.sub(r"\s+", " ", text).strip()
    return text


def build_header_signature_sets(document):
    if not ENABLE_HEADER_DETECTION:
        return {}

    occurrences = {}

    for page_index in range(len(document)):
        page = document[page_index]
        page_height = max(1.0, float(page.rect.height))
        raw_blocks = extract_raw_text_blocks(page)

        for block in raw_blocks:
            bbox = block.get("bbox", {})
            y0 = float(bbox.get("y0", 0))
            rel_y = y0 / page_height

            if rel_y > HEADER_REGION_END_RATIO:
                continue

            text = clean_text(block.get("text", ""))
            if not text:
                continue

            sig = normalize_header_signature(text)
            if sig:
                occurrences.setdefault(sig, []).append((page_index, rel_y))

    repeated = {}
    for signature, items in occurrences.items():
        page_ids = {page_index for page_index, _ in items}
        if len(page_ids) < HEADER_REPEAT_MIN_PAGES:
            continue
        repeated[signature] = {
            "page_count": len(page_ids),
            "typical_y": statistics.median(rel_y for _, rel_y in items),
        }

    return repeated


def block_is_header(block, page_height, repeated_header_signatures):
    if not ENABLE_HEADER_DETECTION:
        return False

    bbox = block.get("bbox", {})
    y0 = float(bbox.get("y0", 0))
    page_height = max(1.0, float(page_height))
    rel_y = y0 / page_height

    if rel_y > HEADER_REGION_END_RATIO:
        return False

    text = clean_text(block.get("text", ""))
    if not text:
        return False

    sig = normalize_header_signature(text)
    info = repeated_header_signatures.get(sig)
    if info:
        typical_y = float(info.get("typical_y", rel_y))
        if abs(rel_y - typical_y) <= HEADER_POSITION_TOLERANCE_RATIO:
            return True

    # Unique first-page letterheads are common, but they must be compact.
    # A large body block near the top may contain an e-mail or URL; placing
    # such a block into a Word header makes Word repeat it on every overflow
    # page. Guard against that by limiting block height, text size and lines.
    y1 = float(bbox.get("y1", y0))
    block_height_ratio = max(0.0, y1 - y0) / page_height
    line_count = len(block.get("lines", []) or [])

    if (
        HEADER_CONTACT_REGEX.search(text)
        and not HEADER_FORM_LABEL_REGEX.search(text)
        and block_height_ratio <= HEADER_UNIQUE_MAX_HEIGHT_RATIO
        and len(text) <= HEADER_UNIQUE_MAX_TEXT_LENGTH
        and line_count <= HEADER_UNIQUE_MAX_LINES
    ):
        return True

    return False


def bbox_similarity_score(a, b):
    """Return a geometry-only similarity score for two bbox dictionaries."""
    if not a or not b:
        return 0.0
    ra = (
        float(a.get("x0", 0)), float(a.get("y0", 0)),
        float(a.get("x1", 0)), float(a.get("y1", 0)),
    )
    rb = (
        float(b.get("x0", 0)), float(b.get("y0", 0)),
        float(b.get("x1", 0)), float(b.get("y1", 0)),
    )
    overlap = bbox_overlap_ratio(ra, rb)
    _, containment = _rect_metrics(ra, rb)
    return max(overlap, containment)


def attach_form_label_translation_elements(
    form_fields,
    raw_blocks,
    elements,
    element_counter,
    page_number,
):
    """Give every reconstructed form-label line its own translation id.

    PDF text blocks often contain several labels in one block.  Asking the
    translator to preserve the original internal line breaks is fragile: a
    six-line block can come back as four or five translated lines, causing
    labels to be assigned to the wrong form fields.  Here each visual label
    line becomes an independent translation element.  This is generic and is
    driven only by bbox geometry.

    Small helper lines directly underneath a primary label and inside the same
    field column are attached to that field as well (for example explanatory
    parenthetical text beneath a date field).
    """
    line_records = []
    for block_index, block in enumerate(raw_blocks):
        for line_index, line in enumerate(block.get("lines", []) or []):
            text = clean_text(line.get("text", ""))
            if not text:
                text = clean_text("".join(
                    span.get("text", "")
                    for span in (line.get("spans", []) or [])
                ))
            if not text:
                continue
            line_records.append({
                "block_index": block_index,
                "line_index": line_index,
                "line": line,
                "text": text,
                "bbox": line.get("bbox") or block.get("bbox", {}),
                "style": block.get("style", {}),
            })

    id_cache = {}

    def translation_id_for(record, role):
        nonlocal element_counter
        bbox = record.get("bbox", {})
        key = (
            page_number,
            round(float(bbox.get("x0", 0)), 2),
            round(float(bbox.get("y0", 0)), 2),
            round(float(bbox.get("x1", 0)), 2),
            round(float(bbox.get("y1", 0)), 2),
            record.get("text", ""),
        )
        if key in id_cache:
            return id_cache[key]

        element_id = f"pdf_{element_counter}"
        element_counter += 1
        elements.append({
            "id": element_id,
            "page": page_number,
            "type": "form_label_line",
            "text": record.get("text", ""),
            "bbox": bbox,
            "role": role,
            "style": record.get("style", {}),
            "lines": [record.get("line", {})],
            "spans": (record.get("line", {}) or {}).get("spans", []),
            "list": None,
        })
        id_cache[key] = element_id
        return element_id

    for field in form_fields:
        label_bbox = field.get("label_bbox")
        if not label_bbox:
            continue

        primary = None
        primary_score = 0.0
        for record in line_records:
            score = bbox_similarity_score(record.get("bbox"), label_bbox)
            if score > primary_score:
                primary_score = score
                primary = record

        if primary is None or primary_score < 0.70:
            continue

        field["label_translation_id"] = translation_id_for(
            primary, "form_field_label"
        )
        field["label_translation_source"] = primary.get("text", "")

        # Find explanatory/helper lines immediately below the primary label,
        # within the same field's horizontal span.  We do not rely on words or
        # language, only local geometry and visual hierarchy.
        px0 = float(field.get("x0", 0))
        px1 = float(field.get("x1", 0))
        pb = primary.get("bbox", {})
        py1 = float(pb.get("y1", 0))
        primary_block = primary.get("block_index")

        helpers = []
        for record in line_records:
            if record is primary:
                continue
            if record.get("block_index") != primary_block:
                continue

            bbox = record.get("bbox", {})
            by0 = float(bbox.get("y0", 0))
            bx0 = float(bbox.get("x0", 0))
            bx1 = float(bbox.get("x1", 0))

            vertical_gap = by0 - py1
            if vertical_gap < -2.0 or vertical_gap > 4.5:
                continue

            overlap = horizontal_overlap(px0 - 4.0, px1 + 4.0, bx0, bx1)
            line_width = max(1.0, bx1 - bx0)
            if overlap / line_width < 0.72:
                continue

            # Section headings are structural text and must remain separate.
            if line_is_form_section_heading(record.get("line", {})):
                continue

            # A helper normally starts close to the same left edge as its
            # primary label.  This prevents a neighbouring column's label from
            # being attached to the current field.
            if abs(bx0 - float(pb.get("x0", bx0))) > 12.0:
                continue

            helpers.append(record)

        helpers.sort(key=lambda r: (
            float(r.get("bbox", {}).get("y0", 0)),
            float(r.get("bbox", {}).get("x0", 0)),
        ))

        field["label_helper_translation_ids"] = [
            translation_id_for(record, "form_field_helper")
            for record in helpers
        ]
        field["label_helper_sources"] = [
            record.get("text", "") for record in helpers
        ]
        # Keep the exact helper rectangles as claimed form text as well.
        # Without this, the same source helper line can be rendered twice:
        # once inside the form field and again by the generic unclaimed-text
        # fallback. This showed up as duplicate translations of the same
        # parenthetical helper text.
        field["label_helper_bboxes"] = [
            record.get("bbox", {}) for record in helpers
        ]

    return element_counter


# ============================================================
# MAIN EXTRACTION
# ============================================================

def extract_pdf_structure(document):
    pages = []
    elements = []

    exact_footer_signatures, relaxed_footer_signatures = (
        build_footer_signature_sets(document)
    )
    repeated_header_signatures = build_header_signature_sets(document)

    element_counter = 1
    table_counter = 1

    for page_index in range(len(document)):
        page = document[page_index]
        page_number = page_index + 1
        page_width = float(page.rect.width)
        page_height = float(page.rect.height)

        drawings, lines, rectangles = normalize_drawings(page)

        horizontal_lines = [
            line
            for line in lines
            if line.get("orientation") == "horizontal"
        ]

        vertical_lines = [
            line
            for line in lines
            if line.get("orientation") == "vertical"
        ]

        form_lines = detect_form_lines(
            horizontal_lines,
            page_width,
        )

        raw_blocks = extract_raw_text_blocks(page)
        raw_blocks, suppressed_spatial_duplicates = suppress_spatial_duplicate_blocks(
            raw_blocks
        )

        for block in raw_blocks:
            detect_span_underlines(
                block.get("spans", []),
                horizontal_lines,
            )

            for line_data in block.get("lines", []):
                detect_span_underlines(
                    line_data.get("spans", []),
                    horizontal_lines,
                )

        font_stats = get_page_font_statistics(raw_blocks)

        form_structure = detect_form_structure(
            form_lines,
            vertical_lines,
            raw_blocks,
        )

        form_fields = form_structure["fields"]
        form_rows = form_structure["rows"]
        form_columns = form_structure["columns"]

        detected_tables = find_page_tables(page)
        table_regions = []
        table_structures = []

        for detected_table in detected_tables:
            try:
                bbox = detected_table.bbox
                table_bbox = (
                    float(bbox[0]),
                    float(bbox[1]),
                    float(bbox[2]),
                    float(bbox[3]),
                )
                matrix = clean_table_matrix(
                    detected_table.extract()
                )
            except Exception as e:
                print("Could not parse table:", str(e))
                continue

            if not matrix:
                continue

            row_count = len(matrix)
            column_count = max(len(row) for row in matrix)

            if row_count <= 0 or column_count <= 0:
                continue

            table_id = f"table_{table_counter}"
            table_counter += 1

            cell_bboxes = extract_table_cell_bboxes(
                detected_table
            )

            table_data = {
                "table_id": table_id,
                "type": "table",
                "page": page_number,
                "bbox": bbox_dict(table_bbox),
                "row_count": row_count,
                "column_count": column_count,
                "rows": [],
            }

            for row_index in range(row_count):
                row_data = []

                for col_index in range(column_count):
                    value = ""

                    if (
                        row_index < len(matrix)
                        and col_index < len(matrix[row_index])
                    ):
                        value = clean_text(
                            matrix[row_index][col_index]
                        )

                    cell_bbox = None

                    if (
                        row_index < len(cell_bboxes)
                        and col_index < len(cell_bboxes[row_index])
                    ):
                        cell_bbox = cell_bboxes[row_index][col_index]

                    if not cell_bbox:
                        cell_bbox = bbox_dict(table_bbox)

                    if value:
                        element_id = f"pdf_{element_counter}"
                        element_counter += 1

                        element = {
                            "id": element_id,
                            "page": page_number,
                            "type": "table_cell",
                            "table_id": table_id,
                            "row": row_index,
                            "col": col_index,
                            "text": value,
                            "bbox": cell_bbox,
                        }

                        elements.append(element)
                    else:
                        element_id = None

                    row_data.append({
                        "id": element_id,
                        "row": row_index,
                        "col": col_index,
                        "text": value,
                        "bbox": cell_bbox,
                    })

                table_data["rows"].append(row_data)

            table_regions.append(table_bbox)
            table_structures.append(table_data)

        text_elements = []
        header_elements = []
        footer_elements = []
        detected_lists = []

        for block in raw_blocks:
            bbox = block.get("bbox", {})

            block_rect = (
                float(bbox.get("x0", 0)),
                float(bbox.get("y0", 0)),
                float(bbox.get("x1", 0)),
                float(bbox.get("y1", 0)),
            )

            inside_table = any(
                rect_center_inside(
                    block_rect,
                    table_rect,
                )
                for table_rect in table_regions
            )

            if inside_table:
                continue

            text = clean_text(block.get("text", ""))
            if not text:
                continue

            is_footer = block_is_footer(
                block,
                page_height,
                exact_footer_signatures,
                relaxed_footer_signatures,
            )
            is_header = (
                not is_footer
                and block_is_header(
                    block,
                    page_height,
                    repeated_header_signatures,
                )
            )

            element_id = f"pdf_{element_counter}"
            element_counter += 1

            if is_footer:
                element = {
                    "id": element_id,
                    "page": page_number,
                    "type": "footer",
                    "text": text,
                    "bbox": block["bbox"],
                    "role": "footer",
                    "style": block.get("style", {}),
                    "lines": block.get("lines", []),
                    "spans": block.get("spans", []),
                    "list": None,
                }

                elements.append(element)
                footer_elements.append(element)
                continue

            if is_header:
                element = {
                    "id": element_id,
                    "page": page_number,
                    "type": "header",
                    "text": text,
                    "bbox": block["bbox"],
                    "role": "header",
                    "style": block.get("style", {}),
                    "lines": block.get("lines", []),
                    "spans": block.get("spans", []),
                    "list": None,
                }
                elements.append(element)
                header_elements.append(element)
                continue

            role = classify_text_role(
                block,
                font_stats,
            )

            list_info = detect_list_info(text)
            element_type = (
                "list_item"
                if list_info
                else "text_block"
            )

            element = {
                "id": element_id,
                "page": page_number,
                "type": element_type,
                "text": text,
                "bbox": block["bbox"],
                "role": role,
                "style": block.get("style", {}),
                "lines": block.get("lines", []),
                "spans": block.get("spans", []),
                "list": list_info,
            }

            elements.append(element)
            text_elements.append(element)

            if list_info:
                detected_lists.append({
                    "id": element_id,
                    "bbox": block["bbox"],
                    "list_type": list_info["list_type"],
                    "marker": list_info["marker"],
                    "text": text,
                })

        # Create independent translation ids for visual form-label lines.
        # They are not rendered as normal text elements; add_form_row() uses
        # them directly so translation line wrapping cannot scramble fields.
        if (
            len(form_fields) >= FORM_PAGE_MIN_FIELDS
            and len(form_rows) >= FORM_PAGE_MIN_ROWS
            and len(form_columns) >= 2
        ):
            element_counter = attach_form_label_translation_elements(
                form_fields,
                raw_blocks,
                elements,
                element_counter,
                page_number,
            )

        columns = detect_columns(
            raw_blocks,
            page_width,
        )

        layout_items = []

        for element in text_elements:
            layout_items.append({
                "type": element["type"],
                "id": element["id"],
                "x0": element["bbox"]["x0"],
                "y0": element["bbox"]["y0"],
            })

        for table in table_structures:
            layout_items.append({
                "type": "table",
                "table_id": table["table_id"],
                "x0": table["bbox"]["x0"],
                "y0": table["bbox"]["y0"],
            })

        layout_items.sort(
            key=lambda item: (
                float(item.get("y0", 0)),
                float(item.get("x0", 0)),
            )
        )

        pages.append({
            "page": page_number,
            "width": round_num(page_width),
            "height": round_num(page_height),
            "font_statistics": font_stats,
            "column_count": columns["column_count"],
            "column_divider_x": columns["divider_x"],
            "text_block_count": len(text_elements),
            "spatial_duplicate_count": len(suppressed_spatial_duplicates),
            "header_block_count": len(header_elements),
            "footer_block_count": len(footer_elements),
            "table_count": len(table_structures),
            "list_count": len(detected_lists),
            "drawing_count": len(drawings),
            "line_count": len(lines),
            "horizontal_line_count": len(horizontal_lines),
            "vertical_line_count": len(vertical_lines),
            "rectangle_count": len(rectangles),
            "form_line_count": len(form_lines),
            "form_field_count": len(form_fields),
            "form_row_count": len(form_rows),
            "form_column_count": len(form_columns),
            "tables": table_structures,
            "lists": detected_lists,
            "drawings": drawings,
            "lines": lines,
            "rectangles": rectangles,
            "form_lines": form_lines,
            "form_fields": form_fields,
            "form_rows": form_rows,
            "form_columns": form_columns,
            "header_blocks": header_elements,
            "footer_blocks": footer_elements,
            "layout": layout_items,
        })

    elements.sort(
        key=lambda item: int(item["id"].split("_")[1])
    )

    translation_elements = [
        {
            "id": element["id"],
            "type": element["type"],
            "text": element["text"],
        }
        for element in elements
    ]

    chunks = [
        translation_elements[i:i + CHUNK_SIZE]
        for i in range(
            0,
            len(translation_elements),
            CHUNK_SIZE,
        )
    ]

    return pages, elements, chunks

# ============================================================
# TRANSLATION JSON
# ============================================================

def parse_translations(raw_value):
    if raw_value is None:
        return {}

    if isinstance(raw_value, bytes):
        raw_value = raw_value.decode("utf-8")

    if isinstance(raw_value, str):
        data = json.loads(raw_value)
    else:
        data = raw_value

    result = {}

    for item in data.get("translations", []):
        element_id = item.get("id")
        if not element_id:
            continue

        text = item.get("text")
        if text is None:
            text = ""

        result[str(element_id)] = str(text)

    return result

# ============================================================
# WORD XML HELPERS
# ============================================================

def configure_section(section, page_width, page_height):
    section.page_width = Pt(page_width)
    section.page_height = Pt(page_height)

    section.top_margin = Pt(0)
    section.bottom_margin = Pt(0)
    section.left_margin = Pt(0)
    section.right_margin = Pt(0)

    section.header_distance = Pt(0)
    section.footer_distance = Pt(0)


def prevent_row_split(row):
    try:
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = tr_pr.find(qn("w:cantSplit"))

        if cant_split is None:
            cant_split = OxmlElement("w:cantSplit")
            tr_pr.append(cant_split)

        cant_split.set(qn("w:val"), "1")
    except Exception:
        pass


def set_cell_margins(
    cell,
    top=20,
    start=30,
    bottom=20,
    end=30,
):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()

    tc_mar = tc_pr.first_child_found_in(
        "w:tcMar"
    )

    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)

    for name, value in [
        ("top", top),
        ("start", start),
        ("bottom", bottom),
        ("end", end),
    ]:
        node = tc_mar.find(
            qn(f"w:{name}")
        )

        if node is None:
            node = OxmlElement(
                f"w:{name}"
            )
            tc_mar.append(node)

        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_border(
    cell,
    top=None,
    left=None,
    bottom=None,
    right=None,
):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in(
        "w:tcBorders"
    )

    if borders is None:
        borders = OxmlElement(
            "w:tcBorders"
        )
        tc_pr.append(borders)

    sides = {
        "top": top,
        "left": left,
        "bottom": bottom,
        "right": right,
    }

    for edge, spec in sides.items():
        tag = f"w:{edge}"
        node = borders.find(qn(tag))

        if node is None:
            node = OxmlElement(tag)
            borders.append(node)

        if spec is None:
            node.set(qn("w:val"), "nil")
            continue

        node.set(
            qn("w:val"),
            spec.get("val", "single"),
        )
        node.set(
            qn("w:sz"),
            str(spec.get("sz", 4)),
        )
        node.set(
            qn("w:space"),
            str(spec.get("space", 0)),
        )
        node.set(
            qn("w:color"),
            spec.get("color", "000000"),
        )


def set_table_borders(table):
    try:
        tbl_pr = table._tbl.tblPr

        borders = (
            tbl_pr.first_child_found_in(
                "w:tblBorders"
            )
        )

        if borders is None:
            borders = OxmlElement(
                "w:tblBorders"
            )
            tbl_pr.append(borders)

        for edge in [
            "top",
            "left",
            "bottom",
            "right",
            "insideH",
            "insideV",
        ]:
            tag = f"w:{edge}"
            border = borders.find(
                qn(tag)
            )

            if border is None:
                border = OxmlElement(tag)
                borders.append(border)

            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "808080")

    except Exception:
        pass


def set_table_indentation(table, left_pt):
    try:
        tbl_pr = table._tbl.tblPr
        tbl_ind = tbl_pr.first_child_found_in(
            "w:tblInd"
        )

        if tbl_ind is None:
            tbl_ind = OxmlElement(
                "w:tblInd"
            )
            tbl_pr.append(tbl_ind)

        tbl_ind.set(
            qn("w:w"),
            str(int(max(0, left_pt) * 20)),
        )
        tbl_ind.set(
            qn("w:type"),
            "dxa",
        )
    except Exception:
        pass


def set_table_fixed_layout(table):
    try:
        tbl_pr = table._tbl.tblPr
        tbl_layout = tbl_pr.first_child_found_in(
            "w:tblLayout"
        )

        if tbl_layout is None:
            tbl_layout = OxmlElement(
                "w:tblLayout"
            )
            tbl_pr.append(tbl_layout)

        tbl_layout.set(
            qn("w:type"),
            "fixed",
        )
    except Exception:
        pass


def set_cell_width(cell, width_pt):
    try:
        width_pt = max(
            FORM_DOCX_MIN_CELL_WIDTH_PT,
            float(width_pt),
        )

        tc_pr = cell._tc.get_or_add_tcPr()
        tc_w = tc_pr.first_child_found_in(
            "w:tcW"
        )

        if tc_w is None:
            tc_w = OxmlElement("w:tcW")
            tc_pr.append(tc_w)

        tc_w.set(
            qn("w:w"),
            str(int(width_pt * 20)),
        )
        tc_w.set(
            qn("w:type"),
            "dxa",
        )

        cell.width = Pt(width_pt)
    except Exception:
        pass


def remove_table_borders(table):
    try:
        tbl_pr = table._tbl.tblPr
        borders = tbl_pr.first_child_found_in(
            "w:tblBorders"
        )

        if borders is None:
            borders = OxmlElement(
                "w:tblBorders"
            )
            tbl_pr.append(borders)

        for edge in [
            "top",
            "left",
            "bottom",
            "right",
            "insideH",
            "insideV",
        ]:
            tag = f"w:{edge}"
            border = borders.find(
                qn(tag)
            )

            if border is None:
                border = OxmlElement(tag)
                borders.append(border)

            border.set(qn("w:val"), "nil")

    except Exception:
        pass

# ============================================================
# BASIC TEXT / TABLE DOCX
# ============================================================

def add_basic_text_block(
    document,
    element,
    translated_text,
    page_width,
    previous_bottom,
):
    bbox = element["bbox"]

    x0 = float(bbox["x0"])
    y0 = float(bbox["y0"])
    x1 = float(bbox["x1"])
    y1 = float(bbox["y1"])

    paragraph = document.add_paragraph()

    paragraph.paragraph_format.left_indent = Pt(
        max(0, x0)
    )

    paragraph.paragraph_format.right_indent = Pt(
        max(0, page_width - x1)
    )

    gap = y0 if previous_bottom is None else y0 - previous_bottom

    gap = clamp(
        gap,
        MIN_VERTICAL_GAP_PT,
        MAX_VERTICAL_GAP_PT,
    )

    paragraph.paragraph_format.space_before = Pt(gap)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(translated_text)

    style = element.get("style", {})

    run.font.name = (
        style.get("font")
        or DEFAULT_FONT_NAME
    )

    run.font.size = Pt(
        float(
            style.get(
                "size",
                DEFAULT_FONT_SIZE,
            )
            or DEFAULT_FONT_SIZE
        )
    )

    run.bold = bool(
        style.get(
            "bold",
            False,
        )
    )

    run.italic = bool(
        style.get(
            "italic",
            False,
        )
    )
    # Preserve source text color (e.g. colored form subsection headings).
    apply_hex_font_color(run, style.get("color", "#000000"))

    return y1


def add_basic_table(
    document,
    table_data,
    translations,
):
    rows = int(
        table_data.get(
            "row_count",
            0,
        )
    )

    cols = int(
        table_data.get(
            "column_count",
            0,
        )
    )

    if rows <= 0 or cols <= 0:
        return

    table = document.add_table(
        rows=rows,
        cols=cols,
    )

    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    set_table_fixed_layout(table)
    set_table_borders(table)

    rows_data = table_data.get(
        "rows",
        [],
    )

    for row_index in range(rows):
        row = table.rows[row_index]
        prevent_row_split(row)

        for col_index in range(cols):
            cell = row.cells[col_index]
            cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

            set_cell_margins(cell)

            info = None

            if (
                row_index < len(rows_data)
                and col_index < len(rows_data[row_index])
            ):
                info = rows_data[row_index][col_index]

            if not info:
                continue

            element_id = info.get("id")
            original = info.get("text", "")

            final_text = (
                translations.get(
                    element_id,
                    original,
                )
                if element_id
                else ""
            )

            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)

            run = paragraph.add_run(final_text)
            run.font.name = DEFAULT_FONT_NAME
            run.font.size = Pt(TABLE_FONT_SIZE)

# ============================================================
# FORM DOCX HELPERS
# ============================================================

def page_is_form_page(page_data):
    """Return True only for pages with a real repeated form grid.

    A few horizontal rules on a normal multi-column/legal page must not be
    enough to switch the whole page into the form renderer.  The previous
    fallback could therefore misclassify dense text pages and explode the
    document into many broken pages.  We now require repeated multi-field
    rows, repeated columns and a reasonable amount of label evidence in
    addition to the existing field/row thresholds.
    """
    if not ENABLE_FORM_DOCX_REBUILD:
        return False
    if page_data.get("table_count", 0) != 0:
        return False

    fields = page_data.get("form_fields", []) or []
    rows = page_data.get("form_rows", []) or []
    columns = page_data.get("form_columns", []) or []

    if len(fields) < FORM_PAGE_MIN_FIELDS or len(rows) < FORM_PAGE_MIN_ROWS:
        return False

    multi_field_rows = sum(
        1 for row in rows
        if int(row.get("field_count", len(row.get("field_ids", []) or []))) >= 2
    )
    repeated_columns = sum(
        1 for column in columns
        if int(column.get("field_count", len(column.get("field_ids", []) or []))) >= 3
    )
    labeled_fields = sum(
        1 for field in fields
        if clean_text(field.get("label", ""))
    )
    label_ratio = labeled_fields / max(1, len(fields))

    return (
        multi_field_rows >= 3
        and repeated_columns >= 2
        and label_ratio >= 0.45
    )


def get_form_render_region(page_data):
    """Bounding region in which form-specific fallback is allowed.

    This deliberately derives the region from detected field geometry rather
    than from page numbers, words or fixed coordinates.  Structural headings
    immediately above a form group are included with a modest top margin,
    while unrelated text elsewhere on the page keeps the normal renderer.
    """
    fields = page_data.get("form_fields", []) or []
    if not fields:
        return None

    x0 = min(float(field.get("x0", 0)) for field in fields)
    x1 = max(float(field.get("x1", 0)) for field in fields)
    y0 = min(float(field.get("y", 0)) for field in fields)
    y1 = max(float(field.get("y", 0)) for field in fields)

    page_width = float(page_data.get("width", x1) or x1)
    page_height = float(page_data.get("height", y1 + 50) or (y1 + 50))

    return (
        max(0.0, x0 - 18.0),
        max(0.0, y0 - 72.0),
        min(page_width, x1 + 18.0),
        min(page_height, y1 + 48.0),
    )


def bbox_center_inside_region(bbox, region):
    if not bbox or not region:
        return False
    x0 = float(bbox.get("x0", 0))
    y0 = float(bbox.get("y0", 0))
    x1 = float(bbox.get("x1", x0))
    y1 = float(bbox.get("y1", y0))
    cx = (x0 + x1) / 2.0
    cy = (y0 + y1) / 2.0
    rx0, ry0, rx1, ry1 = region
    return rx0 <= cx <= rx1 and ry0 <= cy <= ry1


def get_form_label_bboxes(page_data):
    result = []

    for field in page_data.get(
        "form_fields",
        [],
    ):
        bbox = field.get("label_bbox")
        if bbox:
            result.append((
                float(bbox.get("x0", 0)),
                float(bbox.get("y0", 0)),
                float(bbox.get("x1", 0)),
                float(bbox.get("y1", 0)),
            ))

        # Helper text belongs to the field just as much as the primary label.
        # Mark it as consumed so form_element_line_events() cannot emit the
        # same visual line again as an unclaimed fallback event.
        for helper_bbox in field.get("label_helper_bboxes", []) or []:
            if not helper_bbox:
                continue
            result.append((
                float(helper_bbox.get("x0", 0)),
                float(helper_bbox.get("y0", 0)),
                float(helper_bbox.get("x1", 0)),
                float(helper_bbox.get("y1", 0)),
            ))

    return result


def rect_is_used_as_form_label(rect, label_bboxes):
    """Return True only when *this exact rectangle* belongs to a field label.

    Form PDFs often group a section heading, several field labels and small
    helper text into one PyMuPDF text block. Treating the whole block as a
    label caused unrelated lines inside that block to disappear. The matcher
    is therefore reusable at line granularity.
    """
    for label_bbox in label_bboxes:
        overlap = bbox_overlap_ratio(rect, label_bbox)
        _, containment = _rect_metrics(rect, label_bbox)
        if overlap >= 0.35 or containment >= 0.82:
            return True
    return False


def element_is_used_as_form_label(
    element,
    label_bboxes,
):
    bbox = element.get("bbox", {})
    rect = (
        float(bbox.get("x0", 0)),
        float(bbox.get("y0", 0)),
        float(bbox.get("x1", 0)),
        float(bbox.get("y1", 0)),
    )
    return rect_is_used_as_form_label(rect, label_bboxes)


def form_element_line_events(element, label_bboxes):
    """Split mixed form text blocks into safe, unclaimed line events.

    A line that is already the label of a detected writing field is omitted
    here because add_form_row() renders it next to its field. Every other line
    is kept. This creates a generic no-silent-loss rule for section headings,
    locality/place labels and helper text without relying on any language or
    page-specific wording.
    """
    lines = element.get("lines", []) or []
    if not lines:
        return []

    result = []
    matched_any_label_line = False

    for line_index, line_data in enumerate(lines):
        text = clean_text(line_data.get("text", ""))
        if not text:
            text = clean_text("".join(
                span.get("text", "")
                for span in (line_data.get("spans", []) or [])
            ))
        if not text:
            continue

        bbox = line_data.get("bbox") or element.get("bbox", {})
        rect = (
            float(bbox.get("x0", 0)),
            float(bbox.get("y0", 0)),
            float(bbox.get("x1", 0)),
            float(bbox.get("y1", 0)),
        )

        if rect_is_used_as_form_label(rect, label_bboxes):
            matched_any_label_line = True
            continue

        result.append({
            "line_index": line_index,
            "line": line_data,
            "bbox": bbox,
            "text": text,
        })

    # Only split blocks that actually contain at least one field-label line.
    # Otherwise the normal block renderer remains preferable because it
    # preserves multi-line paragraph flow better.
    if not matched_any_label_line:
        return []

    return result


def translated_form_line_text(element, line_index, fallback_text, translations):
    translated_block = clean_text(
        translations.get(element.get("id"), element.get("text", ""))
    )
    translated_lines = [
        clean_text(part)
        for part in translated_block.splitlines()
        if clean_text(part)
    ]

    if line_index is not None and 0 <= int(line_index) < len(translated_lines):
        return translated_lines[int(line_index)]

    # If the translation model merged line breaks, do not drop the source
    # line. Keeping visible source text is safer than silently removing a form
    # heading/label. Normal cases still use the translated line above.
    return clean_text(fallback_text)


def form_line_proxy_element(element, line_data, bbox):
    spans = line_data.get("spans", []) or []
    style = get_block_style(spans) if spans else dict(element.get("style", {}))
    return {
        "id": element.get("id"),
        "page": element.get("page"),
        "type": element.get("type", "text_block"),
        "role": element.get("role"),
        "bbox": bbox,
        "style": style,
        "spans": spans,
        "lines": [line_data],
    }


def add_form_heading_block(
    document,
    element,
    translated_text,
    page_width,
    previous_y,
):
    bbox = element.get(
        "bbox",
        {},
    )

    x0 = float(
        bbox.get(
            "x0",
            0,
        )
    )
    y0 = float(
        bbox.get(
            "y0",
            0,
        )
    )
    x1 = float(
        bbox.get(
            "x1",
            page_width,
        )
    )
    y1 = float(
        bbox.get(
            "y1",
            y0,
        )
    )

    paragraph = document.add_paragraph()

    gap = (
        y0
        if previous_y is None
        else y0 - previous_y
    )

    gap = clamp(
        gap,
        0,
        FORM_DOCX_ROW_GAP_MAX_PT,
    )

    paragraph.paragraph_format.space_before = Pt(
        gap
    )
    paragraph.paragraph_format.space_after = Pt(0)

    paragraph.paragraph_format.left_indent = Pt(
        max(
            0,
            x0,
        )
    )

    paragraph.paragraph_format.right_indent = Pt(
        max(
            0,
            page_width - x1,
        )
    )

    run = paragraph.add_run(
        translated_text
    )

    style = element.get(
        "style",
        {},
    )

    font_size = float(
        style.get(
            "size",
            FORM_DOCX_HEADING_FONT_PT,
        )
        or FORM_DOCX_HEADING_FONT_PT
    )

    font_size = clamp(
        font_size,
        6.5,
        12.0,
    )

    run.font.name = (
        style.get(
            "font"
        )
        or DEFAULT_FONT_NAME
    )
    run.font.size = Pt(font_size)
    run.bold = bool(
        style.get(
            "bold",
            False,
        )
    ) or element.get("role") in (
        "heading_1",
        "heading_2",
        "subheading",
    )
    run.italic = bool(
        style.get(
            "italic",
            False,
        )
    )
    apply_hex_font_color(
        run,
        style.get("color", "#000000"),
    )

    return y1



def find_best_form_label_element(field, page_elements, claimed_label_keys=None):
    """Return the source block and line index for one reconstructed label.

    Multiple form labels may live inside the same PDF text block. Therefore
    ownership is tracked per (element_id, line_index), not per whole element.
    """
    claimed_label_keys = claimed_label_keys or set()
    label_bbox = field.get("label_bbox")
    if not label_bbox:
        return None

    target = (
        float(label_bbox.get("x0", 0)),
        float(label_bbox.get("y0", 0)),
        float(label_bbox.get("x1", 0)),
        float(label_bbox.get("y1", 0)),
    )
    wanted_line_index = field.get("label_line_index")

    best = None
    best_score = -1.0

    for element in page_elements:
        element_id = element.get("id")
        if not element_id:
            continue
        if element.get("type") not in ("text_block", "list_item"):
            continue

        bbox = element.get("bbox", {})
        rect = (
            float(bbox.get("x0", 0)),
            float(bbox.get("y0", 0)),
            float(bbox.get("x1", 0)),
            float(bbox.get("y1", 0)),
        )

        overlap = bbox_overlap_ratio(rect, target)
        _, containment = _rect_metrics(rect, target)
        score = max(overlap, containment)

        if score > best_score:
            key = (str(element_id), wanted_line_index)
            if key in claimed_label_keys:
                continue
            best_score = score
            best = {
                "element": element,
                "line_index": wanted_line_index,
                "claim_key": key,
            }

    if best_score < 0.55:
        return None

    return best


def add_form_row(
    document,
    row_data,
    field_lookup,
    page_width,
    previous_y,
    page_elements,
    translations,
    rendered_ids,
    claimed_label_keys,
):
    fields = [
        field_lookup[field_id]
        for field_id in row_data.get(
            "field_ids",
            [],
        )
        if field_id in field_lookup
    ]

    if not fields:
        return previous_y

    fields = sorted(
        fields,
        key=lambda field: field["x0"],
    )

    row_y = float(
        row_data.get(
            "y",
            0,
        )
    )

    row_left = min(
        float(field["x0"])
        for field in fields
    )

    row_right = max(
        float(field["x1"])
        for field in fields
    )

    if previous_y is not None:
        gap = clamp(
            row_y - previous_y,
            0,
            FORM_DOCX_ROW_GAP_MAX_PT,
        )

        if gap > 0:
            spacer = document.add_paragraph()
            spacer.paragraph_format.space_before = Pt(
                gap
            )
            spacer.paragraph_format.space_after = Pt(0)
            spacer.paragraph_format.line_spacing = 0.5
            spacer.add_run("")

    table = document.add_table(
        rows=1,
        cols=len(fields),
    )

    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    set_table_fixed_layout(table)
    remove_table_borders(table)

    set_table_indentation(
        table,
        max(
            0,
            row_left - FORM_DOCX_LEFT_RIGHT_SAFETY_PT,
        ),
    )

    row = table.rows[0]
    prevent_row_split(row)

    for col_index, field in enumerate(fields):
        cell = row.cells[col_index]
        cell.vertical_alignment = (
            WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        )

        width = float(
            field.get(
                "width",
                50,
            )
        )

        width = min(
            width,
            max(
                FORM_DOCX_MIN_CELL_WIDTH_PT,
                page_width - row_left,
            ),
        )

        set_cell_width(
            cell,
            width,
        )

        set_cell_margins(
            cell,
            top=0,
            start=20,
            bottom=10,
            end=20,
        )

        # Prefer the independent line-level translation id created during
        # extraction.  This prevents a translator from reflowing a multi-line
        # PDF block and shifting labels to neighbouring form fields.
        label = ""
        helper_labels = []
        direct_label_id = field.get("label_translation_id")

        if direct_label_id:
            label = clean_text(translations.get(
                direct_label_id,
                field.get("label_translation_source", field.get("label", "")),
            ))
            rendered_ids.add(direct_label_id)

            helper_ids = field.get("label_helper_translation_ids", []) or []
            helper_sources = field.get("label_helper_sources", []) or []
            for helper_index, helper_id in enumerate(helper_ids):
                source = (
                    helper_sources[helper_index]
                    if helper_index < len(helper_sources)
                    else ""
                )
                helper_text = clean_text(translations.get(helper_id, source))
                if helper_text:
                    helper_labels.append(helper_text)
                rendered_ids.add(helper_id)
        else:
            label_match = find_best_form_label_element(
                field,
                page_elements,
                claimed_label_keys,
            )
            if label_match is not None:
                label_element = label_match["element"]
                label_id = label_element.get("id")
                line_index = label_match.get("line_index")

                translated_block = clean_text(
                    translations.get(label_id, label_element.get("text", ""))
                )
                translated_lines = [
                    clean_text(part)
                    for part in translated_block.splitlines()
                    if clean_text(part)
                ]

                if line_index is not None and 0 <= int(line_index) < len(translated_lines):
                    label = translated_lines[int(line_index)]
                elif len(translated_lines) == 1:
                    label = translated_lines[0]

                claimed_label_keys.add(label_match["claim_key"])
                if label_id:
                    rendered_ids.add(label_id)

        # Remove the default paragraph content and rebuild the field according
        # to the detected PDF geometry. This is intentionally generic: labels
        # below the PDF line stay below it; labels above stay above it.
        first_p = cell.paragraphs[0]
        first_p.text = ""
        label_position = field.get("label_position", "above")

        def _configure(p, blank=False):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1

        def _add_label(p):
            _configure(p)
            run = p.add_run(label if label else " ")
            run.font.name = DEFAULT_FONT_NAME
            run.font.size = Pt(FORM_DOCX_FONT_PT)
            apply_hex_font_color(run, field.get("label_color", "#000000"))

            # Keep helper text in the SAME paragraph as the primary label.
            # Separate helper paragraphs make short form rows much taller and
            # can push only part of a three-column row onto the next Word page
            # near the bottom of the source PDF. A line break preserves the
            # visual hierarchy while keeping the complete field compact.
            for helper_text in helper_labels:
                run.add_break()
                helper_run = p.add_run(helper_text)
                helper_run.font.name = DEFAULT_FONT_NAME
                helper_run.font.size = Pt(max(6.25, FORM_DOCX_FONT_PT - 0.75))
                apply_hex_font_color(
                    helper_run, field.get("label_color", "#000000")
                )

        def _add_writing_line(p):
            _configure(p, blank=True)
            p.paragraph_format.line_spacing = Pt(FORM_FIELD_BLANK_HEIGHT_PT)
            r = p.add_run(" ")
            r.font.size = Pt(FORM_FIELD_BLANK_HEIGHT_PT)
            set_paragraph_bottom_border(p, size=5, color="000000")

        if label_position == "below":
            _add_writing_line(first_p)
            _add_label(cell.add_paragraph())
        else:
            _add_label(first_p)
            _add_writing_line(cell.add_paragraph())

        # Keep only real vertical box boundaries at cell level. The horizontal
        # writing line is rendered as a paragraph border so its relation to the
        # label can be preserved on either side.
        set_cell_border(
            cell,
            top=None,
            left=({"val": "single", "sz": 4, "color": "000000"} if field.get("left_border") else None),
            bottom=None,
            right=({"val": "single", "sz": 4, "color": "000000"} if field.get("right_border") else None),
        )

    return row_y


def build_form_page_events(
    page_data,
    page_elements,
    rendered_ids=None,
):
    rendered_ids = rendered_ids or set()
    label_bboxes = get_form_label_bboxes(
        page_data
    )
    form_region = get_form_render_region(page_data)

    events = []

    for element in page_elements:
        if element.get("id") in rendered_ids:
            continue
        if element.get("type") not in (
            "text_block",
            "list_item",
        ):
            continue

        # Mixed form blocks must be handled line-by-line. A whole PDF block
        # may contain one detected field label plus unrelated structural text
        # (section headings, place/location labels, instructions). Earlier
        # code discarded the whole block as soon as one label overlapped it.
        element_bbox = element.get("bbox", {})
        allow_form_fallback = bbox_center_inside_region(
            element_bbox,
            form_region,
        )

        line_events = (
            form_element_line_events(
                element,
                label_bboxes,
            )
            if allow_form_fallback
            else []
        )

        if line_events:
            for item in line_events:
                bbox = item["bbox"]
                events.append({
                    "kind": "text_line",
                    "y": float(bbox.get("y0", 0)),
                    "x": float(bbox.get("x0", 0)),
                    "element": element,
                    "line_index": item["line_index"],
                    "line": item["line"],
                    "bbox": bbox,
                    "fallback_text": item["text"],
                })
            continue

        if element_is_used_as_form_label(
            element,
            label_bboxes,
        ):
            continue

        bbox = element.get(
            "bbox",
            {},
        )

        events.append({
            "kind": "text",
            "y": float(
                bbox.get(
                    "y0",
                    0,
                )
            ),
            "x": float(
                bbox.get(
                    "x0",
                    0,
                )
            ),
            "element": element,
        })

    for row in page_data.get(
        "form_rows",
        [],
    ):
        events.append({
            "kind": "form_row",
            "y": float(
                row.get(
                    "y",
                    0,
                )
            ),
            "x": 0.0,
            "row": row,
        })

    events.sort(
        key=lambda event: (
            event["y"],
            event["x"],
            0 if event["kind"] == "text" else 1,
        )
    )

    return events


def add_form_page(
    document,
    page_data,
    page_elements,
    translations,
    rendered_ids=None,
):
    rendered_ids = rendered_ids or set()
    page_width = float(
        page_data["width"]
    )

    field_lookup = {
        field["field_id"]: field
        for field in page_data.get(
            "form_fields",
            [],
        )
    }

    events = build_form_page_events(
        page_data,
        page_elements,
        rendered_ids,
    )

    previous_y = None
    claimed_label_keys = set()

    for event in events:
        if event["kind"] == "text":
            element = event["element"]

            text = translations.get(
                element["id"],
                element.get(
                    "text",
                    "",
                ),
            )

            previous_y = add_form_heading_block(
                document,
                element,
                text,
                page_width,
                previous_y,
            )
            rendered_ids.add(element["id"])

        elif event["kind"] == "text_line":
            element = event["element"]
            line_text = translated_form_line_text(
                element,
                event.get("line_index"),
                event.get("fallback_text", ""),
                translations,
            )
            line_element = form_line_proxy_element(
                element,
                event.get("line", {}),
                event.get("bbox", {}),
            )
            previous_y = add_form_heading_block(
                document,
                line_element,
                line_text,
                page_width,
                previous_y,
            )
            # Mark the source block as globally handled. Its field-label lines
            # can still be resolved below because add_form_row() searches the
            # page element list directly rather than filtering rendered_ids.
            rendered_ids.add(element["id"])

        elif event["kind"] == "form_row":
            previous_y = add_form_row(
                document,
                event["row"],
                field_lookup,
                page_width,
                previous_y,
                page_elements,
                translations,
                rendered_ids,
                claimed_label_keys,
            )


# ============================================================
# WORD HEADER / LETTERHEAD
# ============================================================

def clear_header(header):
    try:
        for paragraph in list(header.paragraphs):
            paragraph._element.getparent().remove(paragraph._element)
    except Exception:
        pass


def add_page_header(section, page_data, element_lookup, translations, rendered_ids):
    # Every PDF page is rendered into its own Word section. A newly-created
    # Word section inherits the previous section's header by default. If this
    # PDF page has no header and we return before unlinking/clearing it, Word
    # repeats the previous page's header on this and every overflow page.
    # Always break that inheritance first, even when header_items is empty.
    header = section.header
    header.is_linked_to_previous = False
    clear_header(header)

    header_items = page_data.get("header_blocks", [])
    if not header_items:
        return

    for item in sorted(
        header_items,
        key=lambda x: (float(x.get("bbox", {}).get("y0", 0)), float(x.get("bbox", {}).get("x0", 0))),
    ):
        element_id = item.get("id")
        if not element_id or element_id in rendered_ids:
            continue

        element = element_lookup.get(element_id, item)
        text = translations.get(element_id, element.get("text", ""))
        if not clean_text(text):
            rendered_ids.add(element_id)
            continue

        paragraph = header.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1

        bbox = element.get("bbox", {})
        paragraph.paragraph_format.left_indent = Pt(max(0, float(bbox.get("x0", 0))))

        run = paragraph.add_run(text)
        style = element.get("style", {})
        run.font.name = style.get("font") or DEFAULT_FONT_NAME
        size = float(style.get("size", 7.5) or 7.5)
        run.font.size = Pt(clamp(size, 5.5, 9.5))
        run.bold = bool(style.get("bold", False))
        run.italic = bool(style.get("italic", False))
        rendered_ids.add(element_id)

# ============================================================
# WORD FOOTER
# ============================================================

def clear_footer(footer):
    try:
        for paragraph in list(footer.paragraphs):
            paragraph._element.getparent().remove(paragraph._element)
    except Exception:
        pass


def add_page_footer(
    section,
    page_data,
    element_lookup,
    translations,
    rendered_ids,
):
    # Same rule as headers: new Word sections inherit the previous footer.
    # Break the link and clear it before checking whether this PDF page has
    # footer content.
    footer = section.footer
    footer.is_linked_to_previous = False
    clear_footer(footer)

    footer_items = page_data.get("footer_blocks", [])
    if not footer_items:
        return

    sorted_items = sorted(
        footer_items,
        key=lambda item: (
            float(item.get("bbox", {}).get("y0", 0)),
            float(item.get("bbox", {}).get("x0", 0)),
        ),
    )

    for item in sorted_items:
        element_id = item.get("id")
        if not element_id or element_id in rendered_ids:
            continue
        element = element_lookup.get(element_id, item)

        text = translations.get(
            element_id,
            element.get("text", ""),
        )

        if not clean_text(text):
            continue

        paragraph = footer.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1

        bbox = element.get("bbox", {})
        paragraph.paragraph_format.left_indent = Pt(
            max(0, float(bbox.get("x0", 0)))
        )

        run = paragraph.add_run(text)
        style = element.get("style", {})
        run.font.name = style.get("font") or DEFAULT_FONT_NAME

        size = float(style.get("size", 7.0) or 7.0)
        run.font.size = Pt(clamp(size, 5.5, 8.5))
        run.bold = bool(style.get("bold", False))
        run.italic = bool(style.get("italic", False))
        rendered_ids.add(element_id)

# ============================================================
# CREATE WORD
# ============================================================

def create_word_from_structure(
    pages,
    elements,
    translations,
):
    # IMPORTANT: /create-docx must render the exact structure returned by
    # /extract-pdf. It must never open/re-extract the original PDF here.
    # This keeps pdf_N IDs, form fields, tables, header/footer classification
    # and spatial de-duplication identical between translation and rendering.
    if not isinstance(pages, list) or not pages:
        raise ValueError("Structure pages[] is missing or empty")

    if not isinstance(elements, list):
        raise ValueError("Structure elements[] is missing")

    document = Document()

    if document.paragraphs:
        p = document.paragraphs[0]
        p.text = ""
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    element_lookup = {
        element["id"]: element
        for element in elements
    }

    # One render ledger for the entire DOCX.  An extracted pdf_N element may
    # be consumed by header, footer, form, table/body logic, but never twice.
    global_rendered_ids = set()

    elements_by_page = {}

    for element in elements:
        elements_by_page.setdefault(
            int(element.get("page", 0)),
            [],
        ).append(element)

    for page_index, page_data in enumerate(
        pages
    ):
        page_width = float(
            page_data["width"]
        )
        page_height = float(
            page_data["height"]
        )

        if page_index == 0:
            section = document.sections[0]
        else:
            section = document.add_section(
                WD_SECTION.NEW_PAGE
            )

        configure_section(
            section,
            page_width,
            page_height,
        )

        # Deliberately reuse the document-wide set. Do not reset per page.
        rendered_ids = global_rendered_ids

        add_page_header(
            section,
            page_data,
            element_lookup,
            translations,
            rendered_ids,
        )

        add_page_footer(
            section,
            page_data,
            element_lookup,
            translations,
            rendered_ids,
        )

        page_number = int(
            page_data.get(
                "page",
                page_index + 1,
            )
        )

        page_elements = elements_by_page.get(
            page_number,
            [],
        )

        # ====================================================
        # FORM PAGE
        # ====================================================
        if page_is_form_page(
            page_data
        ):
            add_form_page(
                document,
                page_data,
                page_elements,
                translations,
                rendered_ids,
            )
            continue

        # ====================================================
        # NORMAL PAGE
        # ====================================================
        table_lookup = {
            table["table_id"]: table
            for table in page_data.get(
                "tables",
                [],
            )
        }

        previous_bottom = None

        for item in page_data.get(
            "layout",
            [],
        ):
            if item.get("type") == "table":
                table_data = table_lookup.get(
                    item.get(
                        "table_id"
                    )
                )

                if table_data:
                    # Skip a table only if every translated cell has already
                    # been consumed elsewhere. Otherwise render it once and
                    # claim all of its cell IDs globally.
                    table_cell_ids = [
                        cell.get("id")
                        for row_data in table_data.get("rows", [])
                        for cell in row_data
                        if cell.get("id")
                    ]

                    if not table_cell_ids or not all(
                        cell_id in rendered_ids for cell_id in table_cell_ids
                    ):
                        add_basic_table(
                            document,
                            table_data,
                            translations,
                        )
                        rendered_ids.update(table_cell_ids)

                    previous_bottom = float(
                        table_data[
                            "bbox"
                        ][
                            "y1"
                        ]
                    )

                continue

            element_id = item.get("id")
            if not element_id or element_id in rendered_ids:
                continue
            element = element_lookup.get(
                element_id
            )

            if not element:
                continue

            translated_text = (
                translations.get(
                    element_id,
                    element.get(
                        "text",
                        "",
                    ),
                )
            )

            previous_bottom = add_basic_text_block(
                document,
                element,
                translated_text,
                page_width,
                previous_bottom,
            )
            rendered_ids.add(element_id)

    output = BytesIO()
    document.save(output)
    output.seek(0)

    return output

# ============================================================
# TEMPORARY STRUCTURE STORE
# ============================================================

def _ensure_structure_store():
    os.makedirs(STRUCTURE_STORE_DIR, exist_ok=True)


def _structure_path(structure_id):
    if not isinstance(structure_id, str) or not re.fullmatch(r"[0-9a-f]{32}", structure_id):
        raise ValueError("Invalid structure_id")
    return os.path.join(STRUCTURE_STORE_DIR, f"{structure_id}.json.gz")


def cleanup_expired_structures():
    """Best-effort cleanup of expired temporary structure files."""
    try:
        _ensure_structure_store()
        cutoff = time.time() - STRUCTURE_TTL_SECONDS
        for name in os.listdir(STRUCTURE_STORE_DIR):
            if not name.endswith(".json.gz"):
                continue
            path = os.path.join(STRUCTURE_STORE_DIR, name)
            try:
                if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                    os.remove(path)
            except OSError:
                pass
    except OSError:
        pass


def save_extracted_structure(pages, elements, original_filename=None):
    """Persist the exact extraction result and return a small opaque ID."""
    _ensure_structure_store()
    cleanup_expired_structures()

    structure_id = uuid.uuid4().hex
    path = _structure_path(structure_id)
    tmp_path = f"{path}.tmp"

    payload = {
        "created_at": time.time(),
        "filename": original_filename or "translated.pdf",
        "pages": pages,
        "elements": elements,
    }

    with gzip.open(tmp_path, "wt", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, separators=(",", ":"))

    os.replace(tmp_path, path)
    return structure_id


def load_extracted_structure(structure_id):
    """Load a structure previously created by /extract-pdf."""
    cleanup_expired_structures()
    path = _structure_path(structure_id)

    if not os.path.exists(path):
        raise FileNotFoundError(
            "structure_id was not found or has expired. Run /extract-pdf again."
        )

    age = time.time() - os.path.getmtime(path)
    if age > STRUCTURE_TTL_SECONDS:
        try:
            os.remove(path)
        except OSError:
            pass
        raise FileNotFoundError(
            "structure_id has expired. Run /extract-pdf again."
        )

    with gzip.open(path, "rt", encoding="utf-8") as fh:
        payload = json.load(fh)

    if not isinstance(payload, dict):
        raise ValueError("Stored structure is invalid")

    pages = payload.get("pages")
    elements = payload.get("elements")

    if not isinstance(pages, list) or not pages:
        raise ValueError("Stored structure.pages must be a non-empty array")
    if not isinstance(elements, list):
        raise ValueError("Stored structure.elements must be an array")

    return payload


# ============================================================
# EXTRACT PDF ROUTE
# ============================================================

@app.route("/extract-pdf", methods=["POST"])
def extract_pdf():
    if "file" not in request.files:
        return jsonify({
            "error": "No PDF file received"
        }), 400

    uploaded_file = request.files["file"]

    try:
        pdf_bytes = uploaded_file.read()
        document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf",
        )
    except Exception as e:
        return jsonify({
            "error": "Could not read PDF file",
            "details": str(e),
        }), 400

    try:
        pages, elements, chunks = extract_pdf_structure(
            document
        )

        # Store the exact extraction result server-side. Make.com only needs
        # to carry the small structure_id between /extract-pdf and /create-docx.
        structure_id = save_extracted_structure(
            pages,
            elements,
            uploaded_file.filename,
        )

        return jsonify({
            "filename": uploaded_file.filename,
            "page_count": len(pages),
            "element_count": len(elements),
        "spatial_duplicate_count": sum(
            page.get("spatial_duplicate_count", 0) for page in pages
        ),
            "chunk_size": CHUNK_SIZE,
            "chunk_count": len(chunks),
            "table_count": sum(
                page["table_count"]
                for page in pages
            ),
            "list_count": sum(
                page["list_count"]
                for page in pages
            ),
            "line_count": sum(
                page["line_count"]
                for page in pages
            ),
            "form_line_count": sum(
                page["form_line_count"]
                for page in pages
            ),
            "form_field_count": sum(
                page["form_field_count"]
                for page in pages
            ),
            "form_row_count": sum(
                page["form_row_count"]
                for page in pages
            ),
            "header_block_count": sum(
                page.get("header_block_count", 0)
                for page in pages
            ),
            "footer_block_count": sum(
                page.get("footer_block_count", 0)
                for page in pages
            ),
            "rectangle_count": sum(
                page["rectangle_count"]
                for page in pages
            ),
            "structure_id": structure_id,
            "structure_ttl_seconds": STRUCTURE_TTL_SECONDS,
            "chunks": chunks,
        })

    finally:
        document.close()

# ============================================================
# CREATE DOCX ROUTE
# ============================================================

@app.route("/create-docx", methods=["POST"])
def create_docx():
    translations_raw = request.form.get("translations")
    structure_id = (request.form.get("structure_id") or "").strip()

    if not translations_raw:
        return jsonify({
            "error": "No translations JSON received"
        }), 400

    if not structure_id:
        return jsonify({
            "error": "No structure_id received",
            "expected": {
                "structure_id": "structure_id returned by /extract-pdf"
            }
        }), 400

    try:
        translations = parse_translations(translations_raw)
    except Exception as e:
        return jsonify({
            "error": "Could not parse translations JSON",
            "details": str(e),
        }), 400

    if not translations:
        return jsonify({
            "error": "Translations list is empty"
        }), 400

    try:
        stored_structure = load_extracted_structure(structure_id)
        pages = stored_structure["pages"]
        elements = stored_structure["elements"]

        # Defensive validation: every renderable extracted element should keep
        # the deterministic ID assigned by /extract-pdf.
        seen_ids = set()
        duplicate_ids = []
        for element in elements:
            if not isinstance(element, dict):
                continue
            element_id = element.get("id")
            if not element_id:
                continue
            if element_id in seen_ids:
                duplicate_ids.append(element_id)
            seen_ids.add(element_id)

        if duplicate_ids:
            raise ValueError(
                "Duplicate element IDs in stored structure: "
                + ", ".join(duplicate_ids[:20])
            )

    except FileNotFoundError as e:
        return jsonify({
            "error": "Stored PDF structure not found",
            "details": str(e),
        }), 410
    except Exception as e:
        return jsonify({
            "error": "Could not load stored PDF structure",
            "details": str(e),
        }), 400

    try:
        word_file = create_word_from_structure(
            pages,
            elements,
            translations,
        )
    except Exception as e:
        return jsonify({
            "error": "Could not create DOCX",
            "details": str(e),
        }), 500

    # The original PDF is intentionally NOT required here. A filename can be
    # supplied separately by Make; otherwise use a safe default.
    original_name = (
        request.form.get("filename")
        or stored_structure.get("filename")
        or "translated.pdf"
    )

    if original_name.lower().endswith(".pdf"):
        original_name = original_name[:-4]

    output_filename = f"{original_name}_translated.docx"

    return send_file(
        word_file,
        as_attachment=True,
        download_name=output_filename,
        mimetype=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )

# ============================================================
# LOCAL
# ============================================================

if __name__ == "__main__":
    app.run()
