from flask import Flask, request, jsonify, send_file

import fitz
import json
import re
import statistics

from io import BytesIO

from docx import Document
from docx.shared import Pt
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

HEADING_SIZE_RATIO = 1.18
HEADING_LARGE_RATIO = 1.45

COLUMN_MIN_BLOCKS = 4
COLUMN_GAP_RATIO = 0.08

LINE_TOLERANCE = 1.5


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
        "y1": round_num(rect[3])
    }


def rect_center_inside(inner_rect, outer_rect):
    cx = (inner_rect[0] + inner_rect[2]) / 2
    cy = (inner_rect[1] + inner_rect[3]) / 2

    return (
        outer_rect[0] <= cx <= outer_rect[2]
        and
        outer_rect[1] <= cy <= outer_rect[3]
    )


def rects_intersect(a, b):
    return not (
        a[2] <= b[0]
        or a[0] >= b[2]
        or a[3] <= b[1]
        or a[1] >= b[3]
    )


def clean_text(value):
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\u00ad", "")
    value = value.replace("\u00a0", " ")

    return value.strip()


# ============================================================
# FONT / STYLE HELPERS
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
            "hex": f"#{r:02X}{g:02X}{b:02X}"
        }

    except Exception:
        return {
            "r": 0,
            "g": 0,
            "b": 0,
            "hex": "#000000"
        }


def span_is_bold(span):
    flags = int(span.get("flags", 0))
    font = str(span.get("font", "")).lower()

    return (
        bool(flags & 16)
        or
        "bold" in font
        or
        "black" in font
        or
        "heavy" in font
        or
        "semibold" in font
        or
        "demi" in font
    )


def span_is_italic(span):
    flags = int(span.get("flags", 0))
    font = str(span.get("font", "")).lower()

    return (
        bool(flags & 2)
        or
        "italic" in font
        or
        "oblique" in font
    )


def span_is_monospace(span):
    flags = int(span.get("flags", 0))

    return bool(flags & 8)


def span_is_serif(span):
    flags = int(span.get("flags", 0))

    return bool(flags & 4)


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
        "bbox": bbox_dict(bbox)
    }


def get_block_style(spans):
    valid_spans = [
        span
        for span in spans
        if span.get("text")
    ]

    if not valid_spans:
        return {
            "font": "",
            "size": 0,
            "max_size": 0,
            "bold": False,
            "italic": False,
            "color": "#000000"
        }

    sizes = [
        float(span.get("size", 0))
        for span in valid_spans
        if float(span.get("size", 0)) > 0
    ]

    fonts = [
        span.get("font", "")
        for span in valid_spans
        if span.get("font")
    ]

    colors = [
        span.get("color", {}).get("hex", "#000000")
        for span in valid_spans
    ]

    font = ""

    if fonts:
        try:
            font = statistics.mode(fonts)
        except Exception:
            font = fonts[0]

    color = "#000000"

    if colors:
        try:
            color = statistics.mode(colors)
        except Exception:
            color = colors[0]

    return {
        "font": font,
        "size": round_num(statistics.median(sizes)) if sizes else 0,
        "max_size": round_num(max(sizes)) if sizes else 0,
        "bold": any(span.get("bold") for span in valid_spans),
        "italic": any(span.get("italic") for span in valid_spans),
        "color": color
    }


# ============================================================
# LIST DETECTION
# ============================================================

BULLET_REGEX = re.compile(
    r"^\s*("
    r"[•●▪◦‣⁃·]"
    r"|[\uF0B7]"
    r"|[-–—]"
    r")\s+"
)

NUMBER_LIST_REGEX = re.compile(
    r"^\s*("
    r"\d+[\.\)]"
    r"|\(\d+\)"
    r"|[a-zA-Z][\.\)]"
    r"|\([a-zA-Z]\)"
    r")\s+"
)


def detect_list_info(text):
    if not ENABLE_LIST_DETECTION:
        return None

    if not text:
        return None

    first_line = text.splitlines()[0].strip()

    bullet_match = BULLET_REGEX.match(first_line)

    if bullet_match:
        return {
            "is_list": True,
            "list_type": "bullet",
            "marker": bullet_match.group(1)
        }

    number_match = NUMBER_LIST_REGEX.match(first_line)

    if number_match:
        return {
            "is_list": True,
            "list_type": "numbered",
            "marker": number_match.group(1)
        }

    return None


# ============================================================
# DRAWING / LINE DETECTION
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
            "bbox": (
                bbox_dict(drawing_bbox)
                if drawing_bbox is not None
                else None
            ),
            "fill": str(drawing.get("fill")),
            "color": str(drawing.get("color")),
            "width": round_num(drawing.get("width", 0))
        }

        drawing_counter += 1

        drawing_items = []

        for item in drawing.get("items", []):
            if not item:
                continue

            command = item[0]

            # ------------------------------------------------
            # LINE
            # ------------------------------------------------

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

                line = {
                    "id": f"line_{line_counter}",
                    "type": "line",
                    "x0": round_num(x0),
                    "y0": round_num(y0),
                    "x1": round_num(x1),
                    "y1": round_num(y1),
                    "orientation": orientation,
                    "length": round_num(
                        ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5
                    ),
                    "width": round_num(drawing.get("width", 0))
                }

                line_counter += 1
                lines_out.append(line)
                drawing_items.append(line)

            # ------------------------------------------------
            # RECTANGLE
            # ------------------------------------------------

            elif command == "re" and len(item) >= 2:
                rect = item[1]

                try:
                    rect_tuple = (
                        float(rect.x0),
                        float(rect.y0),
                        float(rect.x1),
                        float(rect.y1)
                    )

                    rect_data = {
                        "id": f"rect_{rect_counter}",
                        "type": "rectangle",
                        "bbox": bbox_dict(rect_tuple),
                        "width": round_num(
                            rect_tuple[2] - rect_tuple[0]
                        ),
                        "height": round_num(
                            rect_tuple[3] - rect_tuple[1]
                        )
                    }

                    rect_counter += 1
                    rectangles_out.append(rect_data)
                    drawing_items.append(rect_data)

                except Exception:
                    pass

        drawing_data["items"] = drawing_items

        if drawing_items:
            drawings_out.append(drawing_data)

    return (
        drawings_out,
        lines_out,
        rectangles_out
    )


# ============================================================
# FORM LINE DETECTION
# ============================================================

def detect_form_lines(lines, page_width):
    form_lines = []

    minimum_form_line = page_width * 0.08
    maximum_form_line = page_width * 0.80

    for line in lines:
        if line.get("orientation") != "horizontal":
            continue

        length = float(line.get("length", 0))

        if (
            length >= minimum_form_line
            and
            length <= maximum_form_line
        ):
            form_lines.append(line)

    return form_lines


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

        cleaned_row = []

        for value in row:
            cleaned_row.append(
                clean_text(value)
            )

        max_columns = max(
            max_columns,
            len(cleaned_row)
        )

        cleaned.append(cleaned_row)

    if max_columns == 0:
        return []

    for row in cleaned:
        while len(row) < max_columns:
            row.append("")

    cleaned = [
        row
        for row in cleaned
        if any(value.strip() for value in row)
    ]

    return cleaned


def extract_table_cell_bboxes(detected_table):
    result = []

    try:
        rows = detected_table.rows

        for row_index, row in enumerate(rows):
            row_result = []

            for col_index, cell in enumerate(row.cells):
                if cell is None:
                    row_result.append(None)
                    continue

                row_result.append(
                    bbox_dict(cell)
                )

            result.append(row_result)

    except Exception:
        pass

    return result


# ============================================================
# TEXT DICTIONARY EXTRACTION
# ============================================================

def extract_raw_text_blocks(page):
    try:
        text_dict = page.get_text(
            "dict",
            flags=fitz.TEXT_PRESERVE_WHITESPACE
        )
    except Exception:
        text_dict = page.get_text("dict")

    blocks_out = []

    for raw_block in text_dict.get("blocks", []):
        if raw_block.get("type") != 0:
            continue

        block_bbox = raw_block.get(
            "bbox",
            (0, 0, 0, 0)
        )

        lines_out = []
        all_spans = []
        text_lines = []

        for raw_line in raw_block.get("lines", []):
            line_spans = []

            for raw_span in raw_line.get("spans", []):
                normalized = normalize_span(
                    raw_span
                )

                if not normalized["text"]:
                    continue

                line_spans.append(
                    normalized
                )

                all_spans.append(
                    normalized
                )

            if not line_spans:
                continue

            line_text = "".join(
                span["text"]
                for span in line_spans
            ).strip()

            if not line_text:
                continue

            line_bbox = raw_line.get(
                "bbox",
                block_bbox
            )

            lines_out.append({
                "text": line_text,
                "bbox": bbox_dict(line_bbox),
                "spans": line_spans
            })

            text_lines.append(
                line_text
            )

        text = "\n".join(
            text_lines
        ).strip()

        if not text:
            continue

        blocks_out.append({
            "text": text,
            "bbox": bbox_dict(block_bbox),
            "lines": lines_out,
            "spans": all_spans,
            "style": get_block_style(all_spans)
        })

    return blocks_out


# ============================================================
# PAGE FONT STATISTICS
# ============================================================

def get_page_font_statistics(blocks):
    sizes = []

    for block in blocks:
        for span in block.get("spans", []):
            size = float(
                span.get("size", 0)
            )

            if size > 0:
                sizes.append(size)

    if not sizes:
        return {
            "median_size": DEFAULT_FONT_SIZE,
            "max_size": DEFAULT_FONT_SIZE
        }

    return {
        "median_size": round_num(
            statistics.median(sizes)
        ),
        "max_size": round_num(
            max(sizes)
        )
    }


# ============================================================
# HEADING DETECTION
# ============================================================

def classify_text_role(block, font_stats):
    text = clean_text(
        block.get("text", "")
    )

    style = block.get(
        "style",
        {}
    )

    median_size = max(
        1,
        float(
            font_stats.get(
                "median_size",
                DEFAULT_FONT_SIZE
            )
        )
    )

    max_size = float(
        style.get(
            "max_size",
            0
        )
    )

    bold = bool(
        style.get(
            "bold",
            False
        )
    )

    line_count = max(
        1,
        len(
            block.get(
                "lines",
                []
            )
        )
    )

    short_text = len(text) <= 120

    if (
        short_text
        and
        max_size >= median_size * HEADING_LARGE_RATIO
    ):
        return "heading_1"

    if (
        short_text
        and
        bold
        and
        max_size >= median_size * HEADING_SIZE_RATIO
    ):
        return "heading_2"

    if (
        short_text
        and
        bold
        and
        line_count <= 2
    ):
        return "subheading"

    return "paragraph"


# ============================================================
# COLUMN DETECTION
# ============================================================

def detect_columns(blocks, page_width):
    if not ENABLE_COLUMN_DETECTION:
        return {
            "column_count": 1,
            "divider_x": None
        }

    candidates = []

    for block in blocks:
        bbox = block.get(
            "bbox",
            {}
        )

        x0 = float(
            bbox.get(
                "x0",
                0
            )
        )

        x1 = float(
            bbox.get(
                "x1",
                0
            )
        )

        width = x1 - x0

        if (
            width > 10
            and
            width < page_width * 0.70
        ):
            candidates.append(
                (x0, x1)
            )

    if len(candidates) < COLUMN_MIN_BLOCKS * 2:
        return {
            "column_count": 1,
            "divider_x": None
        }

    page_center = page_width / 2

    left = [
        item
        for item in candidates
        if item[0] < page_center
    ]

    right = [
        item
        for item in candidates
        if item[0] >= page_center * 0.85
    ]

    if (
        len(left) < COLUMN_MIN_BLOCKS
        or
        len(right) < COLUMN_MIN_BLOCKS
    ):
        return {
            "column_count": 1,
            "divider_x": None
        }

    left_right_edge = statistics.median(
        item[1]
        for item in left
    )

    right_left_edge = statistics.median(
        item[0]
        for item in right
    )

    gap = (
        right_left_edge -
        left_right_edge
    )

    if gap >= page_width * COLUMN_GAP_RATIO:
        return {
            "column_count": 2,
            "divider_x": round_num(
                (
                    left_right_edge +
                    right_left_edge
                ) / 2
            )
        }

    return {
        "column_count": 1,
        "divider_x": None
    }


# ============================================================
# UNDERLINE DETECTION
# ============================================================

def detect_span_underlines(spans, horizontal_lines):
    for span in spans:
        span["underline"] = False

        bbox = span.get(
            "bbox",
            {}
        )

        x0 = float(bbox.get("x0", 0))
        y1 = float(bbox.get("y1", 0))
        x1 = float(bbox.get("x1", 0))

        text_width = max(
            1,
            x1 - x0
        )

        for line in horizontal_lines:
            ly = float(line.get("y0", 0))

            lx0 = min(
                float(line.get("x0", 0)),
                float(line.get("x1", 0))
            )

            lx1 = max(
                float(line.get("x0", 0)),
                float(line.get("x1", 0))
            )

            vertical_distance = abs(
                ly - y1
            )

            overlap = max(
                0,
                min(x1, lx1) -
                max(x0, lx0)
            )

            overlap_ratio = (
                overlap /
                text_width
            )

            if (
                vertical_distance <= 3.0
                and
                overlap_ratio >= 0.50
            ):
                span["underline"] = True
                break


# ============================================================
# MAIN PDF STRUCTURE EXTRACTION
# ============================================================

def extract_pdf_structure(document):
    pages = []
    elements = []

    element_counter = 1
    table_counter = 1

    for page_index in range(
        len(document)
    ):
        page = document[
            page_index
        ]

        page_number = (
            page_index + 1
        )

        page_width = float(
            page.rect.width
        )

        page_height = float(
            page.rect.height
        )

        # ====================================================
        # DRAWINGS
        # ====================================================

        (
            drawings,
            lines,
            rectangles
        ) = normalize_drawings(
            page
        )

        horizontal_lines = [
            line
            for line in lines
            if line.get(
                "orientation"
            ) == "horizontal"
        ]

        vertical_lines = [
            line
            for line in lines
            if line.get(
                "orientation"
            ) == "vertical"
        ]

        form_lines = detect_form_lines(
            horizontal_lines,
            page_width
        )

        # ====================================================
        # TEXT
        # ====================================================

        raw_blocks = extract_raw_text_blocks(
            page
        )

        for block in raw_blocks:
            detect_span_underlines(
                block.get(
                    "spans",
                    []
                ),
                horizontal_lines
            )

            for line_data in block.get(
                "lines",
                []
            ):
                detect_span_underlines(
                    line_data.get(
                        "spans",
                        []
                    ),
                    horizontal_lines
                )

        font_stats = get_page_font_statistics(
            raw_blocks
        )

        # ====================================================
        # TABLES
        # ====================================================

        detected_tables = find_page_tables(
            page
        )

        table_regions = []
        table_structures = []
        table_cell_elements = []

        for detected_table in detected_tables:
            try:
                bbox = detected_table.bbox

                table_bbox = (
                    float(bbox[0]),
                    float(bbox[1]),
                    float(bbox[2]),
                    float(bbox[3])
                )

                raw_matrix = (
                    detected_table.extract()
                )

                matrix = clean_table_matrix(
                    raw_matrix
                )

            except Exception as e:
                print(
                    "Could not parse table:",
                    str(e)
                )
                continue

            if not matrix:
                continue

            row_count = len(matrix)

            column_count = max(
                len(row)
                for row in matrix
            )

            if (
                row_count == 0
                or
                column_count == 0
            ):
                continue

            table_id = (
                f"table_{table_counter}"
            )

            table_counter += 1

            cell_bboxes = (
                extract_table_cell_bboxes(
                    detected_table
                )
            )

            table_data = {
                "table_id": table_id,
                "type": "table",
                "page": page_number,
                "bbox": bbox_dict(
                    table_bbox
                ),
                "row_count": row_count,
                "column_count": column_count,
                "rows": []
            }

            for row_index in range(
                row_count
            ):
                row_data = []

                for col_index in range(
                    column_count
                ):
                    value = ""

                    if (
                        row_index < len(matrix)
                        and
                        col_index < len(
                            matrix[row_index]
                        )
                    ):
                        value = clean_text(
                            matrix[
                                row_index
                            ][
                                col_index
                            ]
                        )

                    cell_bbox = None

                    if (
                        row_index < len(
                            cell_bboxes
                        )
                        and
                        col_index < len(
                            cell_bboxes[row_index]
                        )
                    ):
                        cell_bbox = (
                            cell_bboxes[
                                row_index
                            ][
                                col_index
                            ]
                        )

                    if not cell_bbox:
                        cell_bbox = bbox_dict(
                            table_bbox
                        )

                    if value:
                        element_id = (
                            f"pdf_{element_counter}"
                        )

                        element_counter += 1

                        element = {
                            "id": element_id,
                            "page": page_number,
                            "type": "table_cell",
                            "table_id": table_id,
                            "row": row_index,
                            "col": col_index,
                            "text": value,
                            "bbox": cell_bbox
                        }

                        elements.append(
                            element
                        )

                        table_cell_elements.append(
                            element
                        )

                    else:
                        element_id = None

                    row_data.append({
                        "id": element_id,
                        "row": row_index,
                        "col": col_index,
                        "text": value,
                        "bbox": cell_bbox
                    })

                table_data[
                    "rows"
                ].append(
                    row_data
                )

            table_regions.append(
                table_bbox
            )

            table_structures.append(
                table_data
            )

        # ====================================================
        # NON-TABLE TEXT
        # ====================================================

        text_elements = []
        detected_lists = []

        for block in raw_blocks:
            bbox = block.get(
                "bbox",
                {}
            )

            block_rect = (
                float(bbox.get("x0", 0)),
                float(bbox.get("y0", 0)),
                float(bbox.get("x1", 0)),
                float(bbox.get("y1", 0))
            )

            inside_table = False

            for table_rect in table_regions:
                if rect_center_inside(
                    block_rect,
                    table_rect
                ):
                    inside_table = True
                    break

            if inside_table:
                continue

            text = clean_text(
                block.get(
                    "text",
                    ""
                )
            )

            if not text:
                continue

            element_id = (
                f"pdf_{element_counter}"
            )

            element_counter += 1

            text_role = classify_text_role(
                block,
                font_stats
            )

            list_info = detect_list_info(
                text
            )

            element_type = "text_block"

            if list_info:
                element_type = "list_item"

            element = {
                "id": element_id,
                "page": page_number,
                "type": element_type,
                "text": text,
                "bbox": block["bbox"],
                "role": text_role,
                "style": block.get(
                    "style",
                    {}
                ),
                "lines": block.get(
                    "lines",
                    []
                ),
                "spans": block.get(
                    "spans",
                    []
                ),
                "list": list_info
            }

            elements.append(
                element
            )

            text_elements.append(
                element
            )

            if list_info:
                detected_lists.append({
                    "id": element_id,
                    "bbox": block[
                        "bbox"
                    ],
                    "list_type": list_info[
                        "list_type"
                    ],
                    "marker": list_info[
                        "marker"
                    ],
                    "text": text
                })

        # ====================================================
        # COLUMN DETECTION
        # ====================================================

        columns = detect_columns(
            raw_blocks,
            page_width
        )

        # ====================================================
        # PAGE LAYOUT ORDER
        # ====================================================

        layout_items = []

        for element in text_elements:
            layout_items.append({
                "type": element[
                    "type"
                ],
                "id": element[
                    "id"
                ],
                "x0": element[
                    "bbox"
                ][
                    "x0"
                ],
                "y0": element[
                    "bbox"
                ][
                    "y0"
                ]
            })

        for table in table_structures:
            layout_items.append({
                "type": "table",
                "table_id": table[
                    "table_id"
                ],
                "x0": table[
                    "bbox"
                ][
                    "x0"
                ],
                "y0": table[
                    "bbox"
                ][
                    "y0"
                ]
            })

        layout_items.sort(
            key=lambda item: (
                float(
                    item.get(
                        "y0",
                        0
                    )
                ),
                float(
                    item.get(
                        "x0",
                        0
                    )
                )
            )
        )

        # ====================================================
        # PAGE RESPONSE
        # ====================================================

        pages.append({
            "page": page_number,
            "width": round_num(
                page_width
            ),
            "height": round_num(
                page_height
            ),

            "font_statistics": font_stats,

            "column_count": columns[
                "column_count"
            ],

            "column_divider_x": columns[
                "divider_x"
            ],

            "text_block_count": len(
                text_elements
            ),

            "table_count": len(
                table_structures
            ),

            "list_count": len(
                detected_lists
            ),

            "drawing_count": len(
                drawings
            ),

            "line_count": len(
                lines
            ),

            "horizontal_line_count": len(
                horizontal_lines
            ),

            "vertical_line_count": len(
                vertical_lines
            ),

            "rectangle_count": len(
                rectangles
            ),

            "form_line_count": len(
                form_lines
            ),

            "tables": table_structures,
            "lists": detected_lists,
            "drawings": drawings,
            "lines": lines,
            "rectangles": rectangles,
            "form_lines": form_lines,
            "layout": layout_items
        })

    # ========================================================
    # TRANSLATION CHUNKS
    # ========================================================

    elements.sort(
        key=lambda item: int(
            item[
                "id"
            ].split("_")[1]
        )
    )

    translation_elements = []

    for element in elements:
        translation_elements.append({
            "id": element[
                "id"
            ],
            "type": element[
                "type"
            ],
            "text": element[
                "text"
            ]
        })

    chunks = [
        translation_elements[
            i:i + CHUNK_SIZE
        ]
        for i in range(
            0,
            len(
                translation_elements
            ),
            CHUNK_SIZE
        )
    ]

    return (
        pages,
        elements,
        chunks
    )


# ============================================================
# TRANSLATION JSON
# ============================================================

def parse_translations(raw_value):
    if raw_value is None:
        return {}

    if isinstance(
        raw_value,
        bytes
    ):
        raw_value = raw_value.decode(
            "utf-8"
        )

    if isinstance(
        raw_value,
        str
    ):
        data = json.loads(
            raw_value
        )
    else:
        data = raw_value

    translations = data.get(
        "translations",
        []
    )

    result = {}

    for item in translations:
        element_id = item.get(
            "id"
        )

        text = item.get(
            "text"
        )

        if not element_id:
            continue

        if text is None:
            text = ""

        result[
            str(element_id)
        ] = str(text)

    return result


# ============================================================
# WORD HELPERS
# ============================================================

def configure_section(
    section,
    page_width,
    page_height
):
    section.page_width = Pt(
        page_width
    )

    section.page_height = Pt(
        page_height
    )

    section.top_margin = Pt(0)
    section.bottom_margin = Pt(0)
    section.left_margin = Pt(0)
    section.right_margin = Pt(0)
    section.header_distance = Pt(0)
    section.footer_distance = Pt(0)


def set_cell_margins(
    cell,
    top=40,
    start=60,
    bottom=40,
    end=60
):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()

    tc_mar = tc_pr.first_child_found_in(
        "w:tcMar"
    )

    if tc_mar is None:
        tc_mar = OxmlElement(
            "w:tcMar"
        )

        tc_pr.append(
            tc_mar
        )

    for margin_name, margin_value in [
        ("top", top),
        ("start", start),
        ("bottom", bottom),
        ("end", end)
    ]:
        node = tc_mar.find(
            qn(
                f"w:{margin_name}"
            )
        )

        if node is None:
            node = OxmlElement(
                f"w:{margin_name}"
            )

            tc_mar.append(
                node
            )

        node.set(
            qn("w:w"),
            str(
                margin_value
            )
        )

        node.set(
            qn("w:type"),
            "dxa"
        )


def prevent_row_split(row):
    try:
        tr_pr = (
            row._tr.get_or_add_trPr()
        )

        cant_split = tr_pr.find(
            qn(
                "w:cantSplit"
            )
        )

        if cant_split is None:
            cant_split = OxmlElement(
                "w:cantSplit"
            )

            tr_pr.append(
                cant_split
            )

        cant_split.set(
            qn("w:val"),
            "1"
        )

    except Exception:
        pass


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

            tbl_pr.append(
                borders
            )

        for edge in [
            "top",
            "left",
            "bottom",
            "right",
            "insideH",
            "insideV"
        ]:
            tag = f"w:{edge}"

            border = borders.find(
                qn(tag)
            )

            if border is None:
                border = OxmlElement(
                    tag
                )

                borders.append(
                    border
                )

            border.set(
                qn("w:val"),
                "single"
            )

            border.set(
                qn("w:sz"),
                "4"
            )

            border.set(
                qn("w:space"),
                "0"
            )

            border.set(
                qn("w:color"),
                "808080"
            )

    except Exception:
        pass


# ============================================================
# BASIC WORD OUTPUT
#
# NOTE:
# Bu bölüm henüz yeni form/style bilgilerinin tamamını
# kullanmıyor. Önce extraction tarafını doğruluyoruz.
# ============================================================

def add_basic_text_block(
    document,
    element,
    translated_text,
    page_width,
    previous_bottom
):
    bbox = element[
        "bbox"
    ]

    x0 = float(
        bbox["x0"]
    )

    y0 = float(
        bbox["y0"]
    )

    x1 = float(
        bbox["x1"]
    )

    y1 = float(
        bbox["y1"]
    )

    paragraph = (
        document.add_paragraph()
    )

    paragraph.paragraph_format.left_indent = Pt(
        max(
            0,
            x0
        )
    )

    paragraph.paragraph_format.right_indent = Pt(
        max(
            0,
            page_width - x1
        )
    )

    if previous_bottom is None:
        vertical_gap = y0
    else:
        vertical_gap = (
            y0 -
            previous_bottom
        )

    vertical_gap = clamp(
        vertical_gap,
        MIN_VERTICAL_GAP_PT,
        MAX_VERTICAL_GAP_PT
    )

    paragraph.paragraph_format.space_before = Pt(
        vertical_gap
    )

    paragraph.paragraph_format.space_after = Pt(
        0
    )

    paragraph.paragraph_format.line_spacing = 1

    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.LEFT
    )

    run = paragraph.add_run(
        translated_text
    )

    style = element.get(
        "style",
        {}
    )

    run.font.name = (
        style.get(
            "font"
        )
        or
        DEFAULT_FONT_NAME
    )

    font_size = float(
        style.get(
            "size",
            0
        )
        or
        DEFAULT_FONT_SIZE
    )

    run.font.size = Pt(
        font_size
    )

    run.bold = bool(
        style.get(
            "bold",
            False
        )
    )

    run.italic = bool(
        style.get(
            "italic",
            False
        )
    )

    return y1


def add_basic_table(
    document,
    table_data,
    translations
):
    rows = int(
        table_data.get(
            "row_count",
            0
        )
    )

    cols = int(
        table_data.get(
            "column_count",
            0
        )
    )

    if rows <= 0 or cols <= 0:
        return

    table = document.add_table(
        rows=rows,
        cols=cols
    )

    table.alignment = (
        WD_TABLE_ALIGNMENT.LEFT
    )

    table.autofit = False

    set_table_borders(
        table
    )

    row_data = table_data.get(
        "rows",
        []
    )

    for row_index in range(
        rows
    ):
        word_row = table.rows[
            row_index
        ]

        prevent_row_split(
            word_row
        )

        for col_index in range(
            cols
        ):
            cell = word_row.cells[
                col_index
            ]

            cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

            set_cell_margins(
                cell
            )

            info = None

            if (
                row_index < len(
                    row_data
                )
                and
                col_index < len(
                    row_data[
                        row_index
                    ]
                )
            ):
                info = row_data[
                    row_index
                ][
                    col_index
                ]

            if not info:
                continue

            element_id = info.get(
                "id"
            )

            original = info.get(
                "text",
                ""
            )

            translated = (
                translations.get(
                    element_id,
                    original
                )
                if element_id
                else ""
            )

            paragraph = cell.paragraphs[
                0
            ]

            paragraph.paragraph_format.space_before = Pt(
                0
            )

            paragraph.paragraph_format.space_after = Pt(
                0
            )

            run = paragraph.add_run(
                translated
            )

            run.font.name = (
                DEFAULT_FONT_NAME
            )

            run.font.size = Pt(
                TABLE_FONT_SIZE
            )


def create_word_from_pdf(
    pdf_document,
    translations
):
    (
        pages,
        elements,
        chunks
    ) = extract_pdf_structure(
        pdf_document
    )

    document = Document()

    element_lookup = {
        element["id"]: element
        for element in elements
    }

    if document.paragraphs:
        p = document.paragraphs[0]
        p.text = ""
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    for page_index, page_data in enumerate(
        pages
    ):
        width = float(
            page_data["width"]
        )

        height = float(
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
            width,
            height
        )

        table_lookup = {
            table["table_id"]: table
            for table in page_data.get(
                "tables",
                []
            )
        }

        previous_bottom = None

        for item in page_data.get(
            "layout",
            []
        ):
            item_type = item.get(
                "type"
            )

            if item_type == "table":
                table_id = item.get(
                    "table_id"
                )

                table_data = table_lookup.get(
                    table_id
                )

                if table_data:
                    add_basic_table(
                        document,
                        table_data,
                        translations
                    )

                    previous_bottom = float(
                        table_data[
                            "bbox"
                        ][
                            "y1"
                        ]
                    )

                continue

            element_id = item.get(
                "id"
            )

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
                        ""
                    )
                )
            )

            previous_bottom = (
                add_basic_text_block(
                    document,
                    element,
                    translated_text,
                    width,
                    previous_bottom
                )
            )

    output = BytesIO()

    document.save(
        output
    )

    output.seek(0)

    return output


# ============================================================
# EXTRACT PDF ROUTE
# ============================================================

@app.route(
    "/extract-pdf",
    methods=["POST"]
)
def extract_pdf():
    if "file" not in request.files:
        return jsonify({
            "error":
                "No PDF file received"
        }), 400

    uploaded_file = (
        request.files["file"]
    )

    try:
        pdf_bytes = (
            uploaded_file.read()
        )

        document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

    except Exception as e:
        return jsonify({
            "error":
                "Could not read PDF file",
            "details":
                str(e)
        }), 400

    try:
        (
            pages,
            elements,
            chunks
        ) = extract_pdf_structure(
            document
        )

        table_count = sum(
            page.get(
                "table_count",
                0
            )
            for page in pages
        )

        list_count = sum(
            page.get(
                "list_count",
                0
            )
            for page in pages
        )

        line_count = sum(
            page.get(
                "line_count",
                0
            )
            for page in pages
        )

        form_line_count = sum(
            page.get(
                "form_line_count",
                0
            )
            for page in pages
        )

        rectangle_count = sum(
            page.get(
                "rectangle_count",
                0
            )
            for page in pages
        )

        response = {
            "filename":
                uploaded_file.filename,

            "page_count":
                len(pages),

            "element_count":
                len(elements),

            "chunk_size":
                CHUNK_SIZE,

            "chunk_count":
                len(chunks),

            "table_count":
                table_count,

            "list_count":
                list_count,

            "line_count":
                line_count,

            "form_line_count":
                form_line_count,

            "rectangle_count":
                rectangle_count,

            "pages":
                pages,

            "elements":
                elements,

            "chunks":
                chunks
        }

        return jsonify(
            response
        )

    finally:
        document.close()


# ============================================================
# CREATE DOCX ROUTE
# ============================================================

@app.route(
    "/create-docx",
    methods=["POST"]
)
def create_docx():
    if "file" not in request.files:
        return jsonify({
            "error":
                "No PDF file received"
        }), 400

    uploaded_file = (
        request.files["file"]
    )

    translations_raw = (
        request.form.get(
            "translations"
        )
    )

    if not translations_raw:
        return jsonify({
            "error":
                "No translations JSON received"
        }), 400

    try:
        translations = (
            parse_translations(
                translations_raw
            )
        )

    except Exception as e:
        return jsonify({
            "error":
                "Could not parse translations JSON",
            "details":
                str(e)
        }), 400

    if not translations:
        return jsonify({
            "error":
                "Translations list is empty"
        }), 400

    try:
        pdf_bytes = (
            uploaded_file.read()
        )

        pdf_document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

    except Exception as e:
        return jsonify({
            "error":
                "Could not read original PDF",
            "details":
                str(e)
        }), 400

    try:
        word_file = (
            create_word_from_pdf(
                pdf_document,
                translations
            )
        )

    except Exception as e:
        return jsonify({
            "error":
                "Could not create DOCX",
            "details":
                str(e)
        }), 500

    finally:
        pdf_document.close()

    original_name = (
        uploaded_file.filename
        or
        "translated.pdf"
    )

    if original_name.lower().endswith(
        ".pdf"
    ):
        original_name = (
            original_name[:-4]
        )

    output_filename = (
        f"{original_name}_translated.docx"
    )

    return send_file(
        word_file,
        as_attachment=True,
        download_name=output_filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# ============================================================
# LOCAL
# ============================================================

if __name__ == "__main__":
    app.run()