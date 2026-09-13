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
from docx.enum.table import (
    WD_TABLE_ALIGNMENT,
    WD_CELL_VERTICAL_ALIGNMENT
)
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
ENABLE_FORM_DETECTION = True

HEADING_SIZE_RATIO = 1.18
HEADING_LARGE_RATIO = 1.45

COLUMN_MIN_BLOCKS = 4
COLUMN_GAP_RATIO = 0.08

LINE_TOLERANCE = 1.5

# PDF'de çizgi olarak görünen ince rectangle'lar
THIN_RECT_MAX_THICKNESS = 3.0
MIN_LINE_LENGTH = 8.0

# Form detection
FORM_MIN_LINE_RATIO = 0.08
FORM_MAX_LINE_RATIO = 0.80

FORM_LABEL_MAX_ABOVE_PT = 30.0
FORM_LABEL_MAX_BELOW_PT = 10.0

FORM_ROW_Y_TOLERANCE = 8.0
FORM_COLUMN_X_TOLERANCE = 20.0

FORM_MIN_FIELD_WIDTH = 35.0

# Bir form çizgisinin yakınındaki metni ararken
FORM_LABEL_X_TOLERANCE = 12.0

# Dikey ayırıcıyı bir field sınırı kabul etmek için
FORM_VERTICAL_BORDER_TOLERANCE = 4.0


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


def clean_text(value):
    if value is None:
        return ""

    value = str(value)
    value = value.replace("\u00ad", "")
    value = value.replace("\u00a0", " ")

    return value.strip()


def rect_center_inside(inner_rect, outer_rect):
    cx = (
        inner_rect[0] +
        inner_rect[2]
    ) / 2

    cy = (
        inner_rect[1] +
        inner_rect[3]
    ) / 2

    return (
        outer_rect[0] <= cx <= outer_rect[2]
        and
        outer_rect[1] <= cy <= outer_rect[3]
    )


def rects_intersect(a, b):
    return not (
        a[2] <= b[0]
        or
        a[0] >= b[2]
        or
        a[3] <= b[1]
        or
        a[1] >= b[3]
    )


def horizontal_overlap(
    x0_a,
    x1_a,
    x0_b,
    x1_b
):
    return max(
        0,
        min(x1_a, x1_b) -
        max(x0_a, x0_b)
    )


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
    flags = int(
        span.get(
            "flags",
            0
        )
    )

    font = str(
        span.get(
            "font",
            ""
        )
    ).lower()

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
    flags = int(
        span.get(
            "flags",
            0
        )
    )

    font = str(
        span.get(
            "font",
            ""
        )
    ).lower()

    return (
        bool(flags & 2)
        or
        "italic" in font
        or
        "oblique" in font
    )


def span_is_monospace(span):
    flags = int(
        span.get(
            "flags",
            0
        )
    )

    return bool(
        flags & 8
    )


def span_is_serif(span):
    flags = int(
        span.get(
            "flags",
            0
        )
    )

    return bool(
        flags & 4
    )


def normalize_span(span):
    bbox = span.get(
        "bbox",
        (0, 0, 0, 0)
    )

    return {
        "text": clean_text(
            span.get(
                "text",
                ""
            )
        ),

        "font": str(
            span.get(
                "font",
                ""
            )
        ),

        "size": round_num(
            span.get(
                "size",
                0
            )
        ),

        "bold": span_is_bold(
            span
        ),

        "italic": span_is_italic(
            span
        ),

        "monospace": span_is_monospace(
            span
        ),

        "serif": span_is_serif(
            span
        ),

        "color": rgb_from_int(
            span.get(
                "color",
                0
            )
        ),

        "bbox": bbox_dict(
            bbox
        )
    }


def get_block_style(spans):
    valid = [
        span
        for span in spans
        if span.get("text")
    ]

    if not valid:
        return {
            "font": "",
            "size": 0,
            "max_size": 0,
            "bold": False,
            "italic": False,
            "color": "#000000"
        }

    sizes = [
        float(
            span.get(
                "size",
                0
            )
        )
        for span in valid
        if float(
            span.get(
                "size",
                0
            )
        ) > 0
    ]

    fonts = [
        span.get(
            "font",
            ""
        )
        for span in valid
        if span.get(
            "font"
        )
    ]

    colors = [
        span.get(
            "color",
            {}
        ).get(
            "hex",
            "#000000"
        )
        for span in valid
    ]

    if fonts:
        try:
            font = statistics.mode(
                fonts
            )
        except Exception:
            font = fonts[0]
    else:
        font = ""

    if colors:
        try:
            color = statistics.mode(
                colors
            )
        except Exception:
            color = colors[0]
    else:
        color = "#000000"

    return {
        "font": font,

        "size": (
            round_num(
                statistics.median(
                    sizes
                )
            )
            if sizes
            else 0
        ),

        "max_size": (
            round_num(
                max(
                    sizes
                )
            )
            if sizes
            else 0
        ),

        "bold": any(
            span.get(
                "bold"
            )
            for span in valid
        ),

        "italic": any(
            span.get(
                "italic"
            )
            for span in valid
        ),

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

    first_line = (
        text
        .splitlines()[0]
        .strip()
    )

    match = BULLET_REGEX.match(
        first_line
    )

    if match:
        return {
            "is_list": True,
            "list_type": "bullet",
            "marker": match.group(1)
        }

    match = NUMBER_LIST_REGEX.match(
        first_line
    )

    if match:
        return {
            "is_list": True,
            "list_type": "numbered",
            "marker": match.group(1)
        }

    return None


# ============================================================
# DRAWINGS
# ============================================================

def point_xy(point):
    try:
        return (
            float(point.x),
            float(point.y)
        )

    except Exception:
        try:
            return (
                float(point[0]),
                float(point[1])
            )

        except Exception:
            return (
                0.0,
                0.0
            )


def normalize_drawings(page):
    if not ENABLE_DRAWING_DETECTION:
        return [], [], []

    drawings_out = []
    lines_out = []
    rectangles_out = []

    try:
        drawings = page.get_drawings()

    except Exception as e:
        print(
            "get_drawings error:",
            str(e)
        )

        return [], [], []

    drawing_counter = 1
    line_counter = 1
    rect_counter = 1

    for drawing in drawings:

        drawing_bbox = drawing.get(
            "rect"
        )

        drawing_data = {
            "id": f"drawing_{drawing_counter}",

            "type": "drawing",

            "bbox": (
                bbox_dict(
                    drawing_bbox
                )
                if drawing_bbox is not None
                else None
            ),

            "fill": str(
                drawing.get(
                    "fill"
                )
            ),

            "color": str(
                drawing.get(
                    "color"
                )
            ),

            "width": round_num(
                drawing.get(
                    "width",
                    0
                )
            )
        }

        drawing_counter += 1

        drawing_items = []

        for item in drawing.get(
            "items",
            []
        ):

            if not item:
                continue

            command = item[0]

            # =================================================
            # REAL LINE
            # =================================================

            if (
                command == "l"
                and
                len(item) >= 3
            ):

                x0, y0 = point_xy(
                    item[1]
                )

                x1, y1 = point_xy(
                    item[2]
                )

                horizontal = (
                    abs(
                        y1 - y0
                    )
                    <=
                    LINE_TOLERANCE
                )

                vertical = (
                    abs(
                        x1 - x0
                    )
                    <=
                    LINE_TOLERANCE
                )

                orientation = "diagonal"

                if horizontal:
                    orientation = (
                        "horizontal"
                    )

                elif vertical:
                    orientation = (
                        "vertical"
                    )

                length = (
                    (
                        (x1 - x0) ** 2 +
                        (y1 - y0) ** 2
                    )
                    ** 0.5
                )

                line = {
                    "id":
                        f"line_{line_counter}",

                    "type":
                        "line",

                    "source":
                        "line",

                    "x0":
                        round_num(
                            x0
                        ),

                    "y0":
                        round_num(
                            y0
                        ),

                    "x1":
                        round_num(
                            x1
                        ),

                    "y1":
                        round_num(
                            y1
                        ),

                    "orientation":
                        orientation,

                    "length":
                        round_num(
                            length
                        ),

                    "thickness":
                        round_num(
                            drawing.get(
                                "width",
                                0
                            )
                        )
                }

                line_counter += 1

                lines_out.append(
                    line
                )

                drawing_items.append(
                    line
                )

            # =================================================
            # RECTANGLE
            # =================================================

            elif (
                command == "re"
                and
                len(item) >= 2
            ):

                rect = item[1]

                try:
                    x0 = float(
                        rect.x0
                    )

                    y0 = float(
                        rect.y0
                    )

                    x1 = float(
                        rect.x1
                    )

                    y1 = float(
                        rect.y1
                    )

                    rect_width = abs(
                        x1 - x0
                    )

                    rect_height = abs(
                        y1 - y0
                    )

                    # -----------------------------------------
                    # THIN HORIZONTAL RECTANGLE
                    # -----------------------------------------

                    if (
                        rect_height
                        <=
                        THIN_RECT_MAX_THICKNESS
                        and
                        rect_width
                        >=
                        MIN_LINE_LENGTH
                    ):

                        center_y = (
                            y0 + y1
                        ) / 2

                        line = {
                            "id":
                                f"line_{line_counter}",

                            "type":
                                "line",

                            "source":
                                "thin_rectangle",

                            "x0":
                                round_num(
                                    min(
                                        x0,
                                        x1
                                    )
                                ),

                            "y0":
                                round_num(
                                    center_y
                                ),

                            "x1":
                                round_num(
                                    max(
                                        x0,
                                        x1
                                    )
                                ),

                            "y1":
                                round_num(
                                    center_y
                                ),

                            "orientation":
                                "horizontal",

                            "length":
                                round_num(
                                    rect_width
                                ),

                            "thickness":
                                round_num(
                                    rect_height
                                )
                        }

                        line_counter += 1

                        lines_out.append(
                            line
                        )

                        drawing_items.append(
                            line
                        )

                        continue

                    # -----------------------------------------
                    # THIN VERTICAL RECTANGLE
                    # -----------------------------------------

                    if (
                        rect_width
                        <=
                        THIN_RECT_MAX_THICKNESS
                        and
                        rect_height
                        >=
                        MIN_LINE_LENGTH
                    ):

                        center_x = (
                            x0 + x1
                        ) / 2

                        line = {
                            "id":
                                f"line_{line_counter}",

                            "type":
                                "line",

                            "source":
                                "thin_rectangle",

                            "x0":
                                round_num(
                                    center_x
                                ),

                            "y0":
                                round_num(
                                    min(
                                        y0,
                                        y1
                                    )
                                ),

                            "x1":
                                round_num(
                                    center_x
                                ),

                            "y1":
                                round_num(
                                    max(
                                        y0,
                                        y1
                                    )
                                ),

                            "orientation":
                                "vertical",

                            "length":
                                round_num(
                                    rect_height
                                ),

                            "thickness":
                                round_num(
                                    rect_width
                                )
                        }

                        line_counter += 1

                        lines_out.append(
                            line
                        )

                        drawing_items.append(
                            line
                        )

                        continue

                    # -----------------------------------------
                    # NORMAL RECTANGLE
                    # -----------------------------------------

                    rect_data = {
                        "id":
                            f"rect_{rect_counter}",

                        "type":
                            "rectangle",

                        "bbox":
                            bbox_dict(
                                (
                                    x0,
                                    y0,
                                    x1,
                                    y1
                                )
                            ),

                        "width":
                            round_num(
                                rect_width
                            ),

                        "height":
                            round_num(
                                rect_height
                            )
                    }

                    rect_counter += 1

                    rectangles_out.append(
                        rect_data
                    )

                    drawing_items.append(
                        rect_data
                    )

                except Exception as e:
                    print(
                        "Rectangle parse error:",
                        str(e)
                    )

        drawing_data[
            "items"
        ] = drawing_items

        if drawing_items:
            drawings_out.append(
                drawing_data
            )

    return (
        drawings_out,
        lines_out,
        rectangles_out
    )


# ============================================================
# BASIC FORM LINE DETECTION
# ============================================================

def detect_form_lines(
    lines,
    page_width
):
    result = []

    minimum = (
        page_width *
        FORM_MIN_LINE_RATIO
    )

    maximum = (
        page_width *
        FORM_MAX_LINE_RATIO
    )

    for line in lines:

        if (
            line.get(
                "orientation"
            )
            !=
            "horizontal"
        ):
            continue

        length = float(
            line.get(
                "length",
                0
            )
        )

        if (
            length >= minimum
            and
            length <= maximum
        ):
            result.append(
                line
            )

    return result


# ============================================================
# FORM STRUCTURE DETECTION
# ============================================================

def find_vertical_boundaries(
    field_line,
    vertical_lines
):
    """
    Bir yatay form çizgisinin x0/x1 uçlarında veya
    çizgi boyunca dikey çizgiler var mı diye bakar.
    """

    x0 = float(
        field_line.get(
            "x0",
            0
        )
    )

    x1 = float(
        field_line.get(
            "x1",
            0
        )
    )

    y = float(
        field_line.get(
            "y0",
            0
        )
    )

    left_border = False
    right_border = False

    internal_borders = []

    for line in vertical_lines:

        vx = float(
            line.get(
                "x0",
                0
            )
        )

        vy0 = min(
            float(
                line.get(
                    "y0",
                    0
                )
            ),
            float(
                line.get(
                    "y1",
                    0
                )
            )
        )

        vy1 = max(
            float(
                line.get(
                    "y0",
                    0
                )
            ),
            float(
                line.get(
                    "y1",
                    0
                )
            )
        )

        # Dikey çizgi bu satır seviyesini kesiyor mu?
        if not (
            vy0 -
            FORM_VERTICAL_BORDER_TOLERANCE
            <=
            y
            <=
            vy1 +
            FORM_VERTICAL_BORDER_TOLERANCE
        ):
            continue

        if (
            abs(
                vx - x0
            )
            <=
            FORM_VERTICAL_BORDER_TOLERANCE
        ):
            left_border = True

        elif (
            abs(
                vx - x1
            )
            <=
            FORM_VERTICAL_BORDER_TOLERANCE
        ):
            right_border = True

        elif (
            x0 <
            vx <
            x1
        ):
            internal_borders.append(
                round_num(
                    vx
                )
            )

    internal_borders = sorted(
        list(
            set(
                internal_borders
            )
        )
    )

    return {
        "left_border":
            left_border,

        "right_border":
            right_border,

        "internal_borders":
            internal_borders
    }


def find_form_label(
    field_line,
    raw_blocks
):
    """
    Form çizgisinin hemen üstündeki en uygun metni bulur.
    """

    fx0 = float(
        field_line.get(
            "x0",
            0
        )
    )

    fx1 = float(
        field_line.get(
            "x1",
            0
        )
    )

    fy = float(
        field_line.get(
            "y0",
            0
        )
    )

    candidates = []

    for block in raw_blocks:

        text = clean_text(
            block.get(
                "text",
                ""
            )
        )

        if not text:
            continue

        bbox = block.get(
            "bbox",
            {}
        )

        bx0 = float(
            bbox.get(
                "x0",
                0
            )
        )

        by0 = float(
            bbox.get(
                "y0",
                0
            )
        )

        bx1 = float(
            bbox.get(
                "x1",
                0
            )
        )

        by1 = float(
            bbox.get(
                "y1",
                0
            )
        )

        # Form çizgisinin esas olarak üzerindeki
        # yazıları arıyoruz.
        vertical_distance = (
            fy - by1
        )

        if (
            vertical_distance
            <
            -FORM_LABEL_MAX_BELOW_PT
            or
            vertical_distance
            >
            FORM_LABEL_MAX_ABOVE_PT
        ):
            continue

        overlap = horizontal_overlap(
            fx0 -
            FORM_LABEL_X_TOLERANCE,
            fx1 +
            FORM_LABEL_X_TOLERANCE,
            bx0,
            bx1
        )

        block_width = max(
            1,
            bx1 - bx0
        )

        overlap_ratio = (
            overlap /
            block_width
        )

        # Metin ile form çizgisi hiç yatay ilişki
        # göstermiyorsa kullanma.
        if overlap <= 0:
            continue

        # Skor küçüldükçe daha iyi.
        center_field = (
            fx0 + fx1
        ) / 2

        center_text = (
            bx0 + bx1
        ) / 2

        center_distance = abs(
            center_field -
            center_text
        )

        score = (
            abs(
                vertical_distance
            ) * 4
            +
            center_distance
            -
            overlap_ratio * 20
        )

        candidates.append({
            "score":
                score,

            "text":
                text,

            "bbox":
                block[
                    "bbox"
                ]
        })

    if not candidates:
        return {
            "text": "",
            "bbox": None
        }

    candidates.sort(
        key=lambda item: item[
            "score"
        ]
    )

    best = candidates[0]

    return {
        "text":
            best[
                "text"
            ],

        "bbox":
            best[
                "bbox"
            ]
    }


def detect_field_type(
    label,
    line,
    vertical_info
):
    label_lower = (
        label.lower()
    )

    # checkbox / yes-no benzeri alanlar
    if (
        "ja" in label_lower
        and
        "nein" in label_lower
    ):
        return "choice"

    if (
        "yes" in label_lower
        and
        "no" in label_lower
    ):
        return "choice"

    length = float(
        line.get(
            "length",
            0
        )
    )

    if (
        vertical_info.get(
            "left_border"
        )
        and
        vertical_info.get(
            "right_border"
        )
    ):
        return "boxed_field"

    if length < 80:
        return "short_text_field"

    return "text_field"


def cluster_form_rows(
    fields
):
    if not fields:
        return []

    ordered = sorted(
        fields,
        key=lambda item: (
            item[
                "y"
            ],
            item[
                "x0"
            ]
        )
    )

    row_groups = []

    for field in ordered:

        placed = False

        for group in row_groups:

            if (
                abs(
                    field[
                        "y"
                    ]
                    -
                    group[
                        "average_y"
                    ]
                )
                <=
                FORM_ROW_Y_TOLERANCE
            ):
                group[
                    "fields"
                ].append(
                    field
                )

                group[
                    "average_y"
                ] = (
                    sum(
                        item[
                            "y"
                        ]
                        for item in group[
                            "fields"
                        ]
                    )
                    /
                    len(
                        group[
                            "fields"
                        ]
                    )
                )

                placed = True
                break

        if not placed:
            row_groups.append({
                "average_y":
                    field[
                        "y"
                    ],

                "fields":
                    [
                        field
                    ]
            })

    result = []

    for index, group in enumerate(
        row_groups,
        start=1
    ):

        group_fields = sorted(
            group[
                "fields"
            ],
            key=lambda item: item[
                "x0"
            ]
        )

        result.append({
            "row_id":
                f"form_row_{index}",

            "y":
                round_num(
                    group[
                        "average_y"
                    ]
                ),

            "field_count":
                len(
                    group_fields
                ),

            "field_ids":
                [
                    field[
                        "field_id"
                    ]
                    for field in group_fields
                ]
        })

    return result


def cluster_form_columns(
    fields
):
    if not fields:
        return []

    groups = []

    ordered = sorted(
        fields,
        key=lambda item: item[
            "x0"
        ]
    )

    for field in ordered:

        placed = False

        for group in groups:

            if (
                abs(
                    field[
                        "x0"
                    ]
                    -
                    group[
                        "average_x"
                    ]
                )
                <=
                FORM_COLUMN_X_TOLERANCE
            ):

                group[
                    "fields"
                ].append(
                    field
                )

                group[
                    "average_x"
                ] = (
                    sum(
                        item[
                            "x0"
                        ]
                        for item in group[
                            "fields"
                        ]
                    )
                    /
                    len(
                        group[
                            "fields"
                        ]
                    )
                )

                placed = True
                break

        if not placed:
            groups.append({
                "average_x":
                    field[
                        "x0"
                    ],

                "fields":
                    [
                        field
                    ]
            })

    result = []

    for index, group in enumerate(
        groups,
        start=1
    ):

        result.append({
            "column_id":
                f"form_column_{index}",

            "x":
                round_num(
                    group[
                        "average_x"
                    ]
                ),

            "field_count":
                len(
                    group[
                        "fields"
                    ]
                ),

            "field_ids":
                [
                    field[
                        "field_id"
                    ]
                    for field in group[
                        "fields"
                    ]
                ]
        })

    return result


def detect_form_structure(
    form_lines,
    vertical_lines,
    raw_blocks
):
    """
    Her horizontal form line = potansiyel giriş alanı.

    Daha sonra:
    - label eşleştirir
    - dikey border kontrol eder
    - satırlara gruplar
    - kolonlara gruplar
    """

    if not ENABLE_FORM_DETECTION:
        return {
            "fields": [],
            "rows": [],
            "columns": []
        }

    fields = []

    field_counter = 1

    ordered_lines = sorted(
        form_lines,
        key=lambda line: (
            float(
                line.get(
                    "y0",
                    0
                )
            ),
            float(
                line.get(
                    "x0",
                    0
                )
            )
        )
    )

    for line in ordered_lines:

        x0 = min(
            float(
                line.get(
                    "x0",
                    0
                )
            ),
            float(
                line.get(
                    "x1",
                    0
                )
            )
        )

        x1 = max(
            float(
                line.get(
                    "x0",
                    0
                )
            ),
            float(
                line.get(
                    "x1",
                    0
                )
            )
        )

        y = float(
            line.get(
                "y0",
                0
            )
        )

        width = (
            x1 - x0
        )

        if width < FORM_MIN_FIELD_WIDTH:
            continue

        label_info = find_form_label(
            line,
            raw_blocks
        )

        vertical_info = (
            find_vertical_boundaries(
                line,
                vertical_lines
            )
        )

        field_type = detect_field_type(
            label_info[
                "text"
            ],
            line,
            vertical_info
        )

        field = {
            "field_id":
                f"form_{field_counter}",

            "type":
                field_type,

            "x0":
                round_num(
                    x0
                ),

            "y":
                round_num(
                    y
                ),

            "x1":
                round_num(
                    x1
                ),

            "width":
                round_num(
                    width
                ),

            "line_id":
                line.get(
                    "id"
                ),

            "line_source":
                line.get(
                    "source"
                ),

            "label":
                label_info[
                    "text"
                ],

            "label_bbox":
                label_info[
                    "bbox"
                ],

            "left_border":
                vertical_info[
                    "left_border"
                ],

            "right_border":
                vertical_info[
                    "right_border"
                ],

            "internal_borders":
                vertical_info[
                    "internal_borders"
                ]
        }

        fields.append(
            field
        )

        field_counter += 1

    rows = cluster_form_rows(
        fields
    )

    columns = cluster_form_columns(
        fields
    )

    # field objelerine row / column id ekle
    field_lookup = {
        field[
            "field_id"
        ]:
        field
        for field in fields
    }

    for row in rows:
        for field_id in row[
            "field_ids"
        ]:

            if field_id in field_lookup:
                field_lookup[
                    field_id
                ][
                    "row_id"
                ] = row[
                    "row_id"
                ]

    for column in columns:
        for field_id in column[
            "field_ids"
        ]:

            if field_id in field_lookup:
                field_lookup[
                    field_id
                ][
                    "column_id"
                ] = column[
                    "column_id"
                ]

    return {
        "fields":
            fields,

        "rows":
            rows,

        "columns":
            columns
    }


# ============================================================
# TABLE DETECTION
# ============================================================

def find_page_tables(page):
    if not ENABLE_TABLE_DETECTION:
        return []

    try:
        finder = (
            page.find_tables()
        )

        if finder is None:
            return []

        return list(
            finder.tables
        )

    except Exception as e:
        print(
            "Table detection error:",
            str(e)
        )

        return []


def clean_table_matrix(matrix):
    if not matrix:
        return []

    cleaned = []
    max_columns = 0

    for row in matrix:

        if row is None:
            row = []

        clean_row = [
            clean_text(
                value
            )
            for value in row
        ]

        max_columns = max(
            max_columns,
            len(
                clean_row
            )
        )

        cleaned.append(
            clean_row
        )

    if max_columns == 0:
        return []

    for row in cleaned:
        while (
            len(
                row
            )
            <
            max_columns
        ):
            row.append("")

    return [
        row
        for row in cleaned
        if any(
            value.strip()
            for value in row
        )
    ]


def extract_table_cell_bboxes(
    detected_table
):
    result = []

    try:
        for row in detected_table.rows:

            row_result = []

            for cell in row.cells:

                if cell is None:
                    row_result.append(
                        None
                    )

                else:
                    row_result.append(
                        bbox_dict(
                            cell
                        )
                    )

            result.append(
                row_result
            )

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
            flags=fitz.TEXT_PRESERVE_WHITESPACE
        )

    except Exception:
        text_dict = (
            page.get_text(
                "dict"
            )
        )

    blocks_out = []

    for raw_block in text_dict.get(
        "blocks",
        []
    ):

        if raw_block.get(
            "type"
        ) != 0:
            continue

        block_bbox = raw_block.get(
            "bbox",
            (0, 0, 0, 0)
        )

        lines_out = []
        all_spans = []
        text_lines = []

        for raw_line in raw_block.get(
            "lines",
            []
        ):

            line_spans = []

            for raw_span in raw_line.get(
                "spans",
                []
            ):

                span = normalize_span(
                    raw_span
                )

                if not span[
                    "text"
                ]:
                    continue

                line_spans.append(
                    span
                )

                all_spans.append(
                    span
                )

            if not line_spans:
                continue

            line_text = "".join(
                span[
                    "text"
                ]
                for span in line_spans
            ).strip()

            if not line_text:
                continue

            line_bbox = raw_line.get(
                "bbox",
                block_bbox
            )

            lines_out.append({
                "text":
                    line_text,

                "bbox":
                    bbox_dict(
                        line_bbox
                    ),

                "spans":
                    line_spans
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
            "text":
                text,

            "bbox":
                bbox_dict(
                    block_bbox
                ),

            "lines":
                lines_out,

            "spans":
                all_spans,

            "style":
                get_block_style(
                    all_spans
                )
        })

    return blocks_out


# ============================================================
# FONT STATS / HEADINGS
# ============================================================

def get_page_font_statistics(
    blocks
):
    sizes = []

    for block in blocks:

        for span in block.get(
            "spans",
            []
        ):

            size = float(
                span.get(
                    "size",
                    0
                )
            )

            if size > 0:
                sizes.append(
                    size
                )

    if not sizes:
        return {
            "median_size":
                DEFAULT_FONT_SIZE,

            "max_size":
                DEFAULT_FONT_SIZE
        }

    return {
        "median_size":
            round_num(
                statistics.median(
                    sizes
                )
            ),

        "max_size":
            round_num(
                max(
                    sizes
                )
            )
    }


def classify_text_role(
    block,
    font_stats
):
    text = clean_text(
        block.get(
            "text",
            ""
        )
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

    short = (
        len(
            text
        )
        <=
        120
    )

    if (
        short
        and
        max_size >= (
            median_size *
            HEADING_LARGE_RATIO
        )
    ):
        return "heading_1"

    if (
        short
        and
        bold
        and
        max_size >= (
            median_size *
            HEADING_SIZE_RATIO
        )
    ):
        return "heading_2"

    if (
        short
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

def detect_columns(
    blocks,
    page_width
):
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

        width = (
            x1 - x0
        )

        if (
            width > 10
            and
            width <
            page_width * 0.70
        ):
            candidates.append(
                (
                    x0,
                    x1
                )
            )

    if (
        len(
            candidates
        )
        <
        COLUMN_MIN_BLOCKS * 2
    ):
        return {
            "column_count": 1,
            "divider_x": None
        }

    center = (
        page_width / 2
    )

    left = [
        item
        for item in candidates
        if item[0] < center
    ]

    right = [
        item
        for item in candidates
        if item[0] >= (
            center * 0.85
        )
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

    left_edge = (
        statistics.median(
            item[1]
            for item in left
        )
    )

    right_edge = (
        statistics.median(
            item[0]
            for item in right
        )
    )

    gap = (
        right_edge -
        left_edge
    )

    if (
        gap >= (
            page_width *
            COLUMN_GAP_RATIO
        )
    ):

        return {
            "column_count": 2,

            "divider_x":
                round_num(
                    (
                        left_edge +
                        right_edge
                    )
                    / 2
                )
        }

    return {
        "column_count": 1,
        "divider_x": None
    }


# ============================================================
# UNDERLINE DETECTION
# ============================================================

def detect_span_underlines(
    spans,
    horizontal_lines
):
    for span in spans:

        span[
            "underline"
        ] = False

        bbox = span.get(
            "bbox",
            {}
        )

        x0 = float(
            bbox.get(
                "x0",
                0
            )
        )

        y1 = float(
            bbox.get(
                "y1",
                0
            )
        )

        x1 = float(
            bbox.get(
                "x1",
                0
            )
        )

        text_width = max(
            1,
            x1 - x0
        )

        for line in horizontal_lines:

            ly = float(
                line.get(
                    "y0",
                    0
                )
            )

            lx0 = min(
                float(
                    line.get(
                        "x0",
                        0
                    )
                ),
                float(
                    line.get(
                        "x1",
                        0
                    )
                )
            )

            lx1 = max(
                float(
                    line.get(
                        "x0",
                        0
                    )
                ),
                float(
                    line.get(
                        "x1",
                        0
                    )
                )
            )

            vertical_distance = abs(
                ly - y1
            )

            overlap = max(
                0,
                min(
                    x1,
                    lx1
                )
                -
                max(
                    x0,
                    lx0
                )
            )

            ratio = (
                overlap /
                text_width
            )

            if (
                vertical_distance <= 3
                and
                ratio >= 0.50
            ):

                span[
                    "underline"
                ] = True

                break


# ============================================================
# MAIN STRUCTURE EXTRACTION
# ============================================================

def extract_pdf_structure(
    document
):
    pages = []
    elements = []

    element_counter = 1
    table_counter = 1

    for page_index in range(
        len(
            document
        )
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

        raw_blocks = (
            extract_raw_text_blocks(
                page
            )
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

        font_stats = (
            get_page_font_statistics(
                raw_blocks
            )
        )

        # ====================================================
        # FORM STRUCTURE
        # ====================================================

        form_structure = (
            detect_form_structure(
                form_lines,
                vertical_lines,
                raw_blocks
            )
        )

        form_fields = (
            form_structure[
                "fields"
            ]
        )

        form_rows = (
            form_structure[
                "rows"
            ]
        )

        form_columns = (
            form_structure[
                "columns"
            ]
        )

        # ====================================================
        # TABLES
        # ====================================================

        detected_tables = (
            find_page_tables(
                page
            )
        )

        table_regions = []
        table_structures = []

        for detected_table in detected_tables:

            try:

                bbox = (
                    detected_table.bbox
                )

                table_bbox = (
                    float(
                        bbox[0]
                    ),
                    float(
                        bbox[1]
                    ),
                    float(
                        bbox[2]
                    ),
                    float(
                        bbox[3]
                    )
                )

                matrix = clean_table_matrix(
                    detected_table.extract()
                )

            except Exception as e:

                print(
                    "Could not parse table:",
                    str(e)
                )

                continue

            if not matrix:
                continue

            row_count = len(
                matrix
            )

            column_count = max(
                len(
                    row
                )
                for row in matrix
            )

            if (
                row_count <= 0
                or
                column_count <= 0
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
                "table_id":
                    table_id,

                "type":
                    "table",

                "page":
                    page_number,

                "bbox":
                    bbox_dict(
                        table_bbox
                    ),

                "row_count":
                    row_count,

                "column_count":
                    column_count,

                "rows":
                    []
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
                        row_index < len(
                            matrix
                        )
                        and
                        col_index < len(
                            matrix[
                                row_index
                            ]
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
                            cell_bboxes[
                                row_index
                            ]
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
                            "id":
                                element_id,

                            "page":
                                page_number,

                            "type":
                                "table_cell",

                            "table_id":
                                table_id,

                            "row":
                                row_index,

                            "col":
                                col_index,

                            "text":
                                value,

                            "bbox":
                                cell_bbox
                        }

                        elements.append(
                            element
                        )

                    else:
                        element_id = None

                    row_data.append({
                        "id":
                            element_id,

                        "row":
                            row_index,

                        "col":
                            col_index,

                        "text":
                            value,

                        "bbox":
                            cell_bbox
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
        # NORMAL TEXT
        # ====================================================

        text_elements = []
        detected_lists = []

        for block in raw_blocks:

            bbox = block.get(
                "bbox",
                {}
            )

            block_rect = (
                float(
                    bbox.get(
                        "x0",
                        0
                    )
                ),

                float(
                    bbox.get(
                        "y0",
                        0
                    )
                ),

                float(
                    bbox.get(
                        "x1",
                        0
                    )
                ),

                float(
                    bbox.get(
                        "y1",
                        0
                    )
                )
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

            role = classify_text_role(
                block,
                font_stats
            )

            list_info = detect_list_info(
                text
            )

            element_type = (
                "list_item"
                if list_info
                else
                "text_block"
            )

            element = {
                "id":
                    element_id,

                "page":
                    page_number,

                "type":
                    element_type,

                "text":
                    text,

                "bbox":
                    block[
                        "bbox"
                    ],

                "role":
                    role,

                "style":
                    block.get(
                        "style",
                        {}
                    ),

                "lines":
                    block.get(
                        "lines",
                        []
                    ),

                "spans":
                    block.get(
                        "spans",
                        []
                    ),

                "list":
                    list_info
            }

            elements.append(
                element
            )

            text_elements.append(
                element
            )

            if list_info:

                detected_lists.append({
                    "id":
                        element_id,

                    "bbox":
                        block[
                            "bbox"
                        ],

                    "list_type":
                        list_info[
                            "list_type"
                        ],

                    "marker":
                        list_info[
                            "marker"
                        ],

                    "text":
                        text
                })

        # ====================================================
        # COLUMNS
        # ====================================================

        columns = detect_columns(
            raw_blocks,
            page_width
        )

        # ====================================================
        # LAYOUT
        # ====================================================

        layout_items = []

        for element in text_elements:

            layout_items.append({
                "type":
                    element[
                        "type"
                    ],

                "id":
                    element[
                        "id"
                    ],

                "x0":
                    element[
                        "bbox"
                    ][
                        "x0"
                    ],

                "y0":
                    element[
                        "bbox"
                    ][
                        "y0"
                    ]
            })

        for table in table_structures:

            layout_items.append({
                "type":
                    "table",

                "table_id":
                    table[
                        "table_id"
                    ],

                "x0":
                    table[
                        "bbox"
                    ][
                        "x0"
                    ],

                "y0":
                    table[
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
            "page":
                page_number,

            "width":
                round_num(
                    page_width
                ),

            "height":
                round_num(
                    page_height
                ),

            "font_statistics":
                font_stats,

            "column_count":
                columns[
                    "column_count"
                ],

            "column_divider_x":
                columns[
                    "divider_x"
                ],

            "text_block_count":
                len(
                    text_elements
                ),

            "table_count":
                len(
                    table_structures
                ),

            "list_count":
                len(
                    detected_lists
                ),

            "drawing_count":
                len(
                    drawings
                ),

            "line_count":
                len(
                    lines
                ),

            "horizontal_line_count":
                len(
                    horizontal_lines
                ),

            "vertical_line_count":
                len(
                    vertical_lines
                ),

            "rectangle_count":
                len(
                    rectangles
                ),

            "form_line_count":
                len(
                    form_lines
                ),

            # NEW
            "form_field_count":
                len(
                    form_fields
                ),

            "form_row_count":
                len(
                    form_rows
                ),

            "form_column_count":
                len(
                    form_columns
                ),

            "tables":
                table_structures,

            "lists":
                detected_lists,

            "drawings":
                drawings,

            "lines":
                lines,

            "rectangles":
                rectangles,

            "form_lines":
                form_lines,

            # NEW
            "form_fields":
                form_fields,

            "form_rows":
                form_rows,

            "form_columns":
                form_columns,

            "layout":
                layout_items
        })

    # ========================================================
    # TRANSLATION CHUNKS
    # ========================================================

    elements.sort(
        key=lambda item: int(
            item[
                "id"
            ].split(
                "_"
            )[1]
        )
    )

    translation_elements = []

    for element in elements:

        translation_elements.append({
            "id":
                element[
                    "id"
                ],

            "type":
                element[
                    "type"
                ],

            "text":
                element[
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

def parse_translations(
    raw_value
):
    if raw_value is None:
        return {}

    if isinstance(
        raw_value,
        bytes
    ):
        raw_value = (
            raw_value.decode(
                "utf-8"
            )
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

    result = {}

    for item in data.get(
        "translations",
        []
    ):

        element_id = item.get(
            "id"
        )

        if not element_id:
            continue

        text = item.get(
            "text"
        )

        if text is None:
            text = ""

        result[
            str(
                element_id
            )
        ] = str(
            text
        )

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

    tc_pr = (
        tc.get_or_add_tcPr()
    )

    tc_mar = (
        tc_pr.first_child_found_in(
            "w:tcMar"
        )
    )

    if tc_mar is None:

        tc_mar = OxmlElement(
            "w:tcMar"
        )

        tc_pr.append(
            tc_mar
        )

    for name, value in [
        ("top", top),
        ("start", start),
        ("bottom", bottom),
        ("end", end)
    ]:

        node = tc_mar.find(
            qn(
                f"w:{name}"
            )
        )

        if node is None:

            node = OxmlElement(
                f"w:{name}"
            )

            tc_mar.append(
                node
            )

        node.set(
            qn(
                "w:w"
            ),
            str(
                value
            )
        )

        node.set(
            qn(
                "w:type"
            ),
            "dxa"
        )


def prevent_row_split(
    row
):
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
            qn(
                "w:val"
            ),
            "1"
        )

    except Exception:
        pass


def set_table_borders(
    table
):
    try:

        tbl_pr = (
            table._tbl.tblPr
        )

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

            tag = (
                f"w:{edge}"
            )

            border = borders.find(
                qn(
                    tag
                )
            )

            if border is None:

                border = OxmlElement(
                    tag
                )

                borders.append(
                    border
                )

            border.set(
                qn(
                    "w:val"
                ),
                "single"
            )

            border.set(
                qn(
                    "w:sz"
                ),
                "4"
            )

            border.set(
                qn(
                    "w:space"
                ),
                "0"
            )

            border.set(
                qn(
                    "w:color"
                ),
                "808080"
            )

    except Exception:
        pass


# ============================================================
# BASIC DOCX
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
        bbox[
            "x0"
        ]
    )

    y0 = float(
        bbox[
            "y0"
        ]
    )

    x1 = float(
        bbox[
            "x1"
        ]
    )

    y1 = float(
        bbox[
            "y1"
        ]
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
        gap = y0
    else:
        gap = (
            y0 -
            previous_bottom
        )

    gap = clamp(
        gap,
        MIN_VERTICAL_GAP_PT,
        MAX_VERTICAL_GAP_PT
    )

    paragraph.paragraph_format.space_before = Pt(
        gap
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

    run.font.size = Pt(
        float(
            style.get(
                "size",
                DEFAULT_FONT_SIZE
            )
            or
            DEFAULT_FONT_SIZE
        )
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

    if (
        rows <= 0
        or
        cols <= 0
    ):
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

    rows_data = table_data.get(
        "rows",
        []
    )

    for row_index in range(
        rows
    ):

        row = table.rows[
            row_index
        ]

        prevent_row_split(
            row
        )

        for col_index in range(
            cols
        ):

            cell = row.cells[
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
                    rows_data
                )
                and
                col_index < len(
                    rows_data[
                        row_index
                    ]
                )
            ):
                info = (
                    rows_data[
                        row_index
                    ][
                        col_index
                    ]
                )

            if not info:
                continue

            element_id = info.get(
                "id"
            )

            original = info.get(
                "text",
                ""
            )

            if element_id:

                final_text = (
                    translations.get(
                        element_id,
                        original
                    )
                )

            else:
                final_text = ""

            paragraph = (
                cell.paragraphs[0]
            )

            paragraph.paragraph_format.space_before = Pt(
                0
            )

            paragraph.paragraph_format.space_after = Pt(
                0
            )

            run = paragraph.add_run(
                final_text
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
        element[
            "id"
        ]:
        element

        for element in elements
    }

    if document.paragraphs:

        paragraph = (
            document.paragraphs[0]
        )

        paragraph.text = ""

        paragraph.paragraph_format.space_before = Pt(
            0
        )

        paragraph.paragraph_format.space_after = Pt(
            0
        )

    for page_index, page_data in enumerate(
        pages
    ):

        page_width = float(
            page_data[
                "width"
            ]
        )

        page_height = float(
            page_data[
                "height"
            ]
        )

        if page_index == 0:

            section = (
                document.sections[0]
            )

        else:

            section = (
                document.add_section(
                    WD_SECTION.NEW_PAGE
                )
            )

        configure_section(
            section,
            page_width,
            page_height
        )

        table_lookup = {
            table[
                "table_id"
            ]:
            table

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

            if (
                item.get(
                    "type"
                )
                ==
                "table"
            ):

                table_data = (
                    table_lookup.get(
                        item.get(
                            "table_id"
                        )
                    )
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

            element = (
                element_lookup.get(
                    element_id
                )
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
                    page_width,
                    previous_bottom
                )
            )

    output = BytesIO()

    document.save(
        output
    )

    output.seek(
        0
    )

    return output


# ============================================================
# EXTRACT PDF
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
        request.files[
            "file"
        ]
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

        return jsonify({
            "filename":
                uploaded_file.filename,

            "page_count":
                len(
                    pages
                ),

            "element_count":
                len(
                    elements
                ),

            "chunk_size":
                CHUNK_SIZE,

            "chunk_count":
                len(
                    chunks
                ),

            "table_count":
                sum(
                    page[
                        "table_count"
                    ]
                    for page in pages
                ),

            "list_count":
                sum(
                    page[
                        "list_count"
                    ]
                    for page in pages
                ),

            "line_count":
                sum(
                    page[
                        "line_count"
                    ]
                    for page in pages
                ),

            "form_line_count":
                sum(
                    page[
                        "form_line_count"
                    ]
                    for page in pages
                ),

            # NEW
            "form_field_count":
                sum(
                    page[
                        "form_field_count"
                    ]
                    for page in pages
                ),

            "form_row_count":
                sum(
                    page[
                        "form_row_count"
                    ]
                    for page in pages
                ),

            "rectangle_count":
                sum(
                    page[
                        "rectangle_count"
                    ]
                    for page in pages
                ),

            "pages":
                pages,

            "elements":
                elements,

            "chunks":
                chunks
        })

    finally:

        document.close()


# ============================================================
# CREATE DOCX
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
        request.files[
            "file"
        ]
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