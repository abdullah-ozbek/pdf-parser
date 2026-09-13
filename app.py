from flask import Flask, request, jsonify, send_file
import fitz
import json

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


# ============================================================
# ROOT
# ============================================================

@app.route("/", methods=["GET"])
def home():
    return {
        "status": "ok",
        "service": "PDF Parser"
    }


# ============================================================
# GENERAL HELPERS
# ============================================================

def clamp(value, minimum, maximum):
    return max(
        minimum,
        min(
            maximum,
            value
        )
    )


def rects_intersect(rect1, rect2):
    """
    rect = (x0, y0, x1, y1)
    """

    return not (
        rect1[2] <= rect2[0]
        or rect1[0] >= rect2[2]
        or rect1[3] <= rect2[1]
        or rect1[1] >= rect2[3]
    )


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


# ============================================================
# TABLE HELPERS
# ============================================================

def find_page_tables(page):

    if not ENABLE_TABLE_DETECTION:
        return []

    try:
        finder = page.find_tables()

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

        cleaned_row = []

        for value in row:

            if value is None:
                value = ""

            value = str(
                value
            ).strip()

            cleaned_row.append(
                value
            )

        max_columns = max(
            max_columns,
            len(cleaned_row)
        )

        cleaned.append(
            cleaned_row
        )

    if max_columns == 0:
        return []

    # Normalize row lengths
    for row in cleaned:

        while len(row) < max_columns:
            row.append("")

    # Remove completely empty rows
    cleaned = [
        row
        for row in cleaned
        if any(
            value.strip()
            for value in row
        )
    ]

    if not cleaned:
        return []

    # Remove columns that are completely empty
    used_columns = []

    for col_index in range(
        max_columns
    ):

        has_content = False

        for row in cleaned:

            if (
                col_index < len(row)
                and
                row[col_index].strip()
            ):
                has_content = True
                break

        if has_content:
            used_columns.append(
                col_index
            )

    if not used_columns:
        return []

    final_matrix = []

    for row in cleaned:

        final_matrix.append([
            row[col]
            if col < len(row)
            else ""
            for col in used_columns
        ])

    return final_matrix


def calculate_column_widths_from_table(
    table_bbox,
    column_count
):

    if column_count <= 0:
        return []

    x0, y0, x1, y1 = table_bbox

    total_width = max(
        1.0,
        x1 - x0
    )

    equal_width = (
        total_width /
        column_count
    )

    return [
        equal_width
        for _ in range(
            column_count
        )
    ]


# ============================================================
# PDF EXTRACTION
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

        page_width = float(
            page.rect.width
        )

        page_height = float(
            page.rect.height
        )

        page_elements = []

        # ====================================================
        # DETECT TABLES
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

                matrix = (
                    clean_table_matrix(
                        raw_matrix
                    )
                )

            except Exception as e:

                print(
                    "Could not read table:",
                    str(e)
                )

                continue

            if not matrix:
                continue

            row_count = len(
                matrix
            )

            column_count = max(
                len(row)
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

            column_widths = (
                calculate_column_widths_from_table(
                    table_bbox,
                    column_count
                )
            )

            table_data = {
                "table_id":
                    table_id,

                "page":
                    page_index + 1,

                "type":
                    "table",

                "bbox": {
                    "x0":
                        round(
                            table_bbox[0],
                            2
                        ),

                    "y0":
                        round(
                            table_bbox[1],
                            2
                        ),

                    "x1":
                        round(
                            table_bbox[2],
                            2
                        ),

                    "y1":
                        round(
                            table_bbox[3],
                            2
                        )
                },

                "row_count":
                    row_count,

                "column_count":
                    column_count,

                "column_widths":
                    column_widths,

                "rows":
                    []
            }

            for row_index, row in enumerate(
                matrix
            ):

                row_data = []

                for col_index in range(
                    column_count
                ):

                    text = ""

                    if col_index < len(row):
                        text = (
                            row[col_index]
                            or ""
                        ).strip()

                    if not text:

                        row_data.append({
                            "id": None,
                            "text": "",
                            "row": row_index,
                            "col": col_index
                        })

                        continue

                    element_id = (
                        f"pdf_{element_counter}"
                    )

                    element_counter += 1

                    element = {
                        "id":
                            element_id,

                        "page":
                            page_index + 1,

                        "type":
                            "table_cell",

                        "table_id":
                            table_id,

                        "row":
                            row_index,

                        "col":
                            col_index,

                        "text":
                            text,

                        "bbox": {
                            "x0":
                                round(
                                    table_bbox[0],
                                    2
                                ),

                            "y0":
                                round(
                                    table_bbox[1],
                                    2
                                ),

                            "x1":
                                round(
                                    table_bbox[2],
                                    2
                                ),

                            "y1":
                                round(
                                    table_bbox[3],
                                    2
                                )
                        }
                    }

                    elements.append(
                        element
                    )

                    row_data.append({
                        "id":
                            element_id,

                        "text":
                            text,

                        "row":
                            row_index,

                        "col":
                            col_index
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
        # NORMAL TEXT BLOCKS
        # ====================================================

        blocks = page.get_text(
            "blocks"
        )

        normal_blocks = []

        for block in blocks:

            if len(block) < 5:
                continue

            x0 = float(
                block[0]
            )

            y0 = float(
                block[1]
            )

            x1 = float(
                block[2]
            )

            y1 = float(
                block[3]
            )

            text = str(
                block[4]
            ).strip()

            if not text:
                continue

            block_rect = (
                x0,
                y0,
                x1,
                y1
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

            normal_blocks.append({
                "bbox":
                    block_rect,

                "text":
                    text
            })

        # ====================================================
        # BUILD PAGE ORDER
        # ====================================================

        ordered_items = []

        for block in normal_blocks:

            ordered_items.append({
                "kind":
                    "text",

                "y0":
                    block["bbox"][1],

                "x0":
                    block["bbox"][0],

                "data":
                    block
            })

        for table_data in table_structures:

            ordered_items.append({
                "kind":
                    "table",

                "y0":
                    table_data[
                        "bbox"
                    ][
                        "y0"
                    ],

                "x0":
                    table_data[
                        "bbox"
                    ][
                        "x0"
                    ],

                "data":
                    table_data
            })

        ordered_items.sort(
            key=lambda item: (
                item["y0"],
                item["x0"]
            )
        )

        page_layout = []

        # ====================================================
        # CREATE NORMAL TEXT ELEMENT IDS
        # ====================================================

        for ordered_item in ordered_items:

            if (
                ordered_item[
                    "kind"
                ]
                ==
                "table"
            ):

                table_data = (
                    ordered_item[
                        "data"
                    ]
                )

                page_layout.append({
                    "type":
                        "table",

                    "table_id":
                        table_data[
                            "table_id"
                        ]
                })

                continue

            block = (
                ordered_item[
                    "data"
                ]
            )

            x0, y0, x1, y1 = (
                block[
                    "bbox"
                ]
            )

            element_id = (
                f"pdf_{element_counter}"
            )

            element_counter += 1

            element = {
                "id":
                    element_id,

                "page":
                    page_index + 1,

                "type":
                    "text_block",

                "text":
                    block[
                        "text"
                    ],

                "bbox": {
                    "x0":
                        round(
                            x0,
                            2
                        ),

                    "y0":
                        round(
                            y0,
                            2
                        ),

                    "x1":
                        round(
                            x1,
                            2
                        ),

                    "y1":
                        round(
                            y1,
                            2
                        )
                }
            }

            elements.append(
                element
            )

            page_elements.append(
                element
            )

            page_layout.append({
                "type":
                    "text_block",

                "id":
                    element_id
            })

        # Table cell elements also belong to page_elements
        for table_data in table_structures:

            for row in table_data[
                "rows"
            ]:

                for cell in row:

                    cell_id = cell.get(
                        "id"
                    )

                    if not cell_id:
                        continue

                    for element in elements:

                        if (
                            element[
                                "id"
                            ]
                            ==
                            cell_id
                        ):
                            page_elements.append(
                                element
                            )
                            break

        pages.append({
            "page":
                page_index + 1,

            "width":
                round(
                    page_width,
                    2
                ),

            "height":
                round(
                    page_height,
                    2
                ),

            "element_count":
                len(
                    page_elements
                ),

            "elements":
                page_elements,

            "tables":
                table_structures,

            "layout":
                page_layout
        })

    # Global elements must be in deterministic ID order
    elements.sort(
        key=lambda item: int(
            item[
                "id"
            ].split(
                "_"
            )[1]
        )
    )

    return pages, elements


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

    section.top_margin = Pt(
        0
    )

    section.bottom_margin = Pt(
        0
    )

    section.left_margin = Pt(
        0
    )

    section.right_margin = Pt(
        0
    )

    section.header_distance = Pt(
        0
    )

    section.footer_distance = Pt(
        0
    )


def set_cell_margins(
    cell,
    top=40,
    start=60,
    bottom=40,
    end=60
):

    tc = cell._tc

    tcPr = tc.get_or_add_tcPr()

    tcMar = tcPr.first_child_found_in(
        "w:tcMar"
    )

    if tcMar is None:

        tcMar = OxmlElement(
            "w:tcMar"
        )

        tcPr.append(
            tcMar
        )

    for margin_name, margin_value in [
        ("top", top),
        ("start", start),
        ("bottom", bottom),
        ("end", end)
    ]:

        node = tcMar.find(
            qn(
                f"w:{margin_name}"
            )
        )

        if node is None:

            node = OxmlElement(
                f"w:{margin_name}"
            )

            tcMar.append(
                node
            )

        node.set(
            qn("w:w"),
            str(margin_value)
        )

        node.set(
            qn("w:type"),
            "dxa"
        )


def prevent_row_split(row):

    try:
        trPr = (
            row._tr.get_or_add_trPr()
        )

        cantSplit = trPr.find(
            qn("w:cantSplit")
        )

        if cantSplit is None:

            cantSplit = OxmlElement(
                "w:cantSplit"
            )

            trPr.append(
                cantSplit
            )

        cantSplit.set(
            qn("w:val"),
            "1"
        )

    except Exception:
        pass


def set_table_borders(table):

    try:
        tbl = table._tbl
        tblPr = tbl.tblPr

        borders = tblPr.first_child_found_in(
            "w:tblBorders"
        )

        if borders is None:

            borders = OxmlElement(
                "w:tblBorders"
            )

            tblPr.append(
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

            element = borders.find(
                qn(tag)
            )

            if element is None:

                element = OxmlElement(
                    tag
                )

                borders.append(
                    element
                )

            element.set(
                qn("w:val"),
                "single"
            )

            element.set(
                qn("w:sz"),
                "4"
            )

            element.set(
                qn("w:space"),
                "0"
            )

            element.set(
                qn("w:color"),
                "808080"
            )

    except Exception:
        pass


# ============================================================
# ADD NORMAL TEXT
# ============================================================

def add_text_block(
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

    right_indent = max(
        0,
        page_width - x1
    )

    paragraph.paragraph_format.right_indent = Pt(
        right_indent
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

    run.font.name = (
        DEFAULT_FONT_NAME
    )

    run.font.size = Pt(
        DEFAULT_FONT_SIZE
    )

    return y1


# ============================================================
# ADD WORD TABLE
# ============================================================

def add_word_table(
    document,
    table_data,
    translations,
    page_width,
    previous_bottom
):

    row_count = int(
        table_data[
            "row_count"
        ]
    )

    column_count = int(
        table_data[
            "column_count"
        ]
    )

    if (
        row_count <= 0
        or
        column_count <= 0
    ):
        return previous_bottom

    bbox = (
        table_data[
            "bbox"
        ]
    )

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

    # Vertical gap before table
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

    # Spacer paragraph before table
    spacer = (
        document.add_paragraph()
    )

    spacer.paragraph_format.space_before = Pt(
        vertical_gap
    )

    spacer.paragraph_format.space_after = Pt(
        0
    )

    spacer.paragraph_format.line_spacing = 1

    table = document.add_table(
        rows=row_count,
        cols=column_count
    )

    table.alignment = (
        WD_TABLE_ALIGNMENT.LEFT
    )

    table.autofit = False

    set_table_borders(
        table
    )

    # Approximate table indentation
    try:

        tblPr = (
            table._tbl.tblPr
        )

        tblInd = tblPr.find(
            qn("w:tblInd")
        )

        if tblInd is None:

            tblInd = OxmlElement(
                "w:tblInd"
            )

            tblPr.append(
                tblInd
            )

        # PDF points -> twips
        indent_twips = int(
            max(
                0,
                x0
            )
            *
            20
        )

        tblInd.set(
            qn("w:w"),
            str(
                indent_twips
            )
        )

        tblInd.set(
            qn("w:type"),
            "dxa"
        )

    except Exception:
        pass

    total_pdf_width = max(
        1.0,
        x1 - x0
    )

    available_page_width = max(
        1.0,
        page_width - x0
    )

    word_table_width = min(
        total_pdf_width,
        available_page_width
    )

    column_widths = (
        table_data.get(
            "column_widths"
        )
        or []
    )

    if len(
        column_widths
    ) != column_count:

        column_widths = [
            total_pdf_width /
            column_count
            for _ in range(
                column_count
            )
        ]

    width_sum = sum(
        column_widths
    )

    if width_sum <= 0:

        width_sum = (
            total_pdf_width
        )

    rows_data = (
        table_data[
            "rows"
        ]
    )

    for row_index in range(
        row_count
    ):

        word_row = (
            table.rows[
                row_index
            ]
        )

        prevent_row_split(
            word_row
        )

        for col_index in range(
            column_count
        ):

            cell = (
                word_row.cells[
                    col_index
                ]
            )

            cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

            set_cell_margins(
                cell
            )

            proportional_width = (
                column_widths[
                    col_index
                ]
                /
                width_sum
            )

            cell_width = (
                word_table_width
                *
                proportional_width
            )

            cell.width = Pt(
                max(
                    10,
                    cell_width
                )
            )

            cell_info = None

            if row_index < len(
                rows_data
            ):

                row_data = (
                    rows_data[
                        row_index
                    ]
                )

                if col_index < len(
                    row_data
                ):

                    cell_info = (
                        row_data[
                            col_index
                        ]
                    )

            if cell_info is None:
                continue

            element_id = (
                cell_info.get(
                    "id"
                )
            )

            original_text = (
                cell_info.get(
                    "text"
                )
                or ""
            )

            if element_id:

                final_text = (
                    translations.get(
                        element_id,
                        original_text
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

            paragraph.paragraph_format.line_spacing = 1

            paragraph.alignment = (
                WD_ALIGN_PARAGRAPH.LEFT
            )

            if paragraph.runs:

                paragraph.runs[
                    0
                ].text = final_text

                run = paragraph.runs[
                    0
                ]

            else:

                run = paragraph.add_run(
                    final_text
                )

            run.font.name = (
                DEFAULT_FONT_NAME
            )

            run.font.size = Pt(
                TABLE_FONT_SIZE
            )

    return y1


# ============================================================
# CREATE WORD FROM PDF
# ============================================================

def create_word_from_pdf(
    pdf_document,
    translations
):

    pages, elements = (
        extract_pdf_structure(
            pdf_document
        )
    )

    document = Document()

    # Remove visual impact of default paragraph
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

        paragraph.paragraph_format.line_spacing = 1

    element_lookup = {
        element["id"]:
            element
        for element in elements
    }

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

        for layout_item in page_data.get(
            "layout",
            []
        ):

            item_type = (
                layout_item.get(
                    "type"
                )
            )

            # ================================================
            # NORMAL TEXT
            # ================================================

            if item_type == "text_block":

                element_id = (
                    layout_item.get(
                        "id"
                    )
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
                        element[
                            "text"
                        ]
                    )
                )

                previous_bottom = (
                    add_text_block(
                        document,
                        element,
                        translated_text,
                        page_width,
                        previous_bottom
                    )
                )

            # ================================================
            # TABLE
            # ================================================

            elif item_type == "table":

                table_id = (
                    layout_item.get(
                        "table_id"
                    )
                )

                table_data = (
                    table_lookup.get(
                        table_id
                    )
                )

                if not table_data:
                    continue

                previous_bottom = (
                    add_word_table(
                        document,
                        table_data,
                        translations,
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

        pages, elements = (
            extract_pdf_structure(
                document
            )
        )

        chunks = [
            elements[
                i:i + CHUNK_SIZE
            ]
            for i in range(
                0,
                len(elements),
                CHUNK_SIZE
            )
        ]

        table_count = sum(
            len(
                page.get(
                    "tables",
                    []
                )
            )
            for page in pages
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

            "table_count":
                table_count,

            "chunk_size":
                CHUNK_SIZE,

            "chunk_count":
                len(
                    chunks
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

    # --------------------------------------------------------
    # PDF CHECK
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # TRANSLATIONS CHECK
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # PARSE TRANSLATIONS
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # OPEN ORIGINAL PDF
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # CREATE WORD
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # OUTPUT NAME
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # RETURN DOCX
    # --------------------------------------------------------

    return send_file(
        word_file,
        as_attachment=True,
        download_name=output_filename,
        mimetype=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


# ============================================================
# LOCAL
# ============================================================

if __name__ == "__main__":
    app.run()