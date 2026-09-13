from flask import Flask, request, jsonify
import fitz
from io import BytesIO


app = Flask(__name__)


@app.route("/", methods=["GET"])
def home():
    return {
        "status": "ok",
        "service": "PDF Parser"
    }


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
            filetype="pdf"
        )

    except Exception as e:
        return jsonify({
            "error": "Could not read PDF file",
            "details": str(e)
        }), 400

    pages = []
    elements = []

    element_counter = 1

    for page_index in range(
        len(document)
    ):
        page = document[
            page_index
        ]

        blocks = page.get_text(
            "blocks"
        )

        page_elements = []

        for block in blocks:

            if len(block) < 5:
                continue

            x0 = block[0]
            y0 = block[1]
            x1 = block[2]
            y1 = block[3]
            text = block[4].strip()

            if not text:
                continue

            element = {
                "id": f"pdf_{element_counter}",
                "page": page_index + 1,
                "type": "text_block",
                "text": text,
                "bbox": {
                    "x0": round(x0, 2),
                    "y0": round(y0, 2),
                    "x1": round(x1, 2),
                    "y1": round(y1, 2)
                }
            }

            elements.append(
                element
            )

            page_elements.append(
                element
            )

            element_counter += 1

        pages.append({
            "page": page_index + 1,
            "width": round(
                page.rect.width,
                2
            ),
            "height": round(
                page.rect.height,
                2
            ),
            "element_count": len(
                page_elements
            ),
            "elements": page_elements
        })

    document.close()

    return jsonify({
        "filename": uploaded_file.filename,
        "page_count": len(pages),
        "element_count": len(elements),
        "pages": pages,
        "elements": elements
    })


if __name__ == "__main__":
    app.run()