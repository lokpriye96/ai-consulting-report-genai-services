"""Adds a light author watermark to every page of the deck PDF and locks it against editing.

Run: python src/protect_pdf.py outputs/TCS_AI_Deflation_Strategy.pdf
The PDF opens without a password; editing, annotating and page assembly are disabled.
No PDF protection is unbreakable: the dated GitHub commit history is the real proof of authorship.
"""
import io
import secrets
import sys

from pypdf import PdfReader, PdfWriter
from pypdf.constants import UserAccessPermissions as P
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas

MARK = "© Utkarsh Saxena  ·  github.com/lokpriye96"


def stamp(w, h):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    c.setFillColor(Color(0.26, 0.13, 0.30, alpha=0.07))
    c.setFont("Helvetica-Bold", 34)
    c.translate(w / 2, h / 2)
    c.rotate(28)
    c.drawCentredString(0, 0, MARK)
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def main(path):
    reader = PdfReader(path)
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(stamp(float(page.mediabox.width), float(page.mediabox.height)))
        writer.add_page(page)
    writer.add_metadata({"/Author": "Utkarsh Saxena", "/Title": "Getting ahead of AI deflation: a growth agenda for TCS"})
    writer.encrypt(user_password="", owner_password=secrets.token_hex(16),
                   permissions_flag=P.PRINT | P.PRINT_TO_REPRESENTATION | P.EXTRACT_TEXT_AND_GRAPHICS | P.EXTRACT,
                   algorithm="AES-256")
    with open(path, "wb") as f:
        writer.write(f)
    print("protected", path)


if __name__ == "__main__":
    main(sys.argv[1])
