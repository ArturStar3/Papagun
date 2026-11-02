"""Apply the visible notice to every slide of a demo report."""

import base64
import hashlib
import secrets
from io import BytesIO

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


NOTICE = "Данная демо-версия проекта является ознакомительной и не предназначена для использования"
WATERMARK_COUNT = 5
WATERMARK_ANGLE = 350


def _protect_from_manual_editing(presentation):
    """Ask PowerPoint to open the deck write-protected, without saving a password."""
    salt = secrets.token_bytes(16)
    password = secrets.token_hex(24).encode("utf-16le")
    digest = hashlib.sha512(salt + password).digest()
    spins = 100_000
    for count in range(spins):
        digest = hashlib.sha512(digest + count.to_bytes(4, "little")).digest()
    verifier = OxmlElement("p:modifyVerifier")
    for key, value in {
        "cryptProviderType": "rsaAES",
        "cryptAlgorithmClass": "hash",
        "cryptAlgorithmType": "typeAny",
        "cryptAlgorithmSid": "14",
        "spinCount": str(spins),
        "saltData": base64.b64encode(salt).decode("ascii"),
        "hashData": base64.b64encode(digest).decode("ascii"),
    }.items():
        verifier.set(key, value)
    element = presentation._element
    extension = next((child for child in element if child.tag.endswith("}extLst")), None)
    if extension is None:
        element.append(verifier)
    else:
        element.insert(element.index(extension), verifier)


def apply_watermark(data: bytes) -> bytes:
    presentation = Presentation(BytesIO(data))
    for slide in presentation.slides:
        width, height = Inches(10.2), Inches(0.46)
        center_left = (presentation.slide_width - width) / 2
        for index in range(WATERMARK_COUNT):
            shape = slide.shapes.add_textbox(
                int(center_left + Inches((index - 2) * 0.35)),
                Inches(1.0 + index * 1.25), width, height,
            )
            shape.name = f"Papagun Demo Watermark {index + 1}"
            shape.rotation = WATERMARK_ANGLE
            frame = shape.text_frame
            frame.clear()
            frame.word_wrap = False
            frame.margin_left = frame.margin_right = 0
            frame.margin_top = frame.margin_bottom = 0
            frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            paragraph = frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.CENTER
            run = paragraph.add_run()
            run.text = NOTICE
            run.font.name = "Arial"
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(145, 145, 145)
    _protect_from_manual_editing(presentation)
    output = BytesIO()
    presentation.save(output)
    return output.getvalue()
