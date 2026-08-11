"""Create the public reference DOCX used by the repository example."""

from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "templates" / "public-demo.docx"


def set_font(style, name: str, size: float, bold: bool = False) -> None:
    style.font.name = name
    style._element.rPr.rFonts.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", name
    )
    style.font.size = Pt(size)
    style.font.bold = bold


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

    set_font(document.styles["Normal"], "Aptos", 10.5)
    set_font(document.styles["Title"], "Aptos Display", 20, bold=True)
    set_font(document.styles["Heading 1"], "Aptos", 16, bold=True)
    set_font(document.styles["Heading 2"], "Aptos", 13, bold=True)

    if "Caption" not in [style.name for style in document.styles]:
        caption = document.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
    else:
        caption = document.styles["Caption"]
    set_font(caption, "Aptos", 10.5)
    caption.paragraph_format.space_before = Pt(6)
    caption.paragraph_format.space_after = Pt(6)

    document.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
