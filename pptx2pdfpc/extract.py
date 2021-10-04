import pathlib
import pprint
from typing import Tuple, List, Any

from pptx import Presentation

# logic for extracting speaker notes and text boxes by user fusion on Stack Overflow, see:
# https://stackoverflow.com/questions/63659972/extract-presenter-notes-from-pptx-file-powerpoint
# Many thanks <3


def speaker_notes(input_pptx: pathlib.Path) -> List[Tuple[int, str]]:
    """Extract all speaker notes from a pptx."""
    prs = Presentation(str(input_pptx))
    extracted = []
    for page, slide in enumerate(prs.slides, start=1):
        text = slide.notes_slide.notes_text_frame.text
        extracted.append((page, text))
    return extracted


def text_boxes(input_pptx: pathlib.Path) -> List[Tuple[int, List[Any]]]:
    """Extract all text boxes from all slides."""
    prs = Presentation(str(input_pptx))
    extracted = []
    for page, slide in enumerate(prs.slides, start=1):
        temp = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                temp.append(shape.text)
        extracted.append((page, temp))
    return extracted


def generate_pdfpc(extracted_notes: List[Tuple[int, str]], output_path: pathlib.Path):
    """Generate a config file for pdfpc and writes it to output_path.
    The file ending is .pdfpc and can be used together with a PDF
    version of the pptx presentation.
    From the pdfpc man page: "When pdfpc is invoked with a PDF file, it automatically
    checks for and loads the associated .pdfpc file, if it exists."

    Requirement: Both files need to have the same name, one ending in .pdf, the other in .pdfpc.
    For pdfpc, see its man page or https://man.archlinux.org/man/community/pdfpc/pdfpc.1.en
    """
    DELIMITER = "###"
    with output_path.open("a") as fo:
        fo.write("[notes]\n")
        for slide in extracted_notes:
            page_number = slide[0]
            note_text = slide[1]
            if len(note_text) > 0:
                fo.write(f"{DELIMITER} {page_number}\n")
                fo.write(f"{note_text}\n")
                fo.write("\n")


def generate_output_path(input_path: pathlib.Path) -> pathlib.Path:
    """Generate the path to the notes file ending in .pdfpc. It must have
    the same name as the presentation pdf and be in the same directory
    to be visible during the presentation."""
    filename = input_path.stem
    path = input_path.parent
    return path / pathlib.Path(filename + ".pdfpc")
