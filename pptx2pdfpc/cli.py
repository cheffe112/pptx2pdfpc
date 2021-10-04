import pathlib
import sys

import click


from pptx2pdfpc.extract import speaker_notes, generate_pdfpc, generate_output_path


@click.command()
@click.option(
    "--font-size",
    "-f",
    "font_size",
    nargs=1,
    type=click.INT,
    help="Set optional font size for the speaker notes",
    metavar="pt",
)

@click.argument("input_path", type=click.STRING)
def main(input_path, font_size):
    try:
        pptx = pathlib.Path(input_path)
        notes = speaker_notes(pptx)
    except OSError as e:
        print(f"An error happened while reading in the pptx:\n{e}")
        sys.exit(1)

    output_path = generate_output_path(pptx)

    try:
        generate_pdfpc(notes, output_path, font_size)
    except OSError as e:
        print(f"An error happened while generating the pdfpc notes in the pptx:\n{e}")
        sys.exit(1)
