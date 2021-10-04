import pathlib
import sys

import click


from pptx2pdfpc.extract import speaker_notes, generate_pdfpc, generate_output_path


@click.command()
@click.argument("input_path", type=click.STRING)
def main(input_path):
    try:
        pptx = pathlib.Path(input_path)
        notes = speaker_notes(pptx)
    except OSError as e:
        print(f"An error happened while reading in the pptx:\n{e}")
        sys.exit(1)

    output_path = generate_output_path(pptx)

    try:
        generate_pdfpc(notes, output_path)
    except OSError as e:
        print(f"An error happened while generating the pdfpc notes in the pptx:\n{e}")
        sys.exit(1)
