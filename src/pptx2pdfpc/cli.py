import pathlib
import sys

import click

from pptx2pdfpc.errors import Error
from pptx2pdfpc.extract import speaker_notes, generate_pdfpc, generate_output_path



# Options in pdfpc are not fully documented, but there's open issues and PRs concerning this:
# https://github.com/pdfpc/pdfpc/issues/605
# https://json-schema.app/view/%23?url=https%3A%2F%2Fraw.githubusercontent.com%2Fpdfpc%2Fpdfpc%2Fedb06b2c07e10a9210d4f1ae5861d2e847ca5065%2Fschema%2Fpdfpc.json




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

@click.option(
    "--duration",
    "-d",
    "duration",
    nargs=1,
    type=click.INT,
    help="Duration in minutes of the presentation used for timer display.",
    metavar="minutes",
)
@click.option(
    "--start-time",
    "-t",
    "start_time",
    nargs=1,
    type=click.STRING,
    help="Start time of the presentation to be used as a countdown. (Format: HH:MM (24h))",
    metavar="hh:mm",
)
@click.option(
    "--end-time",
    "-e",
    "end_time",
    nargs=1,
    type=click.STRING,
    help="End time of the presentation. (Format: HH:MM (24h))",
    metavar="hh:mm",
)
@click.option(
    "--last-minutes",
    "-l",
    "last_minutes",
    nargs=1,
    type=click.INT,
    help="Time in minutes, from which on the timer changes its color. (Default: 5 minutes)",
    metavar="mins",
)
@click.option(
    "--end-slide",
    "-s",
    "end_slide",
    nargs=1,
    type=click.INT,
    help="Set the last slide of the presentation.",
    metavar="mins",
)
@click.argument("input_path", type=click.STRING)
def main(input_path, font_size, duration, start_time, end_time, last_minutes, end_slide):
    options = [("font_size", font_size), ("duration", duration), ("start_time", start_time), ("end_time", end_time), ("last_minutes", last_minutes), ("end_slide", end_slide)]
    try:
        pptx = pathlib.Path(input_path)
        notes = speaker_notes(pptx)
    except Error as e:
        print(e)
        sys.exit(1)

    output_path = generate_output_path(pptx)

    generate_pdfpc(notes, options, output_path)
    print(f"The notes have been extracted and can be found here:\n{output_path.name}")