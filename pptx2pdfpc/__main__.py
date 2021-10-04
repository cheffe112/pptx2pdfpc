"""
Entry point module for the execution of `python -m pptx2pdfpc`
See:
https://docs.python.org/3/using/cmdline.html#cmdoption-m
https://www.python.org/dev/peps/pep-0338/

"""
from pptx2pdfpc import cli

if __name__ == "__main__":
    cli.main(prog_name="pptx2pdfpc")  # https://github.com/pallets/click/issues/1399
