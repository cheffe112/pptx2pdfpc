# pptx2pdfpc
Extracts speaker notes from a pptx and transforms it into a .pdfpc file that contains the notes.

PowerPoint is nice to create slides, but there's still a Linux version lacking and the online PowerPoint version does not offer the full Windows capabilities of PowerPoint. It is also good to not depend on an online version of PowerPoint, since you never know when services are unavailable.

A solution for Linux: Create your PowerPoint presentation in your VM of choice, export PowerPoint file as pdf and present it with [pdfpc](https://pdfpc.github.io/). Pdfpc can be used in conjunction with a `.pdfpc` configuration file that contains speaker notes, the file must simply carry the same name as the pdf presentation. Unfortunately, speaker notes are not extracted or transferred when the pdf is generated. It is very cumbersome and inefficient to write the .pdfpc file by hand, copy-pasting the speaker notes into the file.

This tool automates the process: Give it your pptx-file and it extracts all speaker notes, generating the required pdfpc-file. You still need to export the pptx file as pdf by hand.

# Thanks
The tool uses the library [python-pptx](https://python-pptx.readthedocs.io/). I did not immediately find the library's functions to extract the speaker notes, and a search guided me to user [fusion's answer on StackOverflow](https://stackoverflow.com/questions/63659972/extract-presenter-notes-from-pptx-file-powerpoint). My tool extracts the speaker notes based on their SO answer, and I am indebted to and grateful for their solution. Thanks!
