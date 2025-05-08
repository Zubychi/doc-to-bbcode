"""
    This script reads a Word document and prints the text in BBCode.
"""

from docx import Document
from sys import argv
from os.path import exists

if len(argv) > 1:
    filename = argv[1]
else:
    filename = "test.docx"

if not exists(filename):
    print(f"File {filename} does not exist.")
    exit(1)
doc = Document(filename)

stylings = {
    "bold": False,
    "italic": False,
    "underline": False,
    "color": None,
    "size": None,
    "font": None,
    "background": None,
    "strike": False,
    "align": None,
}

alignments = {
    0: "left",
    1: "center",
    2: "right",
    3: "justify",
}

text = ""
par_out = ""
output = ""

for par in doc.paragraphs:
    for run in par.runs:
        text = run.text
        if run.bold:
            text = f"{"[b]"}{text}{"[/b]"}"
        if run.italic:
            text = f"{"[i]"}{text}{"[/i]"}"
        if run.underline:
            text = f"{"[u]"}{text}{"[/u]"}"
        if run.font.color.rgb:
            text = f"{"[color=#"}{run.font.color.rgb}]{text}{"[/color]"}"
        if run.font.size:
            text = f"{"[size="}{run.font.size.pt}]{text}{"[/size]"}"
        if run.font.name:
            text = f"{"[font="}{run.font.name}]{text}{"[/font]"}"
        if run.font.highlight_color:
            text = f"{"[highlight="}{run.font.highlight_color}]{text}{"[/highlight]"}"
        if run.font.strike:
            text = f"{"[strike]"}{text}{"[/strike]"}"
        par_out += text

    if par.alignment:
        if par.alignment in alignments:
            stylings["align"] = alignments[par.alignment]
        else:
            stylings["align"] = "left"
        par_out = f"[align={stylings["align"]}]{par_out}[/align]"
    output += par_out + "\n"
    par_out = ""

print(output)
print("BBCode generated successfully.")