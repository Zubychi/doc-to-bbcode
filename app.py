from flask import Flask, render_template, request
from docx import Document

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload():
    crunched = ""
    if request.method == 'POST':
        file = request.files['file']
        if not file or not file.filename.endswith('.docx'):
            return render_template('upload.html', error="Invalid file type. Please upload a .docx file.")
        crunched = docx_to_bbcode(file)
    return render_template('upload.html', text=crunched)


def docx_to_bbcode(file):
    doc = Document(file)
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
                text = f"[b]{text}[/b]"
            if run.italic:
                text = f"[i]{text}[/i]"
            if run.underline:
                text = f"[u]{text}[/u]"
            if run.font.color and run.font.color.rgb:
                text = f"[color=#{run.font.color.rgb}]{text}[/color]"
            if run.font.size:
                text = f"[size={run.font.size.pt}]{text}[/size]"
            if run.font.name:
                text = f"[font={run.font.name}]{text}[/font]"
            if run.font.highlight_color:
                text = f"[highlight={run.font.highlight_color}]{text}[/highlight]"
            if run.font.strike:
                text = f"[strike]{text}[/strike]"
            par_out += text

        if par.alignment in alignments:
            stylings["align"] = alignments[par.alignment]
        else:
            stylings["align"] = "left"
        par_out = f"[align={stylings['align']}]{par_out}[/align]"
        output += par_out + "\n"
        par_out = ""

    return output


if __name__ == '__main__':
    app.run()