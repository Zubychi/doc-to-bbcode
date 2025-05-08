from flask import Flask, render_template_string, request
from docx import Document

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.docx'):
            return f"File uploaded successfully!"

    return render_template_string(
        """
        <title>Upload a DOCX file</title>
        <h1>Upload a DOCX file</h1>
        <form method=post enctype=multipart/form-data>
          <input type=file name=file>
          <input type=submit value=Upload>
        </form>
        """
    )


if __name__ == '__main__':
    app.run()