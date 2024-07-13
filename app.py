from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import os
from werkzeug.utils import secure_filename
from vba_analysis.extract import extract_vba_code, analyze_vba_code, generate_detailed_documentation, generate_pdf, generate_flowchart

app = Flask(__name__)
app.secret_key = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['PDF_FOLDER'] = 'static/pdfs/'

# Ensure the upload and pdf folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PDF_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No selected file')
        return redirect(url_for('index'))
    
    if file and file.filename.endswith('.xlsm'):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        vba_code = extract_vba_code(file_path)
        if vba_code:
            analysis_results = analyze_vba_code(vba_code)
            documentation = generate_detailed_documentation(vba_code, filename, analysis_results)
            pdf_filename = filename.replace('.xlsm', '_documentation.pdf')
            generate_pdf(documentation, pdf_filename)
            flowchart_path = generate_flowchart(vba_code)
            return render_template('results.html', vba_code=vba_code, analysis_results=analysis_results, documentation=documentation.split('\n'), pdf_filename=pdf_filename, flowchart_path=flowchart_path)
        else:
            flash('Failed to extract VBA code from the file.')
            return redirect(url_for('index'))
    else:
        flash('Invalid file format. Please upload an XLSM file.')
        return redirect(url_for('index'))

@app.route('/download_pdf/<filename>')
def download_pdf(filename):
    return send_from_directory(app.config['PDF_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
