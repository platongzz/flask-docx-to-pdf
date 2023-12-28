from flask import Flask, request, jsonify, send_file
from docx2pdf import convert
from datetime import datetime
import pythoncom

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_docx_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'})
    
    if file and file.filename.endswith('.docx'):
        try:
            pythoncom.CoInitialize()
            
            # Save the uploaded DOCX file
            docx_filename = 'uploaded.docx'
            file.save(docx_filename)
            
            # Convert DOCX to PDF
            pdf_filename = datetime.now().strftime("%Y%m%d%H%M%S") + '_converted.pdf'
            convert(docx_filename, "D:\\TechMeMars\\Docx2Pdf\\" + pdf_filename)
            
            return jsonify({'success': True, 'pdf_filename': "http://127.0.0.1:5000/output/" + pdf_filename})
        except Exception as e:
            return jsonify({'error': f'Conversion failed: {str(e)}'})
    
    return jsonify({'error': 'Invalid file format, please upload a DOCX file'})

@app.route('/output/<path:filename>', methods=['GET'])
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
