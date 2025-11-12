from flask import Flask, request, send_file, jsonify
from pdf2docx import Converter
import os
import tempfile
from flask_cors import CORS
from datetime import datetime
import logging
from docx2pdf import convert
import pythoncom

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# AGPL Compliance - Display license info
@app.route('/')
def home():
    return jsonify({
        "service": "PDF to Word & Word to PDF Converter",
        "version": "2.0",
        "license": "GNU AGPL v3.0",
        "source_code": "https://github.com/your-username/pdf2docx",
        "legal_notice": "This service uses pdf2docx licensed under GNU AGPL v3.0"
    })

@app.route('/health')
def health():
    return jsonify({"status": "healthy", "timestamp": datetime.utcnow().isoformat()})

@app.route('/convert', methods=['POST'])
def convert_pdf_to_word():
    """
    Convert PDF to Word document
    """
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            logger.error("No file provided in request")
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        # Validate file
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400
        
        if not file.filename.lower().endswith('.pdf'):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'File must be a PDF'}), 400
        
        # Check file size (limit to 15MB for free tier)
        file.seek(0, os.SEEK_END)
        file_length = file.tell()
        file.seek(0)
        
        if file_length > 15 * 1024 * 1024:  # 15MB limit
            logger.error(f"File too large: {file_length} bytes")
            return jsonify({'error': 'File size must be less than 15MB'}), 400
        
        # Get conversion parameters
        page_range = request.form.get('page_range', '')
        image_quality = request.form.get('image_quality', 'medium')
        
        logger.info(f"Converting PDF: {file.filename}, size: {file_length} bytes")
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as pdf_temp:
            pdf_path = pdf_temp.name
            file.save(pdf_path)
        
        docx_path = pdf_path.replace('.pdf', '.docx')
        
        try:
            # Convert PDF to Word using pdf2docx
            cv = Converter(pdf_path)
            
            # Set conversion options
            convert_kwargs = {}
            if page_range:
                convert_kwargs['pages'] = parse_page_range(page_range)
            
            if image_quality == 'low':
                convert_kwargs['rotate_page'] = False
            
            # Perform conversion
            cv.convert(docx_path, **convert_kwargs)
            cv.close()
            
            logger.info(f"Conversion successful: {docx_path}")
            
            # Return the converted file
            download_name = file.filename.replace('.pdf', '.docx').replace('.PDF', '.docx')
            
            return send_file(
                docx_path,
                as_attachment=True,
                download_name=download_name,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
            
        except Exception as conversion_error:
            logger.error(f"Conversion failed: {str(conversion_error)}")
            return jsonify({'error': f'Conversion failed: {str(conversion_error)}'}), 500
            
        finally:
            # Clean up temporary files
            cleanup_file(pdf_path)
            cleanup_file(docx_path)
                
    except Exception as e:
        logger.error(f"Server error: {str(e)}")
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/word-to-pdf', methods=['POST'])
def convert_word_to_pdf():
    """
    Convert Word document to PDF
    """
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            logger.error("No file provided in request")
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        # Validate file
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400
        
        # Check for Word document formats
        valid_extensions = ['.docx', '.doc']
        file_ext = os.path.splitext(file.filename.lower())[1]
        if file_ext not in valid_extensions:
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'File must be a Word document (.docx or .doc)'}), 400
        
        # Check file size (limit to 15MB for free tier)
        file.seek(0, os.SEEK_END)
        file_length = file.tell()
        file.seek(0)
        
        if file_length > 15 * 1024 * 1024:  # 15MB limit
            logger.error(f"File too large: {file_length} bytes")
            return jsonify({'error': 'File size must be less than 15MB'}), 400
        
        logger.info(f"Converting Word to PDF: {file.filename}, size: {file_length} bytes")
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(suffix=file_ext, delete=False) as word_temp:
            word_path = word_temp.name
            file.save(word_path)
        
        pdf_path = word_path.replace(file_ext, '.pdf')
        
        try:
            # Initialize COM for Windows (required for docx2pdf)
            pythoncom.CoInitialize()
            
            # Convert Word to PDF
            convert(word_path, pdf_path)
            
            logger.info(f"Conversion successful: {pdf_path}")
            
            # Return the converted file
            download_name = file.filename.replace(file_ext, '.pdf')
            
            return send_file(
                pdf_path,
                as_attachment=True,
                download_name=download_name,
                mimetype='application/pdf'
            )
            
        except Exception as conversion_error:
            logger.error(f"Conversion failed: {str(conversion_error)}")
            return jsonify({'error': f'Conversion failed: {str(conversion_error)}'}), 500
            
        finally:
            # Clean up temporary files
            cleanup_file(word_path)
            cleanup_file(pdf_path)
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
                
    except Exception as e:
        logger.error(f"Server error: {str(e)}")
        return jsonify({'error': f'Server error: {str(e)}'}), 500

def parse_page_range(page_range_str):
    """
    Parse page range string like '1-5,7,9-12'
    """
    if not page_range_str:
        return None
    
    pages = []
    ranges = page_range_str.split(',')
    
    for r in ranges:
        r = r.strip()
        if not r:
            continue
            
        if '-' in r:
            try:
                start, end = map(int, r.split('-'))
                pages.extend(range(start, end + 1))
            except ValueError:
                raise ValueError(f"Invalid range: {r}")
        else:
            try:
                pages.append(int(r))
            except ValueError:
                raise ValueError(f"Invalid page number: {r}")
    
    return pages

def cleanup_file(file_path):
    """Safely remove temporary files"""
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            logger.info(f"Cleaned up: {file_path}")
    except Exception as e:
        logger.warning(f"Could not remove {file_path}: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
