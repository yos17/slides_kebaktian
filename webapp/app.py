#!/usr/bin/env python3
"""
Flask Web App for PowerPoint Song Generator
Allows users to upload song files and generate PowerPoint presentations through a web interface.
"""

import os
import uuid
import time
from datetime import datetime, timedelta
from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
import threading
import json

# Import our generator
from generator import generate_presentation

app = Flask(__name__)
app.secret_key = 'powerpoint-song-generator-secret-key-2024'

# Configuration
UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'generated'
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB
ALLOWED_TEXT_EXTENSIONS = {'txt'}
ALLOWED_PPTX_EXTENSIONS = {'pptx'}
FILE_CLEANUP_HOURS = 2  # Clean up files after 2 hours

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATED_FOLDER'] = GENERATED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

# Global job tracking
processing_jobs = {}

def allowed_file(filename, extensions):
    """Check if file has allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in extensions

def cleanup_old_files():
    """Clean up old uploaded and generated files."""
    cutoff_time = datetime.now() - timedelta(hours=FILE_CLEANUP_HOURS)
    
    for folder in [UPLOAD_FOLDER, GENERATED_FOLDER]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path):
                file_time = datetime.fromtimestamp(os.path.getctime(file_path))
                if file_time < cutoff_time:
                    try:
                        os.remove(file_path)
                        print(f"Cleaned up old file: {file_path}")
                    except OSError:
                        pass

def process_files_async(job_id, song_file_path, template_file_path, generate_toc, output_filename):
    """Process files in background thread."""
    try:
        processing_jobs[job_id]['status'] = 'processing'
        processing_jobs[job_id]['message'] = 'Parsing songs...'
        
        # Generate the presentation
        output_path = os.path.join(GENERATED_FOLDER, output_filename)
        
        success, message, slide_count = generate_presentation(
            song_file_path, 
            output_path, 
            template_file_path, 
            generate_toc
        )
        
        if success:
            processing_jobs[job_id]['status'] = 'completed'
            processing_jobs[job_id]['message'] = f'Successfully generated {slide_count} slides!'
            processing_jobs[job_id]['output_file'] = output_filename
        else:
            processing_jobs[job_id]['status'] = 'error'
            processing_jobs[job_id]['message'] = f'Error: {message}'
            
    except Exception as e:
        processing_jobs[job_id]['status'] = 'error'
        processing_jobs[job_id]['message'] = f'Unexpected error: {str(e)}'

@app.route('/')
def index():
    """Main upload page."""
    cleanup_old_files()  # Clean up old files on each visit
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file upload and start processing."""
    try:
        # Check if files were uploaded
        if 'song_file' not in request.files:
            flash('No song file selected', 'error')
            return redirect(url_for('index'))
        
        song_file = request.files['song_file']
        template_file = request.files.get('template_file')
        generate_toc = 'generate_toc' in request.form
        output_filename = request.form.get('output_filename', 'songs_presentation.pptx')
        
        # Ensure output filename has .pptx extension
        if not output_filename.endswith('.pptx'):
            output_filename += '.pptx'
        
        # Validate song file
        if song_file.filename == '':
            flash('No song file selected', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(song_file.filename, ALLOWED_TEXT_EXTENSIONS):
            flash('Song file must be a .txt file', 'error')
            return redirect(url_for('index'))
        
        # Validate template file if provided
        template_file_path = None
        if template_file and template_file.filename != '':
            if not allowed_file(template_file.filename, ALLOWED_PPTX_EXTENSIONS):
                flash('Template file must be a .pptx file', 'error')
                return redirect(url_for('index'))
            
            # Save template file
            template_filename = secure_filename(f"{uuid.uuid4().hex}_{template_file.filename}")
            template_file_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
            template_file.save(template_file_path)
        
        # Save song file
        song_filename = secure_filename(f"{uuid.uuid4().hex}_{song_file.filename}")
        song_file_path = os.path.join(app.config['UPLOAD_FOLDER'], song_filename)
        song_file.save(song_file_path)
        
        # Create job ID and start processing
        job_id = str(uuid.uuid4())
        processing_jobs[job_id] = {
            'status': 'starting',
            'message': 'Starting processing...',
            'created_at': datetime.now()
        }
        
        # Start background processing
        thread = threading.Thread(
            target=process_files_async,
            args=(job_id, song_file_path, template_file_path, generate_toc, output_filename)
        )
        thread.daemon = True
        thread.start()
        
        return render_template('processing.html', job_id=job_id)
        
    except Exception as e:
        flash(f'Error processing files: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/status/<job_id>')
def get_status(job_id):
    """Get processing status for a job."""
    if job_id not in processing_jobs:
        return jsonify({'status': 'not_found', 'message': 'Job not found'}), 404
    
    job = processing_jobs[job_id]
    response = {
        'status': job['status'],
        'message': job['message']
    }
    
    if job['status'] == 'completed' and 'output_file' in job:
        response['download_url'] = url_for('download_file', filename=job['output_file'])
    
    return jsonify(response)

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated PowerPoint file."""
    file_path = os.path.join(app.config['GENERATED_FOLDER'], filename)
    
    if not os.path.exists(file_path):
        flash('File not found or has expired', 'error')
        return redirect(url_for('index'))
    
    return send_file(file_path, as_attachment=True, download_name=filename)

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error."""
    flash('File is too large. Maximum size is 16MB.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("🎵 PowerPoint Song Generator Web App")
    print("=" * 50)
    print("Starting Flask development server...")
    print("Visit http://localhost:8080 in your browser")
    print("=" * 50)
    
    app.run(debug=True, host='0.0.0.0', port=8080)