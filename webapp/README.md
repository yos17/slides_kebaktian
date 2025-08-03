# PowerPoint Song Generator - Web App

A Flask web application that allows users to upload song files and generate professional PowerPoint presentations through a web interface.

## Features

✨ **User-Friendly Interface**
- Drag & drop file upload
- Bootstrap responsive design
- Real-time processing status
- Professional UI with animations

🎵 **PowerPoint Generation**
- Automatic slide numbering (1/4, 2/4, etc.)  
- Professional Calibri fonts (32pt titles, 28pt content)
- Clickable Table of Contents with 2-column layout
- PowerPoint template support
- Multi-language support (English, Indonesian, German)

🔒 **Secure & Reliable**
- File validation and size limits
- Automatic cleanup of temporary files
- Error handling and user feedback
- Background processing with status updates

## Quick Start

### Prerequisites
```bash
pip3 install flask python-pptx
```

### Running Locally
```bash
python3 app.py
```

Then visit: `http://localhost:5000`

### Usage
1. **Upload Song File**: Drag & drop or browse for your `.txt` song file
2. **Optional Template**: Upload a PowerPoint template (`.pptx`) for branded slides
3. **Configure Options**: Choose whether to generate Table of Contents
4. **Generate**: Click "Generate PowerPoint" and wait for processing
5. **Download**: Download your professional presentation

## Song File Format

Your song text file should follow this format:

```
# Amazing Grace

Amazing Grace, how sweet the sound
That saved a wretch like me

I once was lost, but now am found
Was blind, but now I see

# How Great Thou Art

O Lord my God, when I in awesome wonder
Consider all the worlds Thy hands have made
```

**Rules:**
- Start each song with `#` followed by the song title
- Separate verses/choruses with empty lines
- Each verse/chorus becomes a separate slide
- No special formatting needed - just plain text

## File Structure

```
webapp/
├── app.py                 # Main Flask application
├── generator.py           # PowerPoint generation logic
├── requirements.txt       # Python dependencies
├── templates/
│   ├── index.html         # Main upload interface
│   └── processing.html    # Processing status page
├── static/
│   ├── css/              # Custom styling
│   └── js/               # JavaScript functionality
├── uploads/              # Temporary uploaded files
├── generated/            # Generated PowerPoint files
└── README.md             # This file
```

## API Endpoints

- `GET /` - Main upload interface
- `POST /upload` - Handle file upload and start processing
- `GET /status/<job_id>` - Check processing status
- `GET /download/<filename>` - Download generated files

## Configuration

### Environment Variables
- `FLASK_ENV` - Set to `development` for debug mode
- `FLASK_PORT` - Port to run the application (default: 5000)

### File Limits
- Maximum file size: 16MB
- Supported song file formats: `.txt`
- Supported template formats: `.pptx`
- File cleanup: 2 hours after creation

## Deployment

### Railway (Recommended)
1. Connect your GitHub repository to Railway
2. Railway will automatically detect the Flask app
3. Set environment variables if needed
4. Deploy with one click

### Render
1. Connect GitHub repository
2. Set build command: `pip install -r requirements.txt`
3. Set start command: `gunicorn app:app`
4. Deploy

### Local Development
```bash
# Install dependencies
pip3 install -r requirements.txt

# Run development server
python3 app.py

# Visit http://localhost:5000
```

## Technical Details

- **Framework**: Flask 3.1+
- **PowerPoint Library**: python-pptx 0.6.21
- **Frontend**: Bootstrap 5.3, Font Awesome 6.0
- **File Handling**: Werkzeug secure filename, UUID-based naming
- **Processing**: Background threading with job tracking
- **Security**: File validation, size limits, automatic cleanup

## Example Output

Generated PowerPoint presentations include:
- **Professional formatting** with church-appropriate design
- **Automatic slide numbering** (1/4, 2/4, 3/4, 4/4)
- **Brown-colored counters** matching church theme
- **Table of Contents** with clickable navigation
- **Template integration** preserving your church branding
- **Optimized fonts** for projection and readability

## Troubleshooting

### Common Issues
1. **"No songs found"** - Make sure song titles start with `#`
2. **"File too large"** - Maximum file size is 16MB
3. **"Processing failed"** - Check song file format and try again
4. **Template not working** - Ensure template is `.pptx` format

### File Format Help
- Songs must start with `# Song Title`
- Separate verses with blank lines
- Use plain text only (no special characters)
- Save as `.txt` format with UTF-8 encoding

## Support

For issues or questions:
1. Check the song file format examples
2. Verify file extensions (`.txt` for songs, `.pptx` for templates)
3. Try with a smaller file if processing fails
4. Check browser console for JavaScript errors

## License

This project is designed for church worship services and community use.