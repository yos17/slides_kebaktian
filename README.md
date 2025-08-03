# Simple PowerPoint Song Generator

Automatically generates PowerPoint presentations from worship song collections using Python.

## Features

- **Simple & Fast**: One command to generate presentations
- **Web Interface**: User-friendly web app with drag & drop upload (see `webapp/`)
- **Template Support**: Use existing PowerPoint templates with --master option
- **Table of Contents**: Generate clickable TOC with --toc option
- **Slide Numbering**: Automatic numbering (1/4, 2/4, etc.) in brown color
- **Automatic Song Parsing**: Extracts 115+ songs from text file with # separators  
- **Natural Slide Breaks**: Uses paragraph breaks (empty lines) to split lyrics
- **Clean Formatting**: Calibri fonts with optimized sizing for readability
- **Multi-language Support**: Handles English, Indonesian, and German songs

## Quick Start

### Option 1: Web Interface (Recommended)
```bash
cd webapp/
pip3 install flask python-pptx
python3 app.py
```
Visit `http://localhost:5000` for drag & drop interface!

### Option 2: Command Line
```bash
pip3 install python-pptx
python3 simple_generator.py <input_file.txt> [output_file.pptx] [--master template.pptx] [--toc]
```

**Examples:**
```bash
# Generate with default filename (songs_presentation.pptx)
python3 simple_generator.py kumpulan_lagu_ekklesia.txt

# Generate with custom filename
python3 simple_generator.py kumpulan_lagu_ekklesia.txt my_worship_songs.pptx

# Use existing PowerPoint template
python3 simple_generator.py kumpulan_lagu_ekklesia.txt --master "Master Folie Natal.pptx"

# Generate with Table of Contents
python3 simple_generator.py kumpulan_lagu_ekklesia.txt --toc

# Use template and generate TOC
python3 simple_generator.py kumpulan_lagu_ekklesia.txt --master template.pptx --toc
```

## Output

- **Professional Design**: Song title in header, lyrics left-aligned for readability
- **Optimized Fonts**: 32pt Calibri Bold titles, 28pt Calibri content for congregation viewing
- **Natural Flow**: Slides break at paragraph boundaries for better readability
- **Interactive TOC**: Clickable table of contents with 2-column layout (when --toc used)
- **Template Integration**: Seamlessly works with existing PowerPoint templates
- **Ready to Use**: Generates 400+ slides from 115 songs in seconds

## Song Collection Format

Your `kumpulan_lagu_ekklesia.txt` uses this format:
```
# Song Title

First verse lyrics here
Multiple lines supported

Second verse or chorus
More content

# Next Song Title

Next song content...
```

## File Structure

```
slides_kebaktian/
├── simple_generator.py           # Command-line script
├── kumpulan_lagu_ekklesia.txt    # Song collection (115 songs)
├── Master Folie Natal.pptx      # Template reference
├── webapp/                       # Web application
│   ├── app.py                   # Flask web server
│   ├── generator.py             # Web-optimized generator
│   ├── templates/               # HTML templates
│   └── README.md                # Web app documentation
└── songs_presentation.pptx      # Generated output
```

## Technical Details

- **Library**: python-pptx for PowerPoint generation
- **Parsing**: Splits songs by # markers, slides by empty lines
- **Formatting**: Calibri font, black text, left alignment for better readability
- **Template Support**: Preserves template layouts while adding content
- **Navigation**: Hyperlink-based TOC for easy song navigation
- **Performance**: Processes 115 songs into 400+ slides in under 5 seconds

The system converts your worship song collection into clean, professional PowerPoint presentations suitable for church services.