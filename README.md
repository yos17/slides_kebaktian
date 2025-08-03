# Simple PowerPoint Song Generator

Automatically generates PowerPoint presentations from worship song collections using Python.

## Features

- **Simple & Fast**: One command to generate presentations
- **Automatic Song Parsing**: Extracts 115+ songs from text file with # separators  
- **Natural Slide Breaks**: Uses paragraph breaks (empty lines) to split lyrics
- **Clean Formatting**: Simple slides with title headers and centered content
- **Multi-language Support**: Handles English, Indonesian, and German songs

## Quick Start

### Prerequisites
```bash
pip3 install python-pptx
```

### Usage

```bash
python3 simple_generator.py <input_file.txt> [output_file.pptx]
```

**Examples:**
```bash
# Generate with default filename (songs_presentation.pptx)
python3 simple_generator.py kumpulan_lagu_ekklesia.txt

# Generate with custom filename
python3 simple_generator.py kumpulan_lagu_ekklesia.txt my_worship_songs.pptx

# Automatically adds .pptx extension if missing
python3 simple_generator.py kumpulan_lagu_ekklesia.txt sunday_service
```

## Output

- **Clean Design**: Song title in header, lyrics centered on slide
- **Natural Flow**: Slides break at paragraph boundaries for better readability
- **Large Fonts**: 36pt titles, 28pt content for congregation viewing
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
├── simple_generator.py           # Main script
├── kumpulan_lagu_ekklesia.txt    # Song collection (115 songs)
├── Master Folie Natal.pptx      # Template reference
└── songs_presentation.pptx      # Generated output
```

## Technical Details

- **Library**: python-pptx for PowerPoint generation
- **Parsing**: Splits songs by # markers, slides by empty lines
- **Formatting**: Arial font, black text, centered alignment
- **Performance**: Processes 115 songs into 400+ slides in under 5 seconds

The system converts your worship song collection into clean, professional PowerPoint presentations suitable for church services.