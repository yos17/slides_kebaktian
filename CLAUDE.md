# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a simple Python-based PowerPoint song generator for church services. It converts a text file containing worship songs into a PowerPoint presentation with clean, readable slides.

## Essential Commands

### Prerequisites
```bash
pip3 install python-pptx
```

### Running the Generator
```bash
# Generate with default filename (songs_presentation.pptx)
python3 simple_generator.py kumpulan_lagu_ekklesia.txt

# Generate with custom filename
python3 simple_generator.py kumpulan_lagu_ekklesia.txt my_worship_songs.pptx

# Custom filename without .pptx extension (automatically added)
python3 simple_generator.py kumpulan_lagu_ekklesia.txt sunday_service

# Use template file for consistent branding/formatting
python3 simple_generator.py kumpulan_lagu_ekklesia.txt --master "Master Folie Natal.pptx"

# Template with custom output filename
python3 simple_generator.py kumpulan_lagu_ekklesia.txt christmas_songs.pptx --master "Master Folie Natal.pptx"
```

### Testing the System
```bash
# Test with sample input file
python3 simple_generator.py kumpulan_lagu_ekklesia.txt test_output.pptx
```

## Architecture

**Core Components:**
- `simple_generator.py` - Main script with three key functions:
  - `parse_songs()` - Parses text file using `#` markers as song separators
  - `split_lyrics_into_slides()` - Creates slide breaks at paragraph boundaries (empty lines)
  - `create_slide()` - Generates PowerPoint slides with title headers and centered content

**Input Format:** Text file with songs separated by `#` markers, paragraphs separated by empty lines
**Output:** PowerPoint presentation with clean formatting (Arial font, 36pt titles, 28pt content)

**Data Flow:** Text parsing → Song extraction → Slide splitting → PowerPoint generation

**Template Support:** 
- `--master` argument allows using existing PowerPoint files as templates
- Intelligently uses template placeholders (title/content) when available
- Falls back to manual text boxes when placeholders aren't found
- Preserves template themes, fonts, and formatting while maintaining readability

## File Structure

- `simple_generator.py` - Main generator script  
- `kumpulan_lagu_ekklesia.txt` - Song collection input (115 songs)
- `songs_presentation.pptx` - Default output file
- `Master Folie Natal.pptx` - Template reference