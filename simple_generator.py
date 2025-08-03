#!/usr/bin/env python3
"""
Simple PowerPoint Song Generator
Creates one presentation from kumpulan_lagu_ekklesia.txt with clean, simple slides.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import re


def parse_songs(file_path):
    """Parse songs from text file."""
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Split by song markers (lines starting with #)
    song_sections = re.split(r'\n(?=#)', content)
    songs = []
    
    for section in song_sections:
        if section.strip() and section.startswith('#'):
            lines = section.split('\n')
            
            if lines:
                # Extract title (remove #)
                title = lines[0].replace('#', '').strip()
                
                # Get lyrics (everything after title), preserving empty lines
                lyrics = lines[1:] if len(lines) > 1 else []
                
                if title:  # Only add if we have a title
                    songs.append({
                        'title': title,
                        'lyrics': lyrics
                    })
    
    return songs


def split_lyrics_into_slides(lyrics):
    """Split lyrics into slides at paragraph breaks (empty lines)."""
    slides = []
    current_slide = []
    
    for line in lyrics:
        if line.strip():  # Non-empty line
            current_slide.append(line)
        else:  # Empty line - natural paragraph break
            if current_slide:  # Only create slide if we have content
                slides.append(current_slide.copy())
                current_slide = []
    
    # Add remaining content if any
    if current_slide:
        slides.append(current_slide)
    
    return slides


def create_slide(prs, title, content_lines):
    """Create a simple slide with title and content."""
    # Use blank layout
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Add title at top
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.5),
        prs.slide_width - Inches(1), Inches(1.0)
    )
    title_frame = title_box.text_frame
    title_frame.margin_left = Inches(0.2)
    title_frame.margin_right = Inches(0.2)
    title_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    title_p = title_frame.paragraphs[0]
    title_p.text = title
    title_p.font.size = Pt(36)
    title_p.font.name = "Arial"
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(0, 0, 0)  # Black
    title_p.alignment = PP_ALIGN.CENTER
    
    # Add content below title (closer spacing)
    if content_lines:
        content_box = slide.shapes.add_textbox(
            Inches(1), Inches(1.3),
            prs.slide_width - Inches(2), prs.slide_height - Inches(2.5)
        )
        content_frame = content_box.text_frame
        content_frame.margin_left = Inches(0.2)
        content_frame.margin_right = Inches(0.2)
        content_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        # Add each line
        for i, line in enumerate(content_lines):
            if i == 0:
                p = content_frame.paragraphs[0]
            else:
                p = content_frame.add_paragraph()
            
            p.text = line
            p.font.size = Pt(28)
            p.font.name = "Arial"
            p.font.color.rgb = RGBColor(0, 0, 0)  # Black
            p.alignment = PP_ALIGN.CENTER
            p.space_after = Pt(12)


def main():
    import sys
    
    print("Simple PowerPoint Song Generator")
    print("=" * 40)
    
    # Parse command line arguments
    if len(sys.argv) < 2:
        print("Usage: python3 simple_generator.py <input_file.txt> [output_file.pptx]")
        print("Example: python3 simple_generator.py kumpulan_lagu_ekklesia.txt my_songs.pptx")
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "songs_presentation.pptx"
    
    # Ensure output file has .pptx extension
    if not output_file.endswith('.pptx'):
        output_file += '.pptx'
    
    # Parse songs
    print(f"Reading songs from {input_file}...")
    try:
        songs = parse_songs(input_file)
        print(f"Found {len(songs)} songs")
    except FileNotFoundError:
        print(f"Error: {input_file} not found!")
        return
    except Exception as e:
        print(f"Error reading file: {e}")
        return
    
    # Create presentation
    print("Creating PowerPoint presentation...")
    prs = Presentation()
    
    total_slides = 0
    
    for song in songs:
        title = song['title']
        lyrics = song['lyrics']
        
        # Split lyrics into slides
        lyric_slides = split_lyrics_into_slides(lyrics)
        
        # Create slides for this song
        for slide_content in lyric_slides:
            create_slide(prs, title, slide_content)
            total_slides += 1
    
    # Save presentation
    try:
        prs.save(output_file)
        print(f"✅ Success!")
        print(f"Created {output_file} with {total_slides} slides from {len(songs)} songs")
        print(f"Ready to use for church service!")
    except Exception as e:
        print(f"Error saving presentation: {e}")


if __name__ == "__main__":
    main()