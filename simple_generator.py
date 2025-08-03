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
import argparse


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
    # Use a simple blank layout to avoid placeholder conflicts
    # This ensures consistent behavior across different templates
    try:
        # Use blank layout (index 6) or last available layout
        if len(prs.slide_layouts) > 6:
            layout = prs.slide_layouts[6]  # Blank layout
        else:
            layout = prs.slide_layouts[-1]  # Last available layout
    except (IndexError, AttributeError):
        # Fallback to first available layout
        layout = prs.slide_layouts[0]
    
    slide = prs.slides.add_slide(layout)
    
    # Always use manual text boxes for consistent positioning
    # This avoids conflicts with template placeholders
    
    # Add title at top-left, with distance from red line
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.6),  # More distance from red line at top
        prs.slide_width - Inches(1.0), Inches(1.0)
    )
    title_frame = title_box.text_frame
    title_frame.margin_left = Inches(0.2)
    title_frame.margin_right = Inches(0.2)
    title_frame.vertical_anchor = MSO_ANCHOR.TOP
    title_frame.word_wrap = True
    
    title_p = title_frame.paragraphs[0]
    title_p.text = title
    title_p.font.size = Pt(32)  # Reduced from 40 to 32
    title_p.font.name = "Arial"
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(0, 0, 0)  # Black
    title_p.alignment = PP_ALIGN.LEFT
    
    # Add content below title, positioned much closer
    if content_lines:
        content_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(1.4),  # Much closer to title (reduced from 1.8 to 1.4)
            prs.slide_width - Inches(1.0), prs.slide_height - Inches(1.9)  # Adjusted height accordingly
        )
        content_frame = content_box.text_frame
        content_frame.margin_left = Inches(0.2)  # Match title margin
        content_frame.margin_right = Inches(0.2)
        content_frame.vertical_anchor = MSO_ANCHOR.TOP
        content_frame.word_wrap = True
        
        # Add each line with left alignment
        for i, line in enumerate(content_lines):
            if i == 0:
                p = content_frame.paragraphs[0]
            else:
                p = content_frame.add_paragraph()
            
            p.text = line
            p.font.size = Pt(28)  # Reduced from 32 to 28
            p.font.name = "Arial"
            p.font.color.rgb = RGBColor(0, 0, 0)  # Black
            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(16)


def main():
    print("Simple PowerPoint Song Generator")
    print("=" * 40)
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="Generate PowerPoint presentations from song collections",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python3 simple_generator.py songs.txt
  python3 simple_generator.py songs.txt output.pptx
  python3 simple_generator.py songs.txt --master template.pptx
  python3 simple_generator.py songs.txt output.pptx --master "Master Folie Natal.pptx" """
    )
    
    parser.add_argument('input_file', help='Input text file containing songs')
    parser.add_argument('output_file', nargs='?', default='songs_presentation.pptx',
                       help='Output PowerPoint file (default: songs_presentation.pptx)')
    parser.add_argument('--master', metavar='TEMPLATE', 
                       help='Use existing PowerPoint file as template')
    
    args = parser.parse_args()
    
    input_file = args.input_file
    output_file = args.output_file
    master_file = args.master
    
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
    if master_file:
        print(f"Creating PowerPoint presentation using template: {master_file}")
        try:
            prs = Presentation(master_file)
        except FileNotFoundError:
            print(f"Error: Template file '{master_file}' not found!")
            return
        except Exception as e:
            print(f"Error loading template: {e}")
            return
    else:
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