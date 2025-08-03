#!/usr/bin/env python3
"""Test script for the PowerPoint generator."""

import os
from generator import generate_presentation

def test_generator():
    """Test the generator with sample files."""
    print("🧪 Testing PowerPoint Generator...")
    print("=" * 50)
    
    # Test files
    song_file = "kumpulan_lagu_ekklesia.txt"
    template_file = "Master Folie Natal.pptx"
    output_file = "test_output.pptx"
    
    # Check if input files exist
    if not os.path.exists(song_file):
        print(f"❌ Song file not found: {song_file}")
        return False
    
    if not os.path.exists(template_file):
        print(f"⚠️  Template file not found: {template_file}")
        template_file = None
    
    print(f"📄 Song file: {song_file}")
    print(f"🎨 Template: {template_file or 'None (using default)'}")
    print(f"📊 TOC: Yes")
    print("-" * 50)
    
    # Test the generator
    success, message, slide_count = generate_presentation(
        song_file, 
        output_file, 
        template_file, 
        generate_toc=True
    )
    
    if success:
        print(f"✅ {message}")
        print(f"📋 Output file: {output_file}")
        
        # Check if file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file) / (1024 * 1024)  # MB
            print(f"📏 File size: {file_size:.1f} MB")
        
        return True
    else:
        print(f"❌ {message}")
        return False

if __name__ == "__main__":
    test_generator()