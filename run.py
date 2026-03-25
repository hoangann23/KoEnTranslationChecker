#!/usr/bin/env python3
"""
Korean-English Translation Validator
Highlights glossary terms in both Korean and English documents
Accepts file paths from command line arguments
"""

import subprocess
import sys
import os
import argparse

def main():
    parser = argparse.ArgumentParser(
        description="Highlight glossary terms in Korean and English documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 run.py
  python3 run.py --korean ko/KO-Test.docx --english en/EN-Test.docx --glossary glossary/L2M-OOG-Lingo-0313.xlsx
  python3 run.py --korean ko/custom.docx --english en/custom.docx
        """)
    
    parser.add_argument(
        '--korean', '-k',
        default='ko/KO-Test.docx',
        help='Path to Korean document (default: ko/KO-Test.docx)'
    )
    parser.add_argument(
        '--english', '-e',
        default='en/EN-Test.docx',
        help='Path to English document (default: en/EN-Test.docx)'
    )
    parser.add_argument(
        '--glossary', '-g',
        default='glossary/L2M-OOG-Lingo-0313.xlsx',
        help='Path to glossary file (default: glossary/L2M-OOG-Lingo-0313.xlsx)'
    )
    parser.add_argument(
        '--korean-output', '-ko',
        default='output/KO-Highlighted.docx',
        help='Output path for highlighted Korean document'
    )
    parser.add_argument(
        '--english-output', '-eo',
        default='output/EN-Highlighted.docx',
        help='Output path for highlighted English document'
    )
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("Korean-English Translation Validator")
    print("=" * 80)
    
    # Check if files exist
    for filepath, name in [(args.korean, 'Korean document'),
                           (args.english, 'English document'),
                           (args.glossary, 'Glossary')]:
        if not os.path.exists(filepath):
            print(f"❌ Error: {name} not found: {filepath}")
            sys.exit(1)
    
    print(f"\n📄 Input files:")
    print(f"  Korean:  {args.korean}")
    print(f"  English: {args.english}")
    print(f"  Glossary: {args.glossary}")
    print(f"\n📝 Output files:")
    print(f"  Korean:  {args.korean_output}")
    print(f"  English: {args.english_output}")
    
    # Check and create output directory if it doesn't exist
    output_dir = os.path.dirname(args.korean_output)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"\n✓ Created output directory: {output_dir}")
        except OSError as e:
            print(f"❌ Error: Could not create output directory: {output_dir}")
            print(f"   Details: {e}")
            sys.exit(1)
    
    # Check if dependencies are installed
    try:
        import docx
        import pandas
        print("\n✓ Dependencies already installed")
    except ImportError:
        print("\n📦 Installing dependencies...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✓ Dependencies installed")
    
    # Import the highlighter functions
    from highlight_korean import create_highlighted_korean_doc
    from highlight_english import create_highlighted_doc
    
    print("\n" + "=" * 80)
    print("Processing Korean document...")
    print("=" * 80)
    ko_terms = create_highlighted_korean_doc(
        ko_doc_path=args.korean,
        glossary_path=args.glossary,
        output_path=args.korean_output
    )
    
    print("\n" + "=" * 80)
    print("Processing English document...")
    print("=" * 80)
    en_terms = create_highlighted_doc(
        en_doc_path=args.english,
        glossary_path=args.glossary,
        output_path=args.english_output
    )
    
    # Summary
    print("\n" + "=" * 80)
    print("Summary")
    print("=" * 80)
    print(f"\nKorean Document:")
    print(f"  Total occurrences: {len(ko_terms)}")
    print(f"  Unique terms: {len(set(t['korean'] for t in ko_terms))}")
    
    print(f"\nEnglish Document:")
    print(f"  Total occurrences: {len(en_terms)}")
    print(f"  Unique terms: {len(set(t['english'] for t in en_terms))}")
    
    print(f"\n✨ Done! Check the output files:")
    print(f"  - {args.korean_output}")
    print(f"  - {args.english_output}")

if __name__ == "__main__":
    main()

