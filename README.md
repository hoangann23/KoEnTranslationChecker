# Korean-English Translation Validator

This tool highlights glossary terms in Korean and English documents for easy validation and comparison of translations.

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Folder structure:**
   ```
   KoEnTranslationChecker/
   ├── ko/              # Korean source documents
   ├── en/              # English translated documents
   ├── glossary/        # Glossary files (Excel)
   ├── output/          # Highlighted output documents (auto-created)
   ├── run.py           # Main entry point
   └── ...
   ```

3. **File Requirements:**
   - Place Korean document in `ko/` folder (default: `ko/KO-Test.docx`)
   - Place English document in `en/` folder (default: `en/EN-Test.docx`)
   - Place glossary in `glossary/` folder (default: `glossary/L2M-OOG-Lingo-0313.xlsx`)
   - Glossary should have Korean terms in column F and English translations in column G

## Usage

### Quick Start

Run with default file names:

```bash
python3 run.py
```

This will process:
- `ko/KO-Test.docx` → highlighted to `output/KO-Test_Highlighted.docx`
- `en/EN-Test.docx` → highlighted to `output/EN-Test_Highlighted.docx`

### Custom File Paths

Specify custom document paths:

```bash
python3 run.py --korean ko/my-korean.docx --english en/my-english.docx
```

### All Options

```bash
python3 run.py --help
```

Options:
- `--korean, -k` - Path to Korean document (default: `ko/KO-Test.docx`)
- `--english, -e` - Path to English document (default: `en/EN-Test.docx`)
- `--glossary, -g` - Path to glossary file (default: `glossary/L2M-OOG-Lingo-0313.xlsx`)
- `--korean-output, -ko` - Output path for highlighted Korean document (default: `output/KO-Test_Highlighted.docx`)
- `--english-output, -eo` - Output path for highlighted English document (default: `output/EN-Test_Highlighted.docx`)

### Example Commands

```bash
# Use defaults
python3 run.py

# Custom glossary
python3 run.py --glossary glossary/custom-lingo.xlsx

# Different input files
python3 run.py -k ko/document-v2.docx -e en/document-v2.docx

# Custom output names
python3 run.py --korean-output output/korean-checked.docx --english-output output/english-checked.docx
```

## Output

The script generates two highlighted documents in the `output/` folder:

1. **Highlighted Korean Document** - All glossary terms highlighted in yellow
   - Includes Korean particles filtered out from matching (는, 은, 이, 가, etc.)
   - Processes paragraphs and table content
   
2. **Highlighted English Document** - All glossary terms highlighted in yellow
   - Case-insensitive matching for English terms
   - Processes paragraphs and table content

## Features

- ✅ **Highlights glossary terms** in both documents (yellow background)
- ✅ **Processes tables** - not just regular paragraphs
- ✅ **Filters Korean particles** - automatically ignores common particles
- ✅ **Case-insensitive matching** for English terms
- ✅ **Avoids overlapping matches** - longer terms have priority
- ✅ **Summary statistics** - shows total occurrences and unique terms found
- ✅ **Command-line arguments** - specify custom file paths easily
- ✅ **Organized folders** - keeps inputs and outputs organized

## What It Does

1. Loads glossary terms from Excel file
2. Processes Korean document:
   - Finds all instances of Korean glossary terms
   - Highlights them in yellow
   - Filters out Korean particles
3. Processes English document:
   - Finds all instances of English glossary terms
   - Highlights them in yellow
   - Case-insensitive matching
4. Saves highlighted documents to `output/` folder
5. Displays summary: total occurrences and unique terms in each document

## File Scripts

- **run.py** - Main orchestrator script (runs both highlighters)
- **highlight_korean.py** - Highlights Korean glossary terms
- **highlight_english.py** - Highlights English glossary terms

Can also run individual scripts:
```bash
python3 highlight_korean.py    # Just process Korean
python3 highlight_english.py   # Just process English
```
