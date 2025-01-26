# PowerPoint Text Extractor

📊 A Python CLI tool that extracts text from PowerPoint presentations and exports it to CSV format. Perfect for content analysis, text mining, and presentation archiving.

A simple command-line tool to extract text content from PowerPoint presentations (.pptx/.ppt files) and save it to CSV format.

## Features

- Extracts text from all slides in a PowerPoint presentation
- Preserves slide numbers in the output
- Saves extracted text to a CSV file with the same name as the input file
- Handles both .pptx and .ppt file formats

## Installation

1. Ensure you have Python 3.x installed
2. Install the required dependencies:
```bash
pip install python-pptx
```

## Usage

```bash
python pptx-to-csv.py <path_to_powerpoint_file>
```

Example:
```bash
python pptx-to-csv.py presentation.pptx
```

This will create a CSV file named `presentation.csv` in the same directory as your input file.

## Output Format

The generated CSV file contains two columns:
- `slide_number`: The slide number in the presentation
- `text_content`: All text content extracted from that slide

## Error Handling

The script includes error handling for:
- Non-existent input files
- Invalid file formats (non-PowerPoint files)
- Processing errors during text extraction

## License

MIT License 