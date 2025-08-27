# PDF to Word Converter

This is a simple Python script to convert PDF files to Microsoft Word (.docx) documents. The script extracts text, tables, and images from the PDF and attempts to reconstruct them in a Word document.

## Features

- Extracts text from each page.
- Detects and extracts tables.
- Detects and extracts images.
- Simple command-line interface.

## Requirements

- Python 3
- `pip` for installing dependencies

## Installation

1.  Clone or download this repository.
2.  Install the necessary Python libraries using `pip`:

    ```bash
    pip install pdfplumber python-docx
    ```

## Usage

To convert a PDF file, run the `converter.py` script from your terminal. You need to provide the path to the input PDF file and the desired path for the output .docx file.

### Command

```bash
python converter.py <path_to_your_input.pdf> <path_to_your_output.docx>
```

### Example

```bash
python converter.py sample_syllabus.pdf converted_document.docx
```

This will create a new file named `converted_document.docx` in the same directory, containing the content from `sample_syllabus.pdf`.

## Note

This tool is a best-effort converter. The quality of the conversion, especially regarding layout, can vary greatly depending on the complexity and structure of the source PDF.
