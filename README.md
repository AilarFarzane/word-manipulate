# Word Document Generation and Manipulation Scripts

This document provides an overview of the Python scripts used for generating and updating Microsoft Word documents.

## `add_docx.py`

### Overview

This script generates a Word document (`.docx`) based on content defined in JSON files and a Word template. It populates a template with front matter, chapter content, headers, footers, and page numbers.

### Features

-   **JSON-driven Content:** The script reads chapter text, titles, lists, and image paths from `content/00-frontmatter.json` and `content/01-chapter-01.json`.
-   **Template-based:** It starts from a `template/blank-template.docx` file.
-   **Placeholder Replacement:** Fills in front matter details (e.g., `{{{{author}}}}`, `{{{{title}}}}`) in the initial section of the template.
-   **Section Management:** Creates distinct sections with their own headers and footers.
-   **Dynamic Content Addition:** Adds paragraphs, lists, and images according to the structure of the input JSON files.
-   **Header/Footer Customization:** Adds bordered headers and footers with page numbers.

### Dependencies

-   `python-docx`

Install the required package using:
```bash
pip install python-docx
```

### Usage

Run the script directly from the command line. It will read the predefined JSON files and the template, and output the result to `out-template.docx`.

```bash
python add_docx.py
```

---

## `update_toc.py`

### Overview

This script is designed to update the Table of Contents (TOC) and other fields within an existing Word document. It uses the Microsoft Word application on Windows to perform these updates, which is essential for correctly rendering fields like the TOC.

**Note:** This script requires a Windows environment with Microsoft Word installed.

### Features

-   **Field Updates:** Automatically updates all dynamic fields in the document, including the Table of Contents.
-   **TOC Level Configuration:** Sets the heading levels to be included in the TOC.
-   **Numbering Fix:** Includes a function (`fix_toc_numbering`) to correct formatting for chapter and section numbers in the TOC (e.g., changing `1-2 ` to `1-2- `).
-   **COM Automation:** Leverages the `win32com` library to programmatically control the MS Word application.

### Dependencies

-   `pywin32`

Install the required package using:
```bash
pip install pywin32
```

### Usage

This is a command-line tool that takes an input Word document and an optional output path.

**Syntax:**
```bash
python update_toc.py <input_docx> [output_docx]
```

-   `input_docx`: The path to the Word document that needs to be updated.
-   `output_docx` (Optional): The path to save the updated file. If not provided, the input file will be overwritten.

**Example:**
```bash
python update_toc.py out-template.docx final-document.docx
```
