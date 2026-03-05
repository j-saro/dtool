# dtool

`dtool` is a Python utility package designed for splitting and merging `.docx` files. It provides granular control over how documents are partitioned and facilitates the re-assembly of these fragments.

## Features
*   **Split:** Partition large `.docx` files based on specific character or word counts with adjustable boundary logic.
*   **Merge:** Combine multiple `.docx` files from a folder or an explicit list.
*   **Encryption Removal:** Optional utility to strip encryption settings from the `settings.xml` within the docx structure.

---

## Installation

It is highly recommended to use a virtual environment to manage dependencies.

1. **Create and activate a virtual environment:**
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

Check out `demo.py` for a quick start. Here is a breakdown of how to use the functions.

### Splitting Files
To split a document, use the `split_docx` function with a `Config` object:

```python
from dtool import split_docx, Config, Boundary, Unit

split_docx(
    Config(
        file_path="document.docx",
        output_path="output_folder",
        count=90000,
        unit=Unit.CHARS,             # Options: Unit.CHARS, Unit.WORDS
        boundary=Boundary.NEAREST_SENTENCE # Options: Boundary.STRICT, Boundary.NEAREST_WORD, Boundary.NEAREST_SENTENCE
    )
)
```

**Configuration Notes:**
*   **Boundary Logic:** `STRICT` is precise, potentially cutting inside a word. `NEAREST_WORD` and `NEAREST_SENTENCE` attempt to split at logical breakpoints to maintain readability.
*   **Count Accuracy:** Please note that the `count` is **not identical** to word counts provided by Microsoft Word or LibreOffice. `dtool` does not count text exactly the same way these applications do.

### Merging Files
You can merge files by providing a folder path or an explicit list of file paths.

```python
from dtool import merge_docx

# Merge from a folder (Sorted automatically)
merge_docx(input_files="folder_path", output_path="out.docx", remove_encryption=True)

# Merge from a specific list (Preserves list order)
files = ["0002.docx", "0001.docx"]
merge_docx(input_files=files, output_path="out.docx")
```

**Important Notes for Merging:**
*   **Folder Sorting:** When providing a folder path, files are sorted using Python's default sorting. Keep in mind that string-based sorting puts `10_doc` *before* `2_doc`. To ensure correct order, use naming conventions like `0001_name.docx`.
*   **List Order:** When passing a Python list, the files are merged in the exact order provided. No automatic sorting is performed.

---

## Known Issues

*   **Microsoft Word Compatibility:** Due to the way the files are generated, Microsoft Word may occasionally throw an error when opening the resulting documents because the output is not 100% standard. **However**, after using Word's built-in "Repair" feature, the documents open and function without issues.
*   **LibreOffice:** LibreOffice opens the generated documents without any errors or warnings.

---
