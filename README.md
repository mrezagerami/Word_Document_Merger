# Word Document Merger (.doc & .docx to .txt)

A robust Python script that extracts text from multiple Microsoft Word documents (`.doc` and `.docx`) in a specified directory and merges them into a single continuous text file. 

This tool is especially useful for combining datasets, transcripts, or split documents while strictly preserving the correct file order.

## ✨ Features

- **Multi-Format Support:** Reads both legacy Word files (`.doc`) and modern Word files (`.docx`).
- **Natural Sorting:** Intelligently sorts filenames containing numbers (e.g., `Podcast1`, `Podcast2`, ..., `Podcast10`) instead of using standard alphabetical sorting (which would incorrectly place `10` before `2`).
- **Clean Output:** Automatically adds clear visual separators and the original filename before the text of each document.
- **Safe Processing:** Automatically ignores temporary or hidden Word files (e.g., `~$document.docx`) to prevent crashes.

## 📋 Prerequisites

Due to the nature of legacy `.doc` files, this script relies on the Microsoft Word COM interface. Therefore, the following are required:

- **OS:** Windows
- **Software:** Microsoft Word must be installed on the machine.
- **Python:** Python 3.6 or higher.

### Required Python Libraries
You can install the required dependencies using `pip`:

```bash
pip install python-docx pywin32
