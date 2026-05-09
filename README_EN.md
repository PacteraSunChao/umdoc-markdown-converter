# UMDoc -> Markdown

> **Language**: [简体中文](README.md) | English | [日本語](README_JP.md)

A modern cross-platform (Windows & macOS) desktop tool that converts almost any common document into clean, structured Markdown with a single click.
Powered by Microsoft's open-source MarkItDown engine with the [all] extra dependencies installed, it supports over 30 file formats including Office, PDF, images, audio, ebooks, and more.

## Features

- Universal conversion: Converts Word, Excel, PowerPoint, PDF, images, audio, EPUB, HTML, CSV, JSON, and many other formats into Markdown.
- Drag-and-drop simplicity: Just drag a file onto the window or use the file picker; the tool automatically detects the format and converts it.
- Smart Excel handling: When the file is an Excel workbook, a sheet list appears so you can select which sheets to convert. For all other formats, the sheet selection area is hidden, keeping the interface clean.
- Dual-view preview: Switch instantly between "Source" and "Rendered" views to see the raw Markdown or the formatted output.
- Flexible saving: Save as a Markdown file with a name of your choice, or use "Save to" to pick a folder and automatically generate a `.md` file named after the original document.
- Modern, clean interface: Light theme with rounded card design for a comfortable visual experience.
- Automatic multi-language detection: The app detects your system language on startup and supports Chinese, Japanese, and English, with manual switching also available from the menu.
- Automatic environment management: The `launcher.py` script installs all required dependencies (PySide6, markitdown[all], etc.) into an isolated virtual environment inside the project folder, without polluting your system Python.

## Screenshots

![cn](./screenshots/cn.png)

![en](./screenshots/en.png)

![jp](./screenshots/jp.png)

## Quick Start

### Prerequisites

- Python 3.9+
  macOS: Install via Homebrew with `brew install python`
  Windows: Download from python.org

### Installation and Running

1. Clone the repository
   ```
   git clone git@github.com:PacteraSunChao/umdoc.git
   cd umdoc
   ```

2. Run with the launcher (recommended)
   ```
   python3 launcher.py
   ```
   The first run creates a virtual environment `app_env` inside the project folder and installs all dependencies (PySide6, markitdown[all], etc.). Subsequent runs launch the main window directly without reinstallation.

3. Or install dependencies manually and run the main program
   ```
   python3 -m venv app_env
   source app_env/bin/activate   # macOS/Linux
   app_env\Scripts\activate      # Windows
   pip install PySide6 "markitdown[all]"
   python3 umdoc.py
   ```

### How to Use

- Launch the app and drag any supported file (e.g., .docx, .pdf, .xlsx, .jpg) into the file area, or click the button to choose a file.
- If the file is an Excel workbook, a sheet list will appear; select the sheets you want to convert (Select All / Deselect All is supported) and then click "Convert to Markdown".
- For other formats, the conversion runs automatically and the result is shown in the preview area.
- Use the "Source" and "Rendered" buttons to toggle between the raw Markdown and the formatted preview.
- Click "Save to" and choose a folder; the app will automatically create a `.md` file using the original file name. You can also use the menu "File -> Save Markdown" to save with a custom name and location.

## Supported File Formats

With the `markitdown[all]` dependencies installed, the tool handles the following common file types:

| Category | Formats |
| :--- | :--- |
| Office documents | .docx, .pptx, .xlsx, .xls |
| PDF | .pdf |
| Images | .jpg, .png, etc. (OCR text extraction supported) |
| Audio | .wav, .mp3 (requires Azure AI Speech service) |
| Ebooks | .epub |
| Web & data | .html, .csv, .json, .xml |
| Archives | .zip (automatically extracts and converts recognized files within) |
| Email | .msg (Outlook messages) |
| Other | YouTube links (converted to a Markdown description) |

Note: Advanced features such as audio transcription and YouTube link parsing require additional configuration of Azure AI or an LLM client (e.g., OpenAI). Without these configurations, MarkItDown will still extract available metadata.

## Project Structure

```
.
├── launcher.py          # Bootstrap script (creates virtual environment and installs dependencies)
├── umdoc.py             # Main application (GUI interface and conversion logic)
├── README_EN.md         # Project documentation
└── .gitignore           # Git ignore rules
```

## Technology Stack

- PySide6 - Cross-platform GUI framework
- MarkItDown - Microsoft's open-source document-to-Markdown engine
- openpyxl - Excel file reading (used to retrieve sheet names)
- Additional libraries automatically installed by `markitdown[all]` for PDF, image processing, HTML parsing, etc.

## Multi-language Support

The application supports three interface languages: Chinese, Japanese, and English.
It automatically selects the language based on your system locale on startup, and you can switch at any time via the menu: Language / 言语 / 言語.

## Contributing

Issues and pull requests are welcome to help improve this tool.
If you have feature ideas or find a bug, please open an issue on GitHub.

## License

This project is open source under the MIT License, which permits free use, modification, and distribution. See the LICENSE file for details.