# Tools for company "Brand Images"

## Photo Session Article Helper

A script to automate product photo naming and metadata handling during photo sessions.

Features:
- Reads article data from Excel spreadsheet
- Monitors directory for new photos
- Generate file names based on article numbers
- Adds IPTC metadata to photos

Assumptions:
- The generated file name is copied into the clipboard. The software who tranfers the photo from the camera to the computer is able to 
  - use the file name from the clipboard
  - copy the file into the watch folder

Python Dependencies:
- openpyxl: Excel file handling
- pasteboard: Clipboard operations
- pyexiftool: Photo metadata manipulation
- rich: Enhanced console interface
- watchdog: File system monitoring

OS Dependencies:
- exiftool
  - Download MacOS package, see: https://exiftool.org/index.html
  - To verify succesfull installation, start a terminal and run
    ```bash
    exiftool -ver
    ```
  - Allowed IPTC Fields: https://exiftool.org/TagNames/IPTC.html
- uv
  - Install UV Package Manager (without Homebrew)
    ```bash
    # Download UV binary
    curl -LsSf https://astral.sh/uv/install.sh | sh

    # Add UV to your PATH (for zsh)
    echo 'export PATH="$HOME/.cargo/bin:$PATH"' >> ~/.zshrc
    source ~/.zshrc
    ```
  - To verify succesfull installation, start a terminal and run
    ```bash
    uv --version
    ```

How to run this script (if downloaded)
```bash
uv run photo-session-article-helper.py
```

How to run this script (from Github)
```bash
uv run https://raw.githubusercontent.com/brot/brand-images-tools/refs/heads/main/photo-session-article-helper.py
```
