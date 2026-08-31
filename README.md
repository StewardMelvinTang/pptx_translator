# PPT Translator

A multilingual desktop app for translating text in PowerPoint and PDF documents. The compact CustomTkinter interface supports light, dark, and system appearance on both macOS and Windows.

## Desktop workflow

- Drag PPTX or PDF files anywhere onto the window, or use **Choose files**.
- Review slide/page counts, file sizes, and live per-file translation states.
- Choose or search for source and target languages from compact dropdowns; the animated arrow button swaps them.
- Configure the local API key, translation model, and appearance from **Settings**.
- Translate a batch sequentially and monitor real paragraph-level progress.
- Reopen completed translations from the built-in history, including their full metrics and file locations.
- Ask AI about a translated file inside the main window with rendered Markdown, code blocks, tables, and clickable links.
- Continue a chat later: conversation history is saved separately for each exact translated file.

## Language selection

- Choose from 45 source and target languages.
- Leave the source set to **Auto-detect** to have the model identify each document's dominant language before translating it.
- A batch can contain documents in different source languages; detection runs per file.
- Use the arrow button to exchange the source and target. When Auto-detect is selected, swapping becomes available after one source language has been detected unambiguously.
- PowerPoint text boxes, grouped shapes, and table cells are included in translation.

## Run from source

Python 3.11 or 3.12 is recommended for the broadest binary-wheel and packaging compatibility.

### Windows

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python pptxtranslator.py
```

### macOS

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python pptxtranslator.py
```

## Build a desktop app

Install the build dependencies and run the cross-platform build helper:

```bash
python -m pip install -r requirements-build.txt
python build_app.py
```

The build uses PyInstaller's one-folder mode so CustomTkinter assets and native libraries load reliably and startup remains quick.

- Windows output: `dist\DocumentTranslator\DocumentTranslator.exe`
- macOS output: `dist/DocumentTranslator.app`

PyInstaller packages for the operating system it runs on. Build the Windows app on Windows and the macOS app on macOS; the same source and build command are used on both.

For public distribution, the next release step is to code-sign the Windows executable and sign/notarize the macOS app. An installer can then wrap the Windows output folder, while the macOS `.app` can be distributed in a signed DMG.

## Local data

Preferences and the API key are stored in the current user's native application-data folder:

- Windows: `%APPDATA%\Document Translator\config.json`
- macOS: `~/Library/Application Support/Document Translator/config.json`

The application remembers the selected theme, model, and most recent input/output folders.
Translation history and per-file chat sessions are also stored in this app-data directory.
