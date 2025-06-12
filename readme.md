# UI Automation Assistant

This project contains two main components for UI automation:

1. A GUI-based automation recorder for Windows applications using PyWinAuto
2. A Chrome browser automation utility using Playwright

## PyWinAuto Assistant (asistente_windows_ui.py)

A graphical tool to record and playback UI automation sequences for Windows applications.

### Features

- **Window Selection**: Easily select target windows from a list of open windows
- **Element Identification**: Visual identification of interactive elements with color coding
  - Red outlines: Interactive elements (buttons, textboxes, etc.)
  - Gold outlines: Non-interactive elements
- **Action Recording**:
  - Single and double clicks
  - Text input
  - ComboBox selection
  - Configurable delays
- **Action Management**:
  - Execute individual actions
  - Delete actions
  - Clear all actions
- **Code Generation**:
  - Generate Python code from recorded actions
  - Save generated code to file
  - Execute generated code directly

### Usage

```python
from asistente_windows_ui import PyWinAutoAssistant

if __name__ == "__main__":
    app = PyWinAutoAssistant()
    app.run()
```

## Local Chrome Browser (local_chrome_browser.py)

A Playwright-based Chrome automation utility with built-in stealth features and certificate support.

### Features

- **Context Manager Interface**: Easy to use with `with` statements
- **Certificate Support**: Handle client certificates for authenticated sites
- **Stealth Mode**: Anti-detection measures for automated browsing
- **Session Management**: Prevents restore prompts and maintains clean sessions
- **Random Delays**: Built-in random delay function for natural interaction

### Usage

Basic usage:
```python
from local_chrome_browser import LocalChromeBrowser

with LocalChromeBrowser() as browser:
    browser.goto("https://example.com")
    browser.random_delay()
    browser.page.click("text=SomeButton")
```

With certificates:
```python
with LocalChromeBrowser(
    origin="https://your-server.com",
    pfxPath="./path/to/cert.pfx",
    passphrase="your-password"
) as browser:
    browser.goto("https://your-server.com")
```

## Requirements

- Python 3.11+
- pywinauto
- playwright
- tkinter
- Pillow (PIL)
- pywin32

## Installation

```bash
pip install pywinauto playwright pillow pywin32
```

## Notes

- Chrome must be installed for the LocalChromeBrowser to work
- Both tools are designed for Windows environments