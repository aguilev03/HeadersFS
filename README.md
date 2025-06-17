# HeadersFS
allows for the copy and paste of the new headers

# Stage & Commodity Lookup Tool

A simple AutoHotkey-based GUI application that generates standardized claim notes and labels based on combinations of:

- Stage (e.g., First Time Install, Demo, Reselect)
- Commodity (e.g., Carpet, Tile Wall, Vinyl)
- Prefix (optional: RO, OOW, OCC, etc.)
- Override Mode (with Category + Material dropdowns)

---

## Features

- Lightweight GUI with 3 main dropdowns (Stage, Commodity, Prefix)
- Auto-generated results based on data from an Excel-style lookup table
- Copy to clipboard with one click
- **Override Mode**: Use alternate Category + Material combo to bypass normal logic
- Clean all-caps formatting (e.g., `OOW TILE WALL COMP`)

---

## How to Use

1. Double-click the `.exe` or run the `.ahk` file (requires [AutoHotkey v2](https://www.autohotkey.com/))
2. Select:
    - **Stage**
    - **Commodity**
    - **Prefix** (optional)
3. (Optional) Enable **Override Mode**
    - Pick a **Category** (e.g., "Trip Charge")
    - Pick a **Material** (e.g., "Tile")
4. The generated phrase will appear in the box below
5. Click “Copy to Clipboard” to use it in your workflow

---

## Installation

- Download the `.exe` file — no installation required
- Or clone the repo and run the `.ahk` script (AutoHotkey v2 required)

---

## License

This project is licensed under the [MIT License](https://www.notion.so/LICENSE).

> You are free to use, modify, and distribute this tool — no warranty provided.
> 

---

## Icon License

The clipboard icon was custom-generated using AI and is free for both personal and commercial use.

---

## Author

Built by Evan Aguilar
