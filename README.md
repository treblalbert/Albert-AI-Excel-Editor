# 🤖 Albert's AI Excel Editor

An intelligent spreadsheet editor powered by AI (OpenAI GPT-4o-mini). Edit Excel files using natural language commands.

## Features

- 📊 Load Excel files (.xlsx, .xls, .csv)
- 💬 Natural language editing - Just tell the AI what to change
- ❓ Ask questions about your data
- 🔄 Compare changes - See before/after side by side
- ✨ Highlight AI changes - Toggle to see what was modified
- ↶ Undo support - Revert changes easily
- 🛡️ Row protection - Prevents accidental row deletion
- 🌐 Multi-language - English, Spanish, Catalan
- 📥 Export - Download modified spreadsheet

## Quick Start

### Prerequisites
- [Node.js](https://nodejs.org/) (v18 or higher)
- OpenAI API key

### Setup

```bash
# 1. Install dependencies
npm install

# 2. Test the app
npm start

# 3. Build installer (.exe)
npm run build
```

The installer will be in `dist/` folder.

## Adding Your Custom Icon

**IMPORTANT:** Before building, add your icon:

1. Create a 256x256 (or larger) PNG image
2. Convert to `.ico` at [convertio.co/png-ico](https://convertio.co/png-ico/) or [icoconvert.com](https://icoconvert.com/)
3. Save as `build/icon.ico`

The icon must be a proper `.ico` file with embedded sizes for Windows to display it correctly.

## Project Structure

```
albert-excel-editor/
├── src/
│   ├── index.html
│   └── assets/
│       ├── css/styles.css
│       └── js/
│           ├── app.js
│           └── translations.js
├── build/
│   └── icon.ico          <-- Put your icon here!
├── main.js
├── package.json
└── README.md
```

## Usage

1. **Load your API key** - Create a `.txt` file with your OpenAI API key
2. **Load an Excel file** - Select your spreadsheet
3. **Chat with the AI**:
   - "What's the total of column B?"
   - "Add 10% to all prices"
   - "Sort by date descending"
   - "Change 'Pending' to 'Completed' in Status column"
4. **Download** - Export your modified spreadsheet

## License

MIT License

## Author

Made with ❤️ by Albert
