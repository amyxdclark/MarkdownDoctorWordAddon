# Markdown Doctor - Word Add-in

A Microsoft Word add-in that allows you to convert markdown text to Word format by simply highlighting the text and clicking a button in the toolbar.

## Features

- **Easy to Use**: Highlight markdown text in your document and click the "Convert Markdown" button
- **Comprehensive Support**: Supports headers, bold, italic, lists, blockquotes, code, and links
- **In-place Conversion**: Converts text directly in your document without copy-pasting
- **Beautiful UI**: Clean, modern interface integrated into Word

## Supported Markdown Syntax

- **Headers**: `# H1`, `## H2`, `### H3`, etc.
- **Bold**: `**text**` or `__text__`
- **Italic**: `*text*` or `_text_`
- **Bold Italic**: `***text***` or `___text___`
- **Unordered Lists**: `- item` or `* item`
- **Ordered Lists**: `1. item`, `2. item`, etc.
- **Code**: `` `inline code` `` or ` ```code block``` `
- **Blockquotes**: `> quote`
- **Links**: `[text](url)`

## Installation

### Prerequisites

- Node.js (version 14 or higher)
- Microsoft Word (2016 or later, Microsoft 365, or Word Online)
- **Optional**: Visual Studio 2022 (or later) for IDE support - See [VISUAL_STUDIO.md](VISUAL_STUDIO.md)

### Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/amyxdclark/MarkdownDoctorWordAddon.git
   cd MarkdownDoctorWordAddon
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Generate SSL certificates for local development:
   ```bash
   npx office-addin-dev-certs install
   ```

4. Start the development server:
   ```bash
   npm run dev-server
   ```

5. In another terminal, sideload the add-in to Word:
   ```bash
   npm start
   ```

This will open Microsoft Word with the add-in loaded.

## Usage

1. Open Microsoft Word with the add-in loaded
2. Look for the "Markdown Doctor" button in the Home tab ribbon
3. Click the button to open the task pane
4. Highlight any markdown text in your document
5. Click the "Convert Markdown" button in the task pane
6. Your markdown will be converted to formatted Word text!

## Example

Try converting this markdown text:

```markdown
# My Document

## Introduction

This is a **bold** statement with *italic* text.

### Features

- First item
- Second item
- Third item

> This is a blockquote

Here's some `inline code` and a [link](https://example.com).
```

## Development

### Project Structure

```
MarkdownDoctorWordAddon/
├── manifest.xml          # Add-in manifest file
├── package.json          # Node.js dependencies
├── webpack.config.js     # Webpack configuration
├── taskpane.html         # Main UI
├── taskpane.css          # Styling
├── taskpane.js           # Main logic
├── commands.html         # Command functions
├── assets/               # Icons and images
└── README.md            # This file
```

### Key Files

- **manifest.xml**: Defines the add-in's capabilities and appearance in Word
- **taskpane.html**: The user interface shown in the task pane
- **taskpane.js**: Contains the markdown parsing and Word formatting logic
- **taskpane.css**: Styles for the task pane UI

### Building for Production

To create a production build:

```bash
npm run build
```

The built files will be in the `dist/` directory.

## Troubleshooting

### Add-in doesn't appear in Word

1. Make sure the development server is running (`npm run dev-server`)
2. Try clearing Office's cache and restarting Word
3. Check that the manifest is valid: `npm run validate`

### Conversion not working

1. Make sure you have text selected before clicking "Convert Markdown"
2. Check the browser console (F12) for any errors
3. Ensure your markdown syntax is correct

### SSL Certificate Issues

If you see SSL certificate warnings:

```bash
npx office-addin-dev-certs install --force
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - feel free to use this project for any purpose.

## Support

For issues or questions, please open an issue on GitHub: https://github.com/amyxdclark/MarkdownDoctorWordAddon/issues