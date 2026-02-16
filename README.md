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

- **Visual Studio 2026** (or later) with the following workloads:
  - **ASP.NET and web development** workload
  - **Office/SharePoint development** workload (optional, for advanced Office development tools)
- **.NET 8.0 SDK** or later
- Microsoft Word (2016 or later, Microsoft 365, or Word Online)

### Setup with Visual Studio 2026

1. Clone this repository:
   ```bash
   git clone https://github.com/amyxdclark/MarkdownDoctorWordAddon.git
   cd MarkdownDoctorWordAddon
   ```

2. Open **MarkdownDoctorWordAddon.sln** in Visual Studio 2026

3. Build the solution:
   - Press `Ctrl+Shift+B` or go to Build > Build Solution

4. Run the project:
   - Press `F5` to start debugging
   - The ASP.NET Core server will start on https://localhost:3000

5. Sideload the add-in to Word:
   - In Word, go to Insert > Add-ins > Upload My Add-in
   - Browse to the `manifest.xml` file in the project folder
   - Click Upload

### Setup with .NET CLI (Alternative)

1. Clone this repository:
   ```bash
   git clone https://github.com/amyxdclark/MarkdownDoctorWordAddon.git
   cd MarkdownDoctorWordAddon
   ```

2. Build and run:
   ```bash
   dotnet restore
   dotnet build
   dotnet run
   ```

3. The server will start on https://localhost:3000

4. Sideload the add-in to Word as described above

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
├── MarkdownDoctorWordAddon.sln    # Visual Studio solution file
├── MarkdownDoctorWordAddin.csproj # C# project file
├── Program.cs                     # ASP.NET Core application entry point
├── Properties/
│   └── launchSettings.json        # VS launch configuration
├── appsettings.json              # Application configuration
├── appsettings.Development.json  # Development configuration
├── manifest.xml                  # Office Add-in manifest
├── wwwroot/                      # Static web files
│   ├── taskpane.html            # Main UI
│   ├── taskpane.css             # Styling
│   ├── taskpane.js              # Main logic
│   ├── commands.html            # Command functions
│   └── assets/                  # Icons and images
└── README.md                    # This file
```

### Key Files

- **manifest.xml**: Defines the add-in's capabilities and appearance in Word
- **Program.cs**: ASP.NET Core server configuration
- **wwwroot/taskpane.html**: The user interface shown in the task pane
- **wwwroot/taskpane.js**: Contains the markdown parsing and Word formatting logic
- **wwwroot/taskpane.css**: Styles for the task pane UI

### Building for Production

To create a production build:

```bash
dotnet publish -c Release
```

The published files will be in the `bin/Release/net8.0/publish/` directory.

## Troubleshooting

### Add-in doesn't appear in Word

1. Make sure the ASP.NET Core server is running (https://localhost:3000)
2. Try clearing Office's cache and restarting Word
3. Verify the manifest.xml file is valid

### Server not starting

1. Make sure port 3000 is not in use by another application
2. Check that .NET 8.0 SDK is installed: `dotnet --version`
3. Try running `dotnet restore` to restore dependencies

### SSL Certificate Issues

The ASP.NET Core development server uses a development certificate. If you see SSL warnings:

```bash
dotnet dev-certs https --trust
```

This command trusts the development certificate on your machine.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - feel free to use this project for any purpose.

## Support

For issues or questions, please open an issue on GitHub: https://github.com/amyxdclark/MarkdownDoctorWordAddon/issues