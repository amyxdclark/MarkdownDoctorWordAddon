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
- **.NET 10 SDK** or later
- Microsoft Word Desktop (2016 or later, Microsoft 365) or Word Online

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

5. Sideload the add-in to Word (see [SIDELOAD.md](SIDELOAD.md) for detailed instructions):
   
   **For Microsoft 365 Desktop Word (Recommended Method):**
   - Go to **File** > **Options** > **Add-ins**
   - At the bottom, change the **"Manage:"** dropdown to **"Office Add-ins"** (NOT "Word Add-ins" or "COM Add-ins")
   - Click **Go...** then **Upload My Add-in**
   - Browse to the `manifest.xml` file in the project folder
   - Click **Upload**
   
   **Alternative Method:**
   - In Word, go to **Insert** > **Add-ins** > **My Add-ins** > **Upload My Add-in**
   - Browse to the `manifest.xml` file in the project folder
   - Click **Upload**
   
   > **Important Note**: This is an **Office Add-in** (Web Add-in), not a COM Add-in or Word Add-in template. When using File > Options > Add-ins, make sure to select **"Office Add-ins"** from the "Manage:" dropdown, NOT "Word Add-ins" (which is for document templates). See [SIDELOAD.md](SIDELOAD.md) for detailed step-by-step guidance with explanations.

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

4. Sideload the add-in to Word:
   - See [SIDELOAD.md](SIDELOAD.md) for detailed instructions
   - **Recommended for Desktop Word**: Go to **File** > **Options** > **Add-ins**, change "Manage:" to **"Office Add-ins"**, click **Go**, then **Upload My Add-in**
   - **Alternative**: Go to **Insert** > **Add-ins** > **My Add-ins** > **Upload My Add-in**
   
   > **Important Note**: This is an **Office Add-in** (Web Add-in). When using File > Options > Add-ins, select **"Office Add-ins"** from the "Manage:" dropdown, NOT "Word Add-ins" (which is for document templates) or "COM Add-ins" (which is for legacy add-ins).

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

The published files will be in the `bin/Release/net10.0/publish/` directory.

## Troubleshooting

### Can't find where to add the add-in

This is an **Office Add-in** (also called a "Web Add-in"), not a COM Add-in or Word Add-in template.

**Recommended approach for Microsoft 365 Desktop Word:**
1. Go to **File** > **Options** > **Add-ins**
2. At the bottom, find the **"Manage:"** dropdown
3. Select **"Office Add-ins"** (NOT "Word Add-ins" or "COM Add-ins")
4. Click **Go...** button
5. Click **Upload My Add-in**

**Alternative approaches:**
- Try **Insert** > **Add-ins** > **My Add-ins** > **Upload My Add-in**
- Try **Insert** > **Get Add-ins** > **My Add-ins** > **Upload My Add-in**

**What NOT to use:**
- ❌ **"COM Add-ins"** - These are for legacy add-ins built with COM/VSTO
- ❌ **"Word Add-ins"** - These are for document templates (.dotm files), not Office Add-ins

See [SIDELOAD.md](SIDELOAD.md) for detailed instructions with multiple methods and troubleshooting help.

### Add-in doesn't appear in Word

1. Make sure the ASP.NET Core server is running (https://localhost:3000)
2. Try clearing Office's cache and restarting Word:
   - **Windows**: Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - **Mac**: Delete `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
3. Verify the manifest.xml file is valid
4. Re-sideload the add-in (see [SIDELOAD.md](SIDELOAD.md))

### Server not starting

1. Make sure port 3000 is not in use by another application
2. Check that .NET 10 SDK is installed: `dotnet --version`
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