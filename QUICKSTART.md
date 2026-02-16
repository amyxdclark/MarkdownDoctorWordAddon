# Quick Start Guide

## Getting Started in 5 Minutes

### Step 1: Build the Project

Using Visual Studio 2026:
1. Open `MarkdownDoctorWordAddon.sln` in Visual Studio
2. Press `Ctrl+Shift+B` to build

Or using .NET CLI:
```bash
dotnet restore
dotnet build
```

### Step 2: Trust Development Certificate (First time only)

For local development, trust the .NET development certificate:

```bash
dotnet dev-certs https --trust
```

### Step 3: Run the Server

Using Visual Studio 2026:
- Press `F5` to start debugging

Or using .NET CLI:
```bash
dotnet run
```

The server will run on https://localhost:3000

### Step 4: Sideload the Add-in to Word

1. Open Microsoft Word
2. Go to **Insert** > **Add-ins** > **Upload My Add-in** (or **My Add-ins**)
3. Browse to the `manifest.xml` file in the project folder
4. Click **Upload**
5. The add-in will load and appear in the Home tab

## Using the Add-in

### Opening the Task Pane

1. Look in the **Home** tab of Word's ribbon
2. Find the **Markdown Doctor** section
3. Click **Show Taskpane**

### Converting Markdown

1. Type or paste some markdown text in your Word document. For example:

```markdown
# My Title

This is **bold** and this is *italic*.

- Item 1
- Item 2
```

2. **Select/Highlight** the markdown text
3. Click the **Convert Markdown** button in the task pane
4. Your text will be converted to Word format!

## Example Workflow

Here's a complete example:

1. Start the server (`dotnet run` or `F5` in Visual Studio)
2. Open Word and sideload the add-in (see Step 4)
3. Type this markdown:

```markdown
## Project Status

The project is **on track** and we have completed:

1. Design phase
2. Development
3. Testing

> Note: Final review scheduled for next week.
```

4. Select all the text you just typed
5. Click "Convert Markdown" in the task pane
6. See your markdown transformed into formatted Word text!

## Supported Markdown

✅ **Bold** (`**text**`)  
✅ *Italic* (`*text*`)  
✅ Headers (`# H1`, `## H2`, etc.)  
✅ Lists (`- item`, `1. item`)  
✅ Blockquotes (`> quote`)  
✅ Code (`` `code` ``)  
✅ Links (`[text](url)`)

## Troubleshooting

### Add-in not showing in Word?

- Verify the server is running on https://localhost:3000
- Try clearing the Office add-in cache and reopening Word
- Re-sideload the add-in via Insert > Add-ins

### "Please select some text first" message?

- Make sure you have text highlighted/selected before clicking Convert
- The add-in only works on selected text

### SSL Certificate errors?

Trust the development certificate:
```bash
dotnet dev-certs https --trust
```

Then restart the server.

### Port already in use?

If port 3000 is busy, you can change it in `Properties/launchSettings.json` and update `manifest.xml` with the new URL.

## Next Steps

- Check out [EXAMPLE.md](EXAMPLE.md) for more markdown examples
- Read the full [README.md](README.md) for detailed documentation
- See [VISUAL_STUDIO.md](VISUAL_STUDIO.md) for Visual Studio-specific guidance
- Visit the [GitHub repository](https://github.com/amyxdclark/MarkdownDoctorWordAddon) for updates

## Need Help?

Open an issue on GitHub: https://github.com/amyxdclark/MarkdownDoctorWordAddon/issues
