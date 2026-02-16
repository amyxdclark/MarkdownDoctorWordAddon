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

**Important**: This is an **Office Add-in** (also called a "Web Add-in"), not a COM Add-in or Word Add-in template. This distinction matters when you're looking for where to install it.

#### Recommended Method for Microsoft 365 Desktop Word:

1. Open Microsoft Word
2. Go to **File** > **Options**
3. Select **Add-ins** from the left sidebar
4. At the bottom, find the **"Manage:"** dropdown menu
5. Click the dropdown and select **"Office Add-ins"**
   - **Critical**: Do NOT select "Word Add-ins" (those are document templates) or "COM Add-ins" (those are legacy add-ins)
   - You need **"Office Add-ins"** for this project
6. Click the **Go...** button
7. In the Office Add-ins dialog, click **"Upload My Add-in"**
8. Click **Browse** and navigate to the `manifest.xml` file in the project folder
9. Click **Upload**
10. The add-in will load and appear in the Home tab

#### Alternative Method - Using Insert Tab:

1. Open Microsoft Word
2. Go to **Insert** tab in the ribbon
3. Click **Add-ins** (you may need to click a dropdown arrow)
4. Click **My Add-ins** (look for this in the left sidebar or tabs)
5. Click **Upload My Add-in**
6. Browse to the `manifest.xml` file in the project folder
7. Click **Upload**
8. The add-in will load and appear in the Home tab

For detailed instructions with troubleshooting, see [SIDELOAD.md](SIDELOAD.md).

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

### Can't find where to upload the add-in?

This is an **Office Add-in** (Web Add-in). Common places to look:

**Method 1 (Recommended for Desktop Word):**
- **File** > **Options** > **Add-ins**
- At the bottom, change **"Manage:"** dropdown to **"Office Add-ins"** (NOT "Word Add-ins" or "COM Add-ins")
- Click **Go...** > **Upload My Add-in**

**Method 2 (Alternative):**
- **Insert** > **Add-ins** > **My Add-ins** > **Upload My Add-in**

**What NOT to use:**
- ❌ "COM Add-ins" (for legacy add-ins)
- ❌ "Word Add-ins" (for document templates)
- ✅ You need "Office Add-ins" for this project

See [SIDELOAD.md](SIDELOAD.md) for detailed, step-by-step instructions.

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
- See [SIDELOAD.md](SIDELOAD.md) for detailed add-in installation instructions
- Visit the [GitHub repository](https://github.com/amyxdclark/MarkdownDoctorWordAddon) for updates

## Need Help?

Open an issue on GitHub: https://github.com/amyxdclark/MarkdownDoctorWordAddon/issues
