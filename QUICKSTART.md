# Quick Start Guide

## Getting Started in 5 Minutes

### Step 1: Install Dependencies

```bash
npm install
```

### Step 2: Generate SSL Certificates

For local development, you need SSL certificates:

```bash
npx office-addin-dev-certs install
```

### Step 3: Start the Development Server

```bash
npm run dev-server
```

Keep this terminal window open. The server will run on https://localhost:3000

### Step 4: Sideload the Add-in

Open a new terminal and run:

```bash
npm start
```

This will open Microsoft Word with your add-in loaded.

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

1. Open Word (via `npm start`)
2. Type this markdown:

```markdown
## Project Status

The project is **on track** and we have completed:

1. Design phase
2. Development
3. Testing

> Note: Final review scheduled for next week.
```

3. Select all the text you just typed
4. Click "Convert Markdown" in the task pane
5. See your markdown transformed into formatted Word text!

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

- Verify the dev server is running on https://localhost:3000
- Check that you ran `npm start` to sideload the add-in
- Try closing and reopening Word

### "Please select some text first" message?

- Make sure you have text highlighted/selected before clicking Convert
- The add-in only works on selected text

### SSL Certificate errors?

Run:
```bash
npx office-addin-dev-certs install --force
```

Then restart the dev server.

## Next Steps

- Check out [EXAMPLE.md](EXAMPLE.md) for more markdown examples
- Read the full [README.md](README.md) for detailed documentation
- Visit the [GitHub repository](https://github.com/amyxdclark/MarkdownDoctorWordAddon) for updates

## Need Help?

Open an issue on GitHub: https://github.com/amyxdclark/MarkdownDoctorWordAddon/issues
