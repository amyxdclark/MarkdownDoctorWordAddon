# Sideloading the Add-in to Microsoft Word

This guide explains how to install ("sideload") the Markdown Doctor add-in into Microsoft Word for development and testing.

## Important: Web Add-ins vs. COM Add-ins

**Markdown Doctor is a Web Add-in (Office Add-in)**, not a COM Add-in. When looking for where to add it in Word, you need to find the option for **web-based Office Add-ins**, not "COM Add-ins" or "Word Add-ins" in the traditional sense.

- **Web Add-ins / Office Add-ins**: Modern, web-based add-ins that run in a task pane (what this project is)
- **COM Add-ins**: Legacy add-ins built with COM/VSTO technology
- **Word Add-ins (Templates)**: Document templates with macros

## Prerequisites

Before sideloading:
1. Make sure the server is running on https://localhost:3000
2. Trust the development certificate if you haven't already:
   ```bash
   dotnet dev-certs https --trust
   ```
3. Verify the server is working by visiting https://localhost:3000/taskpane.html in your browser

## Sideloading Instructions

### Option 1: Using the Insert Tab (Most Common)

1. Open **Microsoft Word**
2. Click the **Insert** tab in the ribbon
3. Look for **Add-ins** in the ribbon (you may need to look in the "Add-ins" group)
4. Click the small **dropdown arrow** next to "Add-ins" or click on "Get Add-ins"
5. In the dialog that appears, click **My Add-ins** (in the left sidebar or top tabs)
6. Click **Upload My Add-in** or look for a "Manage My Add-ins" link
7. Click **Browse** and navigate to the `manifest.xml` file in this project folder
8. Click **Upload**

### Option 2: Using the Home Tab (If Available)

Some versions of Word show Add-ins options in the Home tab:

1. Open **Microsoft Word**
2. Click the **Home** tab
3. Look for **Add-ins** button in the ribbon
4. Click it and follow the same steps as Option 1

### Option 3: Using File > Options (Word Desktop)

If you can't find the Add-ins option:

1. Go to **File** > **Options**
2. Select **Add-ins** from the left sidebar
3. At the bottom, next to "Manage:", select **COM Add-ins** dropdown
4. **Change it to "Office Add-ins"** (NOT COM Add-ins!)
5. Click **Go...**
6. In the Office Add-ins dialog, click **Upload My Add-in**
7. Browse to the `manifest.xml` file and upload it

### Option 4: Network Share Method (Windows)

For more persistent sideloading during development:

1. Create a folder to share, e.g., `C:\OfficeAddins`
2. Copy `manifest.xml` to this folder
3. Share this folder on your network or use `\\localhost\OfficeAddins`
4. In Word, go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
5. Click **Trusted Add-in Catalogs**
6. Add the network path to your shared folder as a catalog URL
7. Check the **Show in Menu** checkbox
8. Restart Word
9. Go to **Insert** > **Add-ins** > **My Add-ins** > **SHARED FOLDER** tab
10. Select Markdown Doctor and click **Add**

## Verifying the Add-in is Loaded

Once successfully sideloaded:

1. Look in the **Home** tab of Word's ribbon
2. You should see a **Markdown Doctor** group
3. Click **Show Taskpane** to open the add-in's panel

## Troubleshooting Sideloading Issues

### "Upload My Add-in" option not visible

- Make sure you're signed into your Microsoft account in Word
- Try using the File > Options > Add-ins method (Option 3 above)
- Ensure you're looking at Office/Web Add-ins, not COM Add-ins

### Add-in doesn't appear after upload

1. Check that the server is running at https://localhost:3000
2. Clear the Office add-in cache:
   - **Windows**: Delete the contents of:
     - `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
     - `%appdata%\Microsoft\Teams\manifests\`
   - **Mac**: 
     - Close Word completely
     - Delete `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
3. Restart Word completely (not just close and reopen the document)
4. Try sideloading again

### SSL/Certificate errors

If you see security warnings:

```bash
# Trust the development certificate
dotnet dev-certs https --clean
dotnet dev-certs https --trust
```

Then restart the server and try again.

### "Manifest contains errors" message

1. Verify your `manifest.xml` is valid XML
2. Check that all URLs in the manifest point to https://localhost:3000
3. Make sure there are no typos in the manifest file

### Add-in icon appears but clicking does nothing

1. Open browser developer tools (F12) in Word to check for JavaScript errors
2. Verify https://localhost:3000/taskpane.html loads correctly in a browser
3. Check that the server is still running

## Removing the Sideloaded Add-in

To remove the add-in:

1. Go to **Insert** > **Add-ins** > **My Add-ins**
2. Find **Markdown Doctor** in the list
3. Click the **...** (three dots) menu next to it
4. Select **Remove**

Or clear the Office add-in cache (see troubleshooting section above).

## Different Word Versions

The exact menu locations may vary slightly between Word versions:

- **Microsoft 365 / Office 365**: Insert > Add-ins > Get Add-ins > My Add-ins
- **Word 2019/2021**: Insert > Add-ins > My Add-ins
- **Word 2016**: Insert > My Add-ins
- **Word Online**: Insert > Office Add-ins > My Add-ins > Upload Add-in

## Need More Help?

- See the [Microsoft documentation on sideloading Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- Open an issue on [GitHub](https://github.com/amyxdclark/MarkdownDoctorWordAddon/issues)
