# Sideloading the Add-in to Microsoft Word

This guide explains how to install ("sideload") the Markdown Doctor add-in into Microsoft Word for development and testing.

## Important: Understanding Office Add-ins vs. Other Add-in Types

**Markdown Doctor is an Office Add-in (Web Add-in)**, not a COM Add-in or Word Add-in template. This is crucial to understand when installing:

### What Type of Add-in Is This?
- ✅ **Office Add-ins** (also called "Web Add-ins"): Modern, web-based add-ins that run in a task pane - **THIS IS WHAT MARKDOWN DOCTOR IS**
- ❌ **COM Add-ins**: Legacy add-ins built with COM/VSTO technology - **NOT THIS**
- ❌ **Word Add-ins**: Document templates with macros (.dotm files) - **NOT THIS**

### Why Does This Matter?
When you go to **File > Options > Add-ins** in Word, you'll see a "Manage:" dropdown at the bottom. This dropdown has multiple options like:
- "Word Add-ins" (for document templates)
- "COM Add-ins" (for legacy add-ins)
- **"Office Add-ins"** ← **YOU NEED TO SELECT THIS ONE**

**Common Confusion**: Many users select "Word Add-ins" thinking it means "add-ins for Word", but that's actually for document templates. You need to select **"Office Add-ins"** for this project.

## Prerequisites

Before sideloading:
1. Make sure the server is running on https://localhost:3000
2. Trust the development certificate if you haven't already:
   ```bash
   dotnet dev-certs https --trust
   ```
3. Verify the server is working by visiting https://localhost:3000/taskpane.html in your browser

## Sideloading Instructions for Microsoft 365 Desktop Word

### Method 1: Using File > Options (Most Reliable for Desktop Word)

This is the **recommended method** for Microsoft 365 desktop Word:

1. Go to **File** > **Options**
2. Select **Add-ins** from the left sidebar
3. At the bottom of the screen, next to **"Manage:"**, you'll see a dropdown menu
4. Click the dropdown menu (it may currently show "COM Add-ins" or "Word Add-ins")
5. **Select "Office Add-ins"** from the dropdown
   - **Important**: Do NOT select "Word Add-ins" or "COM Add-ins" - these are for different types of add-ins
   - You need to select **"Office Add-ins"** (sometimes also called "Web Add-ins")
6. Click the **Go...** button
7. In the Office Add-ins dialog that appears, look for **"Upload My Add-in"** or **"MY ADD-INS"** tab
8. Click **"Upload My Add-in"** (you may see a folder icon)
9. Click **Browse...** and navigate to the `manifest.xml` file in the project folder
10. Select the file and click **Upload**
11. The add-in should now appear in Word's Home tab ribbon

### Method 2: Using the Insert Tab (Alternative)

If you prefer using the ribbon:

1. Open **Microsoft Word**
2. Click the **Insert** tab in the ribbon
3. Look for **Add-ins** in the ribbon (you may need to look in the "Add-ins" group)
4. Click the small **dropdown arrow** next to "Add-ins" or click on "Get Add-ins"
5. In the dialog that appears, click **My Add-ins** (in the left sidebar or top tabs)
6. Click **Upload My Add-in** or look for a "Manage My Add-ins" link
7. Click **Browse** and navigate to the `manifest.xml` file in this project folder
8. Click **Upload**

### Method 3: Using the Home Tab (If Available)

Some versions of Word show Add-ins options in the Home tab:

1. Open **Microsoft Word**
2. Click the **Home** tab
3. Look for **Add-ins** button in the ribbon
4. Click it and follow the same steps as Method 2

### Method 4: Network Share Method (Windows - For Advanced Users)

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
