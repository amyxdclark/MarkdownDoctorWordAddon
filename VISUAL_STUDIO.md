# Opening in Visual Studio

This project can now be opened in Visual Studio 2022 (or later) using the modern JavaScript/Node.js project system.

## Requirements

- **Visual Studio 2022** (version 17.0 or later) with the following workloads:
  - **Office/SharePoint development** workload (for Office Add-in tools)
  - **Node.js development** workload (for JavaScript/Node.js support)
  
- **Node.js** (version 14 or later) - Download from https://nodejs.org/

## Opening the Project

1. Open **MarkdownDoctorWordAddon.sln** in Visual Studio 2022
2. Visual Studio will automatically detect the Node.js project structure
3. The Solution Explorer will show all project files

## Building and Running

### Using Visual Studio

- **Run the project**: Press `F5` or click the "Start" button in Visual Studio
  - This will execute `npm start` which sideloads the add-in into Word
  
- **Build**: Build > Build Solution (or `Ctrl+Shift+B`)
  - This will execute `npm run build` which creates production files in the `dist/` folder

### Using npm (Alternative)

You can still use npm commands directly:

```bash
# Install dependencies (first time)
npm install

# Start development server
npm run dev-server

# Sideload add-in to Word (in separate terminal)
npm start

# Build for production
npm run build

# Validate manifest
npm run validate
```

## Project Structure

The project uses the modern **ESPROJ** format (JavaScript Project) introduced in Visual Studio 2022:

- **MarkdownDoctorWordAddon.sln** - Solution file
- **MarkdownDoctorWordAddon.esproj** - Project file (JavaScript/Node.js project)
- **package.json** - npm package configuration
- **webpack.config.js** - Webpack bundler configuration
- **manifest.xml** - Office Add-in manifest

## Debugging

1. Start the development server: `npm run dev-server`
2. Press `F5` in Visual Studio to launch Word with the add-in
3. Use browser developer tools (F12) within Word to debug JavaScript

## Troubleshooting

### "Cannot find SDK" error
- Make sure Visual Studio 2022 is updated to the latest version
- Install the "Office/SharePoint development" workload from Visual Studio Installer

### npm commands not found
- Install Node.js from https://nodejs.org/
- Restart Visual Studio after installing Node.js

### Add-in doesn't load
1. Run `npm install` to install dependencies
2. Run `npx office-addin-dev-certs install` to install SSL certificates
3. Restart Visual Studio and try again

## More Information

For detailed add-in documentation, see the main [README.md](README.md) file.
