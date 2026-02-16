# Opening in Visual Studio 2026

This project is designed for Visual Studio 2026 using C# and ASP.NET Core.

## Requirements

- **Visual Studio 2026** (or Visual Studio 2022 version 17.8+) with the following workloads:
  - **ASP.NET and web development** workload
  - **Office/SharePoint development** workload (optional, for advanced Office tools)
  
- **.NET 8.0 SDK** or later - Download from https://dotnet.microsoft.com/download

## Opening the Project

1. Open **MarkdownDoctorWordAddon.sln** in Visual Studio 2026
2. Visual Studio will automatically restore NuGet packages
3. The Solution Explorer will show the C# project structure

## Building and Running

### Using Visual Studio

- **Run the project**: Press `F5` or click the "Start" button in Visual Studio
  - This will start the ASP.NET Core server on https://localhost:3000
  - A browser window will open to the taskpane
  
- **Build**: Build > Build Solution (or `Ctrl+Shift+B`)
  - This compiles the C# code and copies static files to the output directory

- **Publish**: Build > Publish
  - Creates deployment-ready files for hosting

### Using .NET CLI (Alternative)

You can use the .NET CLI directly:

```bash
# Restore dependencies
dotnet restore

# Build the project
dotnet build

# Run the project
dotnet run

# Run in watch mode (auto-rebuild on changes)
dotnet watch run

# Publish for production
dotnet publish -c Release
```

## Project Structure

The project uses the standard **ASP.NET Core Web** project format:

```
MarkdownDoctorWordAddon/
├── MarkdownDoctorWordAddon.sln    # Visual Studio solution file
├── MarkdownDoctorWordAddin.csproj # C# project file
├── Program.cs                     # ASP.NET Core entry point
├── Properties/
│   └── launchSettings.json        # VS launch configuration
├── appsettings.json              # Application configuration
├── appsettings.Development.json  # Development configuration
├── manifest.xml                  # Office Add-in manifest
└── wwwroot/                      # Static web assets
    ├── taskpane.html            # Main add-in UI
    ├── taskpane.css             # Styling
    ├── taskpane.js              # Client-side JavaScript
    ├── commands.html            # Command functions
    └── assets/                  # Icons (SVG)
```

## Debugging

1. Press `F5` in Visual Studio to start debugging
2. The ASP.NET Core server starts with the debugger attached
3. Use browser developer tools (F12) in Word to debug JavaScript
4. Set breakpoints in Program.cs for server-side debugging

## Sideloading the Add-in

After starting the server:

1. Open Microsoft Word
2. Go to **Insert** > **Add-ins** > **Upload My Add-in** (or **My Add-ins**)
3. Browse to `manifest.xml` in the project folder
4. Click **Upload**
5. The add-in will appear in the Home tab ribbon

## Troubleshooting

### Server won't start

1. Check if port 3000 is already in use:
   - Windows: `netstat -an | findstr :3000`
   - Change the port in `Properties/launchSettings.json` if needed

2. Ensure .NET 8.0 SDK is installed:
   ```bash
   dotnet --list-sdks
   ```

### SSL Certificate Issues

Trust the development certificate:
```bash
dotnet dev-certs https --trust
```

### Add-in doesn't load

1. Verify the server is running at https://localhost:3000
2. Check that `manifest.xml` points to the correct URLs
3. Clear the Office add-in cache:
   - Windows: Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
4. Restart Word and re-sideload the add-in

### Build errors

1. Restore packages: `dotnet restore`
2. Clean and rebuild: `dotnet clean && dotnet build`
3. Check for SDK version compatibility

## More Information

For detailed add-in documentation and usage, see the main [README.md](README.md) file.
