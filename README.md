# eDrawings ActiveX Viewer

A clean, focused application for integrating SolidWorks eDrawings ActiveX control with modern .NET 9.0 Blazor applications.

## Features

- **Direct ActiveX Integration**: Native Windows Forms hosting of eDrawings ActiveX control
- **File Support**: View eDrawings files (.edrw, .eprt, .easm, .edwg) and SOLIDWORKS files (.sldprt, .sldasm, .slddrw)
- **Interactive Controls**: 
  - File browsing and loading
  - Measurement tools
  - 3D view controls (Top, Front, Isometric)
  - Zoom to fit functionality
  - Toolbar visibility controls
- **Clean Interface**: Modern Blazor web interface with single-click launcher

## System Requirements

- Windows Operating System
- eDrawings Viewer installed
- .NET 9.0 Runtime
- ActiveX controls enabled
- Windows Forms support
- COM interop enabled

## Project Structure

```
EdrawingsWebViewer/
├── Components/
│   ├── Layout/           # Navigation and layout components
│   └── Pages/
│       ├── Home.razor    # Main launcher page
│       └── Error.razor   # Error handling
├── WindowsForms/
│   └── DirectEDrawingsForm.cs  # Main eDrawings ActiveX host
└── wwwroot/
    ├── js/
    │   └── edrawings.js  # JavaScript utilities
    └── app.css           # Application styles
```

## Usage

1. Run the application
2. Click "Launch eDrawings Viewer" on the home page
3. Use "Browse..." to select eDrawings or SOLIDWORKS files
4. Click "Load File" to view the 3D model
5. Use the toolbar controls for measurement, views, and zoom
6. Click "Show Toolbar" to enable the native eDrawings toolbar

## Technical Details

- **Target Framework**: net9.0-windows
- **ActiveX CLSID**: 22945A69-1191-4DCF-9E6F-409BDE94D101
- **Threading**: STA thread for Windows Forms hosting
- **COM Interop**: Reflection-based method invocation for flexible API access

## Notes

- ActiveX controls only work in Windows environments
- eDrawings Viewer must be installed and registered
- The application automatically detects and tests eDrawings availability
- Multiple fallback methods ensure compatibility across different eDrawings versions 