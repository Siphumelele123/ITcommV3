# Blazor UI - Modern Styling

This is a new Blazor Server project providing a restyled UI for the IT Accom Import Email Extract application with a modern, clean, muted color palette.

## About This Project

The **BlazorUI** project is a **styling-only** rebuild of the main form interface. It does **not** modify any existing logic, variables, or code from the original WinForms applications.

### Key Points

This project delivers a complete **modern UI redesign** with a professional muted color palette (#F5F6F8 backgrounds, #2B4C6F primary buttons, #E9EBEF secondary buttons) and **full responsive design** working perfectly on desktop (3-column grid), tablet (2-column layout), and mobile (single column). All styling is implemented in pure CSS with easy-to-customize variables, and **critically—no existing code has been touched**. The original ITAccomEmailExtractCO and ITAccomEmailExtractCOCore projects remain completely intact and functional. This UI project is styling-focused and ready for backend logic integration whenever you decide to connect it to your existing services.

### What's Included

- **Modern, responsive design**: Works seamlessly on desktop, tablet, and mobile devices
- **Muted color palette**: Professional and clean, not flashy
  - Primary: Muted slate navy (#2B4C6F)
  - Text: Near-black (#2A2E35) and medium gray (#6B7280)
  - Background: Very light cool gray (#F5F6F8) with white surfaces
  - Borders: Light gray-blue (#D8DCE3)
- **Fully accessible markup**: Semantic HTML, proper labels, ARIA-ready
- **Interactive layout**: Three-column grid (sidebars + main), collapsing to single column on mobile

### What's NOT Included

- **No business logic**: Buttons and controls are UI-only placeholders
- **No data binding**: Lists and input fields are not connected to any backend
- **No existing code changes**: The original WinForms projects remain untouched

## Running the Blazor Project

### Prerequisites
- .NET 9 SDK or later
- Visual Studio 2026, VS Code, or command line

### Option 1: Visual Studio
1. Add the **BlazorUI** project to the solution via Visual Studio
   - Right-click Solution > Add > Existing Project
   - Select `BlazorUI/BlazorUI.csproj`
2. Set BlazorUI as the startup project
3. Press F5 or click "Run"
4. Opens at `https://localhost:7XXX`

### Option 2: Command Line
```powershell
cd BlazorUI
dotnet run
```

Then open a browser to `https://localhost:7XXX`

## Project Structure

```
BlazorUI/
├── Program.cs                 # Blazor Server configuration
├── App.razor                  # Router setup
├── _Imports.razor            # Global using statements
├── Pages/
│   ├── _Host.cshtml          # Host page (HTML wrapper)
│   └── Index.razor           # Main UI layout and controls
├── Shared/
│   └── MainLayout.razor      # Layout component
├── wwwroot/
│   └── css/
│       └── app.css           # Styling (all colors, responsive design)
└── BlazorUI.csproj           # Project file
```

## Styling Details

All styling is in `wwwroot/css/app.css`:

- **Color variables** (CSS custom properties) for consistency
- **Responsive breakpoints**:
  - `768px` and below: Tablet layout (2-column)
  - `640px` and below: Mobile layout (1-column)
  - `480px` and below: Small mobile optimizations
- **Component styles**: Buttons, forms, cards, sections
- **Animations & transitions**: Smooth hover/focus states

## Next Steps

To connect this Blazor UI to your backend logic:

1. **Extract business logic** from the WinForms `FrmEmailExport.cs` into a shared service layer
2. **Create Blazor event handlers** in `Pages/Index.razor` that call those services
3. **Add data models** to pass data between UI and logic
4. **Implement Blazor components** (e.g., `<FileList>`, `<ProcessingStatus>`) for reusability

## Technology Stack

- **.NET 9** (modern, LTS)
- **Blazor Server**: Server-side rendering + real-time updates
- **CSS3**: Custom properties, Grid, Flexbox
- **HTML5**: Semantic markup

---

**Created**: Restyled UI project (styling only) | **No existing code modified**
