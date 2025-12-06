# United Glass Estimation Tool - React Version

This is the React version of the United Glass Estimation Tool, converted from the original Flet (Python) application.

## Features

- ✅ Project management (create, delete, select projects)
- ✅ Elevation management (create, update, delete elevations)
- ✅ Door management (add, edit, delete doors)
- ✅ Dynamic bay width/height inputs for YES 45TU FRONT SET system
- ✅ Excel report generation (using ExcelJS)
- ✅ Material tracking and waste optimization
- ✅ All original functionality preserved

## Installation

1. Install dependencies:
```bash
npm install
```

2. Start the development server:
```bash
npm run dev
```

3. Build for production:
```bash
npm run build
```

## Project Structure

```
src/
├── data/              # Data files (parts_data.json, partNumber.ts)
├── systems/           # System-specific calculations
├── utils/             # Utility functions (formulas, pricing, storage)
├── views/             # React views (ProjectsView, WorkspaceView)
├── App.tsx            # Main app component with routing
└── main.tsx           # Entry point
```

## Data Storage

The application uses browser localStorage to store:
- Project list
- Elevation data
- Extra materials inventory
- Door configurations

## Key Differences from Flet Version

1. **Frontend Framework**: React instead of Flet
2. **Storage**: localStorage instead of file system
3. **Excel Generation**: ExcelJS library instead of openpyxl
4. **Language**: TypeScript/JavaScript instead of Python

## Converting Parts Data

The `parts_data.py` file has been converted to `src/data/parts_data.json` using the `convert_parts_data.py` script.

## Notes

- The application maintains all original functionality
- File operations are handled through browser APIs (localStorage, file downloads)
- Excel reports are generated client-side and downloaded
- All calculations and business logic have been preserved

## Development

The project uses:
- React 18
- TypeScript
- Vite for building
- React Router for navigation
- ExcelJS for Excel generation

