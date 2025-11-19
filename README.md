# United Glass Estimation Tool

A monorepo containing the backend (Python) and frontend (React + Vite) for the United Glass Estimation Calculation Tool.

## Project Structure

```
.
├── backend/          # Python backend application
│   ├── data/        # Data files and database
│   ├── systems/     # System calculation modules
│   ├── utils/       # Utility modules
│   └── main.py      # Main application entry point
├── frontend/        # React + Vite frontend application
│   ├── src/         # React source code
│   └── public/      # Static assets
└── reports/         # Generated Excel reports
```

## Backend (Python)

The backend is a Python application that handles all business logic, calculations, and data management.

### Setup

```bash
cd backend
pip install -r requirements.txt
```

### Running

```bash
python main.py
```

## Frontend (React + Vite)

The frontend is a modern React application built with Vite and managed with Bun.

### Setup

```bash
cd frontend
bun install
```

### Development

```bash
bun run dev
```

The application will be available at `http://localhost:5173`

### Build

```bash
bun run build
```

## Technology Stack

- **Backend**: Python 3.x
- **Frontend**: React 19, Vite 7, Bun
- **Package Manager**: Bun
