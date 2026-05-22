# Nandy Calculations Dashboard

A web-based measurement verification tool that calculates **No × L × B × D** and compares it against a given Content value.

## Features
- Manual entry spreadsheet with live calculations
- **No** column supports two values with multiply (No1 × No2)
- Paste from Excel (multi-row tab-separated)
- Undo/Redo (Ctrl+Z / Ctrl+Y)
- Arrow key, Tab, Enter navigation
- CSV export
- Summary totals and match/mismatch indicators

## Quick Start

### Local Development
```bash
cd backend
pip install flask flask-cors pandas waitress
python app.py
```
Open http://localhost:5000

### Production (Windows Server)
```bash
cd backend
python app.py --prod --port 5000
```
Share `http://<server-ip>:5000` with your team.

## Project Structure
```
measurement_dashboard/
├── backend/
│   └── app.py          # Flask server
├── frontend/
│   └── index.html      # Single-page dashboard
├── start_server.bat    # One-click production launcher
├── requirements.txt
└── .gitignore
```
