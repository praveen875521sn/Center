# Center Location Dashboard

A Flask-based web dashboard for visualizing center data, manpower, employer, migrant, and impact information.

## Features
- 📍 Active Centers — interactive map with filters
- 📋 QP & Criteria — sector and business vertical analysis
- 📊 Impact — enrollment, certification, placement tracking
- 👥 Manpower — staff deployment map and filters
- 🏢 Employer Master — placement analytics
- 🌍 Migrant Data — migrant enrollment breakdown

## Setup

### Local Development
```bash
pip install -r requirements.txt
python app.py
```
Visit `http://localhost:5000`

### Data Files
Place these Excel files in the `data/` folder:
- `MASTER_NEW.xlsx`
- `Impact.xlsx`
- `Cluster_Master.xlsx`
- `Criteria.xlsx`
- `Employer_Master.xlsx`
- `Manpower.xlsx`
- `migrant_Data.xlsx`

## Deploy to Render
1. Push this repo to GitHub
2. Go to [render.com](https://render.com) → New → Web Service
3. Connect your GitHub repo
4. Render auto-detects `render.yaml` and deploys

## Project Structure
```
├── app.py              # Flask backend
├── templates/
│   └── index.html      # Frontend dashboard
├── data/               # Excel data files (not committed)
├── requirements.txt
├── render.yaml
└── README.md
```
