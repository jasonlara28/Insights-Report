# 📊 Category Insights Dashboard → PowerPoint Exporter

A Streamlit web app that analyzes category sales data, displays interactive Plotly charts, auto-generates insights, and exports a polished PowerPoint deck with one click.

## Features

- **CSV Upload or Sample Data** — Upload your own category data or explore with built-in sample data (Baking Chips, Cookie Dough, Evaporated Milk, Pumpkin Puree).
- **Interactive Plotly Charts** — Grouped bar chart, line trend, and donut chart with clean, presentation-ready styling.
- **Auto-Generated Insights** — Key metrics and bullet-point takeaways derived directly from the data.
- **One-Click PowerPoint Export** — Uses Kaleido to render high-res chart PNGs and python-pptx to assemble a widescreen 16:9 `.pptx` deck — no manual chart-building required.

## Expected CSV Format

| Column | Description |
|---|---|
| `Category` | Product category name |
| `Month` | Three-letter month abbreviation (Jan–Dec) |
| `Sales` | Revenue in dollars |
| `Units` | Units sold |
| `YoY_Change` | Year-over-year percent change |

## Local Setup

```bash
# Clone the repo
git clone https://github.com/<your-username>/category-insights-dashboard.git
cd category-insights-dashboard

# Create a virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate        # macOS/Linux
venv\Scripts\activate           # Windows

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run pptx_demo_app.py
```

The app will open in your browser at `http://localhost:8501`.

## Deployment (Streamlit Community Cloud)

This app is configured for one-click deployment on [Streamlit Community Cloud](https://streamlit.io/cloud):

1. Push this repo to a **public** GitHub repository.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **"Create app"** → select your repo, branch (`main`), and set the main file to `pptx_demo_app.py`.
4. Click **Deploy**. Your app will be live in minutes.

## Tech Stack

| Library | Purpose |
|---|---|
| [Streamlit](https://streamlit.io) | Web app framework |
| [Plotly](https://plotly.com/python/) | Interactive charts |
| [Pandas](https://pandas.pydata.org) | Data manipulation |
| [python-pptx](https://python-pptx.readthedocs.io) | PowerPoint generation |
| [Kaleido](https://github.com/plotly/Kaleido) | Static chart image export |

## License

MIT
