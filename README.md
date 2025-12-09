# Demand Forecast Hackathon

Automated demand forecasting and inventory planning tool using Python and DBT with Snowflake.

## Quick Start

1. Clone this repository
2. Copy `.env.example` to `.env` and fill in your Snowflake credentials
3. Set up Python environment:
   ```bash
   python -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
4. See [SETUP.md](SETUP.md) for detailed onboarding instructions

## Project Structure

```
demand-forecast-hackathon/
├── notebooks/          # Jupyter notebooks for analysis
├── scripts/            # Python scripts
├── dbt/                # DBT project for SQL transformations
│   ├── models/         # DBT models
│   ├── macros/         # DBT macros
│   └── seeds/          # Seed data files
├── data/               # Local data files (gitignored)
├── requirements.txt    # Python dependencies
└── .env.example        # Environment variable template
```

## Team

- 5 team members collaborating on demand forecasting

## Resources

- [SETUP.md](SETUP.md) - Detailed setup instructions
- [DBT Documentation](https://docs.getdbt.com/)
- [Prophet Documentation](https://facebook.github.io/prophet/)
