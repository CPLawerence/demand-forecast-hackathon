# Setup Guide

Complete onboarding instructions for the demand forecast hackathon project.

## Prerequisites

- Python 3.9+ (recommend using pyenv)
- Git
- Snowflake credentials (shared service account)

## Step 1: Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/demand-forecast-hackathon.git
cd demand-forecast-hackathon
```

## Step 2: Get Snowflake Credentials

Contact the project lead for the shared Snowflake service account credentials:
- Account
- Username
- Password
- Database
- Warehouse
- Schema

## Step 3: Configure Environment Variables

```bash
cp .env.example .env
```

Edit `.env` and fill in your Snowflake credentials:
```
SNOWFLAKE_ACCT=your_account_here
SNOWFLAKE_USER=your_username_here
SNOWFLAKE_PASSWORD=your_password_here
SNOWFLAKE_DATABASE=your_database_here
SNOWFLAKE_WAREHOUSE=etl_warehouse
SNOWFLAKE_SCHEMA=your_schema_here
SNOWFLAKE_THREADS=4
```

**Important:** Never commit `.env` to version control!

## Step 4: Set Up Python Environment

### Option A: Using venv (recommended)

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
```

### Option B: Using pyenv

```bash
pyenv install 3.11.0
pyenv virtualenv 3.11.0 demand-forecast
pyenv local demand-forecast
pip install --upgrade pip
pip install -r requirements.txt
```

## Step 5: Configure DBT

Copy the DBT profile template to your home directory:

```bash
# If you don't have an existing profiles.yml
mkdir -p ~/.dbt
cp dbt/profiles.yml.example ~/.dbt/profiles.yml

# If you have an existing profiles.yml, merge the demand_forecast profile
```

Make sure your `.env` variables are loaded before running DBT:
```bash
source .env  # or use python-dotenv
```

## Step 6: Verify Setup

### Test Python & Snowflake Connection

```python
import os
from dotenv import load_dotenv
import snowflake.connector

load_dotenv()

conn = snowflake.connector.connect(
    account=os.getenv('SNOWFLAKE_ACCT'),
    user=os.getenv('SNOWFLAKE_USER'),
    password=os.getenv('SNOWFLAKE_PASSWORD'),
    database=os.getenv('SNOWFLAKE_DATABASE'),
    warehouse=os.getenv('SNOWFLAKE_WAREHOUSE'),
    schema=os.getenv('SNOWFLAKE_SCHEMA')
)

cursor = conn.cursor()
cursor.execute("SELECT CURRENT_VERSION()")
print(cursor.fetchone())
conn.close()
```

### Test DBT Connection

```bash
cd dbt
dbt debug
```

## Step 7: Start Jupyter

```bash
jupyter lab
```

## Troubleshooting

### "Environment variable not found" errors
Make sure you've created `.env` from `.env.example` and filled in all values.

### DBT connection errors
1. Verify your `.env` variables are set correctly
2. Check that `~/.dbt/profiles.yml` exists and has the `demand_forecast` profile
3. Run `dbt debug` for detailed error messages

### Snowflake authentication errors
Contact the project lead to verify your service account credentials are correct.
