# Mining Manager

A self-hosted Django web application for managing and analyzing a Bitcoin mining operation across multiple remote hosting platforms. Track your ASIC miners, monitor BTC payouts, log expenses, and forecast profitability — all from a single dashboard.

## Features

### Dashboards

- **Overview Dashboard** — Fleet-wide KPIs: total hashrate, power consumption, efficiency, energy costs, revenue, and ROI. Filterable by platform.
- **Income Dashboard** — Monthly BTC and USD income breakdowns with charts, grouped by platform.
- **CAPEX/OPEX Dashboard** — Expense analysis with monthly CAPEX vs OPEX breakdowns and platform-level cost tracking.
- **Forecasting Dashboard** — Projects mining profitability using live network difficulty, hashrate, block rewards, and fees.

### Data Management

- **Platforms** — Manage hosting providers with contact details, energy pricing, and portal links.
- **Miners** — Track ASIC hardware: model, hashrate, power, efficiency, purchase price, serial numbers, and active/inactive status.
- **Payouts** — Record BTC mining payouts with automatic USD valuation via historical closing prices.
- **Expenses** — Log CAPEX and OPEX with categories, amounts, invoice/receipt links, and notes.
- **Top-Ups** — Track platform account deposits and funding.

### Integrations

- **CoinMarketCap API** — Live BTC/USD price.
- **CryptoCompare API** — Historical BTC closing prices for payout valuation.
- **mempool.space API** — Network hashrate, difficulty, and 24h average block fees.

### Other

- Excel import/export for all data models (`.xlsx`)
- Dark mode
- Optional password protection
- Background API fetching with progress tracking
- Responsive Bootstrap UI

## Tech Stack

| Component  | Technology               |
| ---------- | ------------------------ |
| Backend    | Django 5.2, Python 3.13  |
| Database   | PostgreSQL 17            |
| Server     | Gunicorn                 |
| Frontend   | Django Templates, Bootstrap 5, Chart.js |
| Infra      | Docker, Docker Compose   |

## Prerequisites

- [Docker](https://docs.docker.com/get-docker/) and [Docker Compose](https://docs.docker.com/compose/install/)
- A [CoinMarketCap API key](https://coinmarketcap.com/api/) (free tier works)

### Getting a CoinMarketCap API Key

1. Go to [https://coinmarketcap.com/api/](https://coinmarketcap.com/api/)
2. Click **"Get Your Free API Key"**
3. Create an account and verify your email
4. Copy the API key from your dashboard
5. Paste it into your `.env` file as `COINMARKETCAP_API_KEY`

## Installation

### 1. Clone the repository

```bash
git clone <your-repo-url>
cd mining
```

### 2. Configure environment variables

```bash
cp .env.example .env
```

Edit `.env` with your values:

```env
DJANGO_SECRET_KEY=your-random-secret-key
DJANGO_DEBUG=True
DJANGO_ALLOWED_HOSTS=localhost,127.0.0.1
COINMARKETCAP_API_KEY=your-cmc-api-key
APP_PASSWORD=

# PostgreSQL
POSTGRES_DB=mining
POSTGRES_USER=mining
POSTGRES_PASSWORD=choose-a-strong-password
POSTGRES_HOST=db
POSTGRES_PORT=5432

# Gunicorn
GUNICORN_WORKERS=2

# Host port mapping
WEB_PORT=8000
```

| Variable                | Description                                                                 |
| ----------------------- | --------------------------------------------------------------------------- |
| `DJANGO_SECRET_KEY`     | Django secret key. Generate one with `python3 -c "import secrets; print(secrets.token_urlsafe(50))"` |
| `DJANGO_DEBUG`          | Set to `False` in production                                                |
| `DJANGO_ALLOWED_HOSTS`  | Comma-separated list of allowed hostnames                                   |
| `COINMARKETCAP_API_KEY` | Your CoinMarketCap API key                                                  |
| `APP_PASSWORD`          | Set a password to protect the app. Leave empty to disable                   |
| `POSTGRES_PASSWORD`     | PostgreSQL password. Change from default                                    |
| `GUNICORN_WORKERS`      | Number of Gunicorn workers. `2` is fine for most systems                    |
| `WEB_PORT`              | Port the app listens on                                                     |

### 3. Build and start

```bash
docker compose up -d --build
```

The app will automatically:
- Wait for PostgreSQL to be ready
- Run database migrations
- Collect static files
- Start the Gunicorn server

Open [http://localhost:8000](http://localhost:8000) in your browser.

### 4. Import your data

If you have data to import, go to each section (Platforms, Miners, Payouts, Expenses, Top-Ups) and use the **Import** button. Download the Excel template first to see the expected format.

**Import order matters:** Platforms first, then Miners, Payouts, Expenses, and Top-Ups.

## Docker Commands

### Start / Stop / Restart

```bash
# Start in background
docker compose up -d

# Stop (keeps data)
docker compose stop

# Restart
docker compose restart

# Stop and remove containers (keeps data volumes)
docker compose down
```

### Build / Rebuild

```bash
# Build and start
docker compose up -d --build

# Force full rebuild (no cache)
docker compose build --no-cache && docker compose up -d
```

### Logs

```bash
# All logs
docker compose logs

# Follow logs in real-time
docker compose logs -f

# Web container only
docker compose logs -f web

# Database container only
docker compose logs -f db
```

### Database

```bash
# Open a Django shell
docker compose exec web python manage.py shell

# Open a PostgreSQL shell
docker compose exec db psql -U mining -d mining

# Run migrations manually
docker compose exec web python manage.py migrate

# Create a Django superuser (for /admin)
docker compose exec web python manage.py createsuperuser
```

### Updating After a Git Pull

```bash
git pull
docker compose up -d --build
```

Migrations and static file collection run automatically on container start.

### Full Reset

```bash
# Stop containers and delete all data (database, static files)
docker compose down -v

# Rebuild from scratch
docker compose up -d --build
```

## Project Structure

```
mining/
├── docker-compose.yml          # Service definitions (web + db)
├── Dockerfile                  # Python 3.13 + Gunicorn image
├── entrypoint.sh               # Startup script (migrate, collectstatic, run)
├── .env                        # Environment variables (not in repo)
├── .env.example                # Template for .env
├── CLAUDE.md                   # AI assistant instructions
├── STATUS.md                   # Project changelog
├── app/                        # Django project (mounted into container)
│   ├── manage.py
│   ├── requirements.txt
│   ├── mining/                 # Main Django app
│   │   ├── models.py           # Data models
│   │   ├── views/              # View modules (dashboards, CRUD, exports, imports, tasks)
│   │   ├── services.py         # Shared dashboard data logic
│   │   ├── api_utils.py        # External API integrations
│   │   ├── forms.py            # Django forms
│   │   ├── middleware.py        # Password protection
│   │   ├── urls.py             # URL routing
│   │   └── migrations/         # Database migrations
│   ├── remote_mining_management_platform/
│   │   ├── settings.py         # Django settings
│   │   ├── urls.py             # Root URL config
│   │   └── wsgi.py             # WSGI entry point
│   ├── templates/              # HTML templates
│   └── static/                 # CSS, JS, images
└── media/                      # User uploads (platform logos, miner images)
```

## License

This project is for personal use. No license has been specified.
