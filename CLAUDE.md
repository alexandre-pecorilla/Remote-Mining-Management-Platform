# PROJECT OVERVIEW

A Django web application for managing and analyzing a Bitcoin mining operation across multiple remote hosting platforms. The app tracks mining hardware (ASIC miners), records BTC payouts, logs expenses (CAPEX/OPEX) and platform top-ups, and pulls live network data (BTC price, hashrate, difficulty, block fees) from CoinMarketCap, CryptoCompare, and mempool.space APIs. Four dashboards provide analytics: an Overview of fleet and financial KPIs, an Income dashboard with monthly BTC/USD breakdowns, a CAPEX/OPEX expense dashboard, and a Forecasting dashboard that projects mining profitability using difficulty-based calculations. Data can be filtered by platform, imported/exported via Excel (.xls), and the UI supports dark mode.

## Tech Stack
- **Backend:** Django 5.2, Python 3.13, SQLite
- **Frontend:** Django templates, Bootstrap, Chart.js
- **APIs:** CoinMarketCap (BTC price), CryptoCompare (historical prices), mempool.space (hashrate, difficulty, block fees)

## Key Structure
- `mining/` — main Django app (models, views, forms, templates, API utilities)
- `remote_mining_management_platform/` — Django project settings and root URL config
- `templates/` — HTML templates (base layout + mining app templates)
- `static/` — CSS, JS, images
- `media/` — user-uploaded miner/platform images

## Running
```
python manage.py runserver
```
