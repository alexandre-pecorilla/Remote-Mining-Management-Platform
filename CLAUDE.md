# PROJECT OVERVIEW

A Django web application for managing and analyzing a Bitcoin mining operation across multiple remote hosting platforms. The app tracks mining hardware (ASIC miners), records BTC payouts, logs expenses (CAPEX/OPEX) and platform top-ups, and pulls live network data (BTC price, hashrate, difficulty, block fees) from CoinMarketCap, CryptoCompare, and mempool.space APIs. Four dashboards provide analytics: an Overview of fleet and financial KPIs, an Income dashboard with monthly BTC/USD breakdowns, a CAPEX/OPEX expense dashboard, and a Forecasting dashboard that projects mining profitability using difficulty-based calculations. Data can be filtered by platform, imported/exported via Excel (.xlsx), and the UI supports dark mode.

## Tech Stack
- **Backend:** Django 5.2, Python 3.13, PostgreSQL 17, Gunicorn
- **Frontend:** Django templates, Bootstrap, Chart.js
- **Infra:** Docker Compose (web + db containers)
- **APIs:** CoinMarketCap (BTC price), CryptoCompare (historical prices), mempool.space (hashrate, difficulty, block fees)

## Key Structure
- `app/` — Django project root (mounted into container at `/app`)
  - `mining/` — main Django app (models, views, forms, templates, API utilities)
  - `remote_mining_management_platform/` — Django project settings and root URL config
  - `templates/` — HTML templates (base layout + mining app templates)
  - `static/` — CSS, JS, images
- `media/` — user-uploaded miner/platform images (mounted volume)
- `Dockerfile` — Python 3.13 + Gunicorn image
- `docker-compose.yml` — web + PostgreSQL services
- `entrypoint.sh` — waits for DB, runs migrations, collects static, starts Gunicorn

## Running
```
cp .env.example .env   # edit with your values
docker compose up -d --build
```

# CONTEXT & WORK METHODOLOGY

Consistency across this project is critical. To prevent codebase drift, strictly adhere to the following protocol.

## 1. Session Initialization

At the start of **every** new task or session, **read `STATUS.md` first.** It contains the current project state, active features, and established technical decisions. Do not proceed without reading it.

## 2. Implementation Standards

- Before writing new logic, check if a similar mechanism already exists or is documented in `STATUS.md`.
- Maintain strict naming conventions throughout the codebase.
- **No hardcoded secrets.** API keys, database credentials, and configuration values go in environment variables (loaded via `docker-compose.yml` or `.env`).
- **Keep dependencies minimal.** Do not add libraries unless clearly justified.

## 3. `STATUS.md` Protocol (CRITICAL)

`STATUS.md` is the **single source of truth** that preserves context across sessions. Without it, context is lost, decisions are forgotten, and work gets duplicated.

- **Update proactively and immediately** — after each meaningful change (feature, fix, decision, file added/removed). Do not batch updates or wait to be asked.
- You do **not** need user approval to update `STATUS.md`. It is always safe and expected.
- **Minimum update triggers:** (a) after implementing a feature or fix, (b) after making a technical decision, (c) before and after any session wrap-up.
- **When in doubt, update.** An over-documented `STATUS.md` is far better than a stale one.
- **Format:** All entries must be timestamped and log: what was done, decisions made, bugs fixed, and exact next steps for the next session.
- **Commit tracking:** Each STATUS.md entry must include a **commit ID**. Since the hash is only known after committing, write `PENDING` as the commit ID. At the **next** commit, resolve `HEAD` and replace the previous `PENDING` with the actual hash before staging. The latest entry will always show `PENDING` until the following commit.

## 4. Temporary Resources

`tmp_dir/` contains files, screenshots, and other assets needed to understand the current task. Read and use its contents as context for your work. **Empty this folder at the end of each task.**

## 5. Permissions

- **File Editing → AUTONOMOUS.** Full pre-approved clearance to create, edit, delete files and to update `STATUS.md`. No permission needed.
- **Git Commits → RESTRICTED.** Do **not** run `git commit`, `git push`, or alter git history unless explicitly commanded.
- **Docker Container Lifecycle → RESTRICTED.** Starting, stopping, deleting, and rebuilding containers (`docker compose up`, `down`, `stop`, `build`, etc.) is reserved for the user. Do **not** run these commands unless explicitly commanded.
- **Docker Exec → AUTONOMOUS.** Free to `docker compose exec` into running containers to run commands (e.g., `psql`, `python`, inspecting logs).

# LANGUAGE RULES

- **All code, comments, documentation, commit messages, and `STATUS.md` entries → English.**
