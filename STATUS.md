# Status

## Recent Changes
- `PENDING` Fix tasks.py: broken import path for get_historical_btc_price, missing get_object_or_404 import, duplicate declarations
- `58f7c28` Fix API fetch hanging: add connect+read timeout tuple (10s, 30s) and User-Agent header to all requests in api_utils.py
- `e68b1e7` Add Meta.ordering to Miner model to fix UnorderedObjectListWarning on paginated list views
- `602eb16` Split views.py into views/ package: dashboards.py, crud.py, exports.py, imports.py, tasks.py
- `606a65a` Add db_index=True to Payout.payout_date, Expense.expense_date, Expense.category, Miner.is_active
- `febde82` Migrate from xlwt/xlrd (.xls) to openpyxl (.xlsx): replace all imports/exports, update templates, update requirements.txt
- `507b498` Combine duplicate mempool.space API calls into single get_bitcoin_hashrate_and_difficulty() function
- `baab179` Add @require_POST to state-changing endpoints: toggle_miner_active, fetch_closing_price, bulk_fetch_closing_prices, trigger_fetch_api_data
- `4ad5880` Replace hasattr field validation in imports with _meta.get_fields() to prevent matching methods/properties/reverse relations
- `9a0fcd2` Replace all bare except: clauses with specific exception types and add logging to import functions and context_processors
- `9d80b70` Replace module-level mutable dicts for background task state with Django cache framework, making status shareable across processes in production
- `1470588` Fix broken import in import_expense_data: replace nonexistent xldate module with xlrd.xldate_as_datetime
- `78f7ce0` Extract shared data-gathering logic into mining/services.py, eliminating ~900 lines of duplication between dashboard views and their export counterparts
- `42deabb` Fix N+1 query problems: add select_related('platform') to all miner, payout, expense, and topup querysets that access platform fields
- `028d125` Include CLAUDE.md rules and fix missed staging
- `8c83f5f` Fix STATUS.md to follow PENDING/hash commit tracking convention
- `9459609` Update STATUS.md with commit hashes and security hardening entry
- `98bf580` Security hardening: moved SECRET_KEY and DEBUG to environment variables (.env), removed staticfiles/ from tracking, rewrote .gitignore exhaustively, scrubbed SECRET_KEY and staticfiles from entire git history using git-filter-repo
- `a49892c` Added CLAUDE.md with comprehensive project overview, tech stack, structure, and run instructions
- `b97decf` Fetching API data is now a background task
- `c877ff5` Bulk fetching of closing prices
- `f87dd86` Pagination changes for miners
- `26b831e` Fix issue with next/previous buttons
- `0b64215` Added filtering by platform on the overview dashboard
