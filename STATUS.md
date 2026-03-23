# Status

## Recent Changes
- `98bf580` Security hardening: moved SECRET_KEY and DEBUG to environment variables (.env), removed staticfiles/ from tracking, rewrote .gitignore exhaustively, scrubbed SECRET_KEY and staticfiles from entire git history using git-filter-repo
- `a49892c` Added CLAUDE.md with comprehensive project overview, tech stack, structure, and run instructions
- `b97decf` Fetching API data is now a background task
- `c877ff5` Bulk fetching of closing prices
- `f87dd86` Pagination changes for miners
- `26b831e` Fix issue with next/previous buttons
- `0b64215` Added filtering by platform on the overview dashboard
