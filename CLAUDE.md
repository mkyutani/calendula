# CLAUDE.md

calendula is an MCP server (Python, `mcp` SDK 2.x `MCPServer`) that reads Outlook (classic) from WSL. User-facing docs are in `README.md`; the specification is `docs/spec.md` and the test design `docs/test-design.md`.

## Layout

- `src/calendula/server.py`: the tools and the server instructions the model reads
- `src/calendula/cache.py`: the in-memory cache (today and later: TTL; past days fetched after they ended: kept)
- `src/calendula/outlook.py`: the bridge to Windows (`wscript.exe` + `run_hidden.vbs` + `get_calendar.ps1`), and `restart`
- `src/calendula/events.py`: parsing `events.json`, formatting, picking the event at a time
- `src/calendula/cli.py`: `calendula [serve|list|show]`

## Rules

- Do not fetch the user's real calendar from a development session unless asked; the events are private. Test with fakes (`Calendar(fetch=...)`, `build_server(calendar, restart=...)`, `read_result` on a temp dir)
- To check the Windows plumbing without Outlook, run `outlook._run_powershell` with a probe script that writes `events.json` itself
- Never write fetched events to disk, and never return older copies when a fetch fails
- Keep the tool names, arguments and the output shape stable: skills tell the model to use these tools
- Change behavior in this order: `docs/spec.md` (the rule and its ID), `docs/test-design.md` (the test case and its ID, the traceability table), the test (`test_<case id>_...`), then the code. A test without a case in the design, or a rule without a test, is a gap to fix
- `uv run pytest`, `uv run ruff check` and `uv run ruff format --check` must pass; CI runs them on Python 3.11-3.14
- On WSL, also run the Windows tests (`CALENDULA_WINDOWS_TESTS=1 uv run pytest tests/test_windows.py`) and, after changing a rule, the mutation check (`CALENDULA_WINDOWS_TESTS=1 uv run python scripts/mutation_check.py`); add a mutation for a new rule that is easy to get wrong
- Docs, comments and messages are in English
