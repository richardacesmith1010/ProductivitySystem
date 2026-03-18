# CLAUDE.md

## Project Overview

Personal Productivity System — a lightweight task management platform built on Excel 2016+, with Python automation for recurring task processing. Designed for single-user personal workflows.

## Repository Structure

```
ProductivitySystem/
├── ProductivitySystem.xlsx   # Main Excel workbook (Tasks + Logs sheets)
├── task_manager.py           # Python script for recurring task automation
├── config.json               # Version tracking (semantic versioning)
├── README.md                 # Project overview
├── guide.md                  # UI/UX design guidelines for the Excel interface
└── CLAUDE.md                 # This file
```

This is a flat repository with no subdirectories. All source files live at the root.

## Tech Stack

- **Python 3.x** with `openpyxl` for Excel file manipulation
- **Excel 2016+** (.xlsx) as the primary user interface
- **VBA** mentioned for future Excel automation (not yet implemented)

## Key Files

### task_manager.py
Core automation script. Processes the `ProductivitySystem.xlsx` workbook:
1. Reads tasks from the **Tasks** sheet (columns: Task ID, Task Name, Due Date, Status, Recurrence, Notes)
2. Finds tasks with Status = "Done"
3. Archives completed tasks to the **Logs** sheet
4. Creates new recurring tasks (Daily/Weekly/Monthly) with updated due dates
5. Deletes processed rows from Tasks

**Run with:**
```bash
python task_manager.py --file /path/to/ProductivitySystem.xlsx
```

### ProductivitySystem.xlsx
Excel workbook with at least two sheets:
- **Tasks** — active task list (6 columns: Task ID, Task Name, Due Date, Status, Recurrence, Notes)
- **Logs** — archive of completed tasks (5 columns: Task ID, Task Name, Completion Date, Original Due Date, Notes)

### config.json
Tracks the system version. Update `version` and `last_updated` when making releases.

## Dependencies

- `openpyxl` (external) — Excel read/write
- `argparse`, `datetime` (stdlib)

No requirements.txt exists. Install manually: `pip install openpyxl`

## Development Notes

- **No test suite** — there are no automated tests. Validate changes manually against the workbook.
- **No linting/formatting config** — no pylint, black, flake8, or pyproject.toml configured.
- **No CI/CD** — no GitHub Actions or other pipelines.
- **No .gitignore** — all files are tracked.

## Conventions

- Python code uses type hints for function signatures.
- Module-level docstrings document usage and expected Excel column layout.
- Excel design should follow the guidelines in `guide.md` (modern spacing, floating labels, conditional formatting for status).
- Monthly recurrence is approximated as 30 days.
- Task IDs are auto-incrementing integers.
- Rows are deleted bottom-to-top after processing to avoid index shifting issues.

## Common Tasks

| Task | Command |
|------|---------|
| Process recurring tasks | `python task_manager.py --file ProductivitySystem.xlsx` |
| Install dependencies | `pip install openpyxl` |
