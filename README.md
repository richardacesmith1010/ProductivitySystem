# Personal Productivity System

Welcome! This guide turns the repository tour into a hands-on onboarding so you can start using the system right away.

## Start Here (10-minute setup)
1. **Install dependencies**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
   pip install openpyxl
   ```
2. **Grab the workbook** – open `ProductivitySystem.xlsx` in Excel (ensure macros are enabled if you plan to add VBA later).
3. **Make a personal copy** – save a version with today’s date so you can experiment freely.

> ✅ Tip: Close Excel before running any Python automation so `openpyxl` can edit the file.

## Quickstart Checklist
Work through these boxes as you explore:
- [ ] Enter three tasks on the **Tasks** sheet (include due dates and at least one recurring task).
- [ ] Add a completion entry to see it appear in the **Logs** sheet.
- [ ] Update `config.json` with a new `last_updated` date to record your session.
- [ ] Run the automation (`python task_manager.py --file ProductivitySystem.xlsx`) and confirm:
  - [ ] Completed tasks moved to **Logs**.
  - [ ] Recurring items were rescheduled with status reset to “Pending”.

## Guided Tour & Mini-Exercises
| Area | Explore | Mini-Exercise |
| --- | --- | --- |
| Excel Interface | Review how the **Tasks** and **Logs** sheets are structured. | Add conditional formatting to highlight overdue tasks using Excel rules. |
| Automation Script (`task_manager.py`) | Skim the main loop that processes each row. | Change the recurrence logic for weekly tasks (e.g., to skip weekends) and test the impact. |
| Design Guide (`guide.md`) | Note the UX principles for consistent layouts. | Draft a new sheet mockup following the spacing & typography rules. |
| Versioning (`config.json`) | Observe version metadata fields. | Increment the version after you successfully run the automation. |

## Run Your First Automation Cycle
1. Open the workbook, mark one task as **Done**, and set a recurrence (Daily/Weekly/Monthly).
2. Close Excel.
3. Execute:
   ```bash
   python task_manager.py --file ProductivitySystem.xlsx
   ```
4. Re-open the workbook. Check that the completed task now appears in **Logs** and a fresh copy exists in **Tasks** with a new due date.
5. Capture a short journal entry in **Logs** about what you automated—this builds a history for later review.

## Track Your Progress
Use this tracker to see where you are:
| Stage | Definition | Status |
| --- | --- | --- |
| 🧭 Orientation | You understand the workbook layout and key scripts. | ☐ In progress ☐ Complete |
| ⚙️ Automation | You can run `task_manager.py` without errors. | ☐ In progress ☐ Complete |
| 🎨 Customization | You’ve applied at least one UX improvement from `guide.md`. | ☐ In progress ☐ Complete |
| 📈 Iteration | You’ve noted enhancements to build next. | ☐ In progress ☐ Complete |

## What to Learn Next
- **Excel automation basics:** Deepen knowledge of tables, data validation, and conditional formatting to evolve the interface.
- **Python + openpyxl:** Extend `task_manager.py` with features like summary reports or email reminders.
- **VBA or Office Scripts:** Add buttons inside Excel that trigger scripts so the workbook feels like an app.
- **Collaboration habits:** Update `config.json`, document changes in `guide.md`, and consider introducing automated tests for workbook integrity.

Keep iterating—each session you complete above adds to your personal productivity playbook!
