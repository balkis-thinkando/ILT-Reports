
# PowerPoint Audit Tool

This script periodically collects info about PowerPoint files (.pptx/.pptm), including:
- Number of slides
- Number of images
- **Estimated** screenshots (heuristic)
- Text stats (word/char counts)
- Keywords (naive frequency-based)
- A short summary (first few meaningful lines detected)

It writes a CSV (and optional JSON) you can feed into dashboards or share in reports.

## Setup

1) Make sure you have Python 3.9+ installed.
2) Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

```bash
python ppt_audit.py --root "/path/to/presentations" --output "/path/to/report.csv" --json "/path/to/report.json"
```

Options:
- `--nonrecursive` : only scan the top-level directory
- `--since-days N` : include only files modified in the last N days

### Example

```bash
python ppt_audit.py --root "D:/Team Decks" --output "D:/Reports/ppt_report.csv" --json "D:/Reports/ppt_report.json" --since-days 30
```

## Scheduling

### macOS / Linux (cron)

Run every Monday at 08:00:

```
0 8 * * 1 /usr/bin/env python3 /path/to/ppt_audit.py --root "/srv/presentations" --output "/srv/reports/ppt_report.csv" --json "/srv/reports/ppt_report.json" --since-days 35 >> /srv/reports/ppt_audit.log 2>&1
```

### Windows (Task Scheduler)

1. Open **Task Scheduler** → **Create Task…**
2. **Triggers** → New… → Weekly → Monday 08:00.
3. **Actions** → New…:
   - Program/script: `python`
   - Add arguments: `"C:\\path\\to\\ppt_audit.py" --root "C:\\path\\to\\presentations" --output "C:\\path\\to\\reports\\ppt_report.csv" --json "C:\\path\\to\\reports\\ppt_report.json" --since-days 35`
   - Start in: `C:\\path\\to`

## Notes & Limitations

- Works with `.pptx` and `.pptm` (not legacy `.ppt`).
- "Screenshots" is an **estimate** based on image size/aspect ratio heuristics. Accurate detection would require computer vision and/or filename metadata.
- The short summary is intentionally simple (no external API needed). If you want more robust summaries, you can plug the extracted text into your preferred LLM pipeline.
- If your decks are on OneDrive/SharePoint/Google Drive, run the script on a machine that has the folder synced locally, or extend it to enumerate files via the respective APIs.
