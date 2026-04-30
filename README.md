# Telegram Chat Analyzer Website

This project is a Streamlit website for analyzing Telegram/support chats and manager performance.

The website can read different input formats, convert them into one standard table, analyze the conversation, and generate one combined Excel report.

## New website features

- Upload one or many files in the browser
- Supports different input formats:
  - Telegram Desktop JSON export
  - CSV table
  - Excel table
  - Plain text chat
- Input format selector: Auto detect / Telegram JSON / CSV / Excel .xlsx / Plain text
- Field mapping for CSV and Excel files
- Manager detection by role column or manager-name keywords
- Preview normalized input before analysis
- Combined report download as `.xlsx`
- Personal recommendation cards for each detected manager
- Separate CSV downloads for every sheet
- Normalized input export as CSV or JSON

## Report sheets

The Excel report contains:

1. `overview` — high-level score for every chat
2. `manager_summary` — manager performance and advice
3. `personal_recommendations` — personal coaching recommendation for each detected manager
4. `mood_flow` — mood/tone analysis by message windows
5. `topics` — detected topics and mention counts
6. `critical_cases` — risky conversations and bad service signals
7. `messages` — normalized source messages

## Input format rules

### 1. Telegram Desktop JSON

Expected structure:

```json
{
  "name": "Chat name",
  "messages": [
    {
      "id": 1,
      "date": "2026-04-21T09:00:00",
      "from": "Client Name",
      "text": "Hello"
    }
  ]
}
```

### 2. CSV or Excel .xlsx table

Recommended columns:

```text
timestamp, sender, role, text
```

The `role` column is optional. If it is missing, the website detects the manager by keywords such as `Manager`, `Sales`, or `Support`.

### 3. Plain text

Supported examples:

```text
2026-04-21 10:00 - Miras Client: Hello, I need a lamp.
2026-04-21 10:03 - Sales Manager Ruslan: Hello, how can I help?
```

or:

```text
Miras Client: Hello, I need a lamp.
Sales Manager Ruslan: Hello, how can I help?
```

## How to run the website

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Run the website:

```bash
python3 -m streamlit run app.py
```

Open the local link shown in the terminal, usually:

```text
http://localhost:8501
```

## How to use sample data

The project includes `sample_data/` with 10 synthetic Telegram JSON exports. In the website, select:

```text
Use included sample_data instead of upload
```

Then click:

```text
Analyze chats
```

## Command-line batch mode

You can still generate one combined Excel report without opening the website:

```bash
python3 batch_runner_combined.py --input_dir sample_data --output reports/combined_report.xlsx
```

For a specific format:

```bash
python3 batch_runner_combined.py --input_dir sample_data --output reports/combined_report.xlsx --input_format telegram_json
```

With manager hints:

```bash
python3 batch_runner_combined.py --input_dir sample_data --output reports/combined_report.xlsx --manager_hints "Dana,Ruslan,Manager,Support,Sales"
```

## Project structure

```text
telegram-chat-analyzer-website/
├── app.py
├── analyzer.py
├── report_exporter.py
├── batch_runner_combined.py
├── requirements.txt
├── README.md
├── .gitignore
└── sample_data/
    ├── 01_reading_lamp_good_consultation.json
    ├── 02_aggressive_manager_conflict.json
    └── ...
```
