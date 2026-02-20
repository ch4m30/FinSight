# FinSight — SME Financial Analysis Tool

FinSight is an internal accountant's tool for analysing SME client financial statements, generating key metrics, benchmarking against ATO small business benchmarks, and producing professional commentary as talking points for client meetings.

## Features

- **File Upload**: Accept Xero CSV/Excel exports or PDF financial statements
- **Metrics**: 20+ financial ratios with traffic-light status (liquidity, profitability, efficiency, leverage, growth)
- **ATO Benchmarking**: Compare client metrics against ATO small business industry benchmarks
- **AI Commentary**: Claude-powered professional commentary structured as meeting talking points
- **Red Flag Detection**: Automatic detection of financial warning signs
- **Export**: PDF report, Excel workbook, and Word document export
- **Demo Mode**: Built-in sample data for demonstrations

## Setup

1. Copy `.env.example` to `.env` and add your Anthropic API key:
   ```
   cp finsight/.env.example finsight/.env
   ```

2. Install dependencies:
   ```
   pip install -r finsight/requirements.txt
   ```

3. Run the app from the `finsight/` directory:
   ```
   cd finsight && streamlit run app.py
   ```

## Project Structure

```
finsight/
├── app.py                  # Main Streamlit app
├── parser/
│   ├── xero_parser.py      # Xero CSV/Excel parsing
│   └── pdf_parser.py       # PDF parsing with pdfplumber
├── metrics/
│   └── calculator.py       # Financial metric calculations
├── benchmarks/
│   ├── ato_fetcher.py      # ATO benchmark fetch logic
│   └── ato_benchmarks.json # Offline fallback benchmark data
├── commentary/
│   └── claude_commentary.py # Anthropic API commentary generation
├── exports/
│   ├── pdf_export.py       # PDF report generation
│   ├── excel_export.py     # Excel workbook generation
│   └── word_export.py      # Word document generation
├── assets/
│   └── logo_placeholder.png
├── .env                    # ANTHROPIC_API_KEY (gitignored)
├── requirements.txt
└── README.md
```

## Usage

1. Launch the app and complete the **Session Setup** in the sidebar (client name, industry, financial year, etc.)
2. Upload financial statements (Xero export or PDF) via the sidebar file uploader
3. If uploading a PDF, review and confirm the parsed figures in the data confirmation step
4. Navigate the tabs to explore metrics, benchmarks, and AI commentary
5. Edit the AI commentary if needed, then export your report

## Notes

- ATO benchmarks are embedded as a fallback JSON; the app will attempt to fetch current data from the ATO website
- ATO benchmark data is typically updated annually and may lag by one year
- PDF parsing is imperfect — always review parsed figures before running analysis
- All reports are marked "Prepared for internal use only — not for distribution without review"
