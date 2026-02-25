# FinSight — SME Financial Analysis Tool

FinSight is an internal accountant's tool for analysing SME client financial statements, generating key ratios, benchmarking against ATO small business benchmarks, and producing professional commentary as talking points for client meetings.

**All client data stays on your machine** — AI commentary is generated via [Ollama](https://ollama.com) (a local LLM server). No data is sent to external APIs.

---

## Features

- **File Upload**: Accept Xero CSV/Excel exports or PDF financial statements
- **Smart Period Detection**: Automatically detects and sorts date columns (FY2024, 2023/24, "Year ended 30 June 2024", etc.)
- **Metrics**: 20+ financial ratios with traffic-light status (liquidity, profitability, efficiency, leverage, growth)
- **ATO Benchmarking**: Compare client metrics against ATO small business industry benchmarks (22 industries, offline fallback)
- **AI Commentary**: Local AI commentary via Ollama — data never leaves your machine
- **Red Flag Detection**: Automatic detection of 7 financial warning signs
- **Polished Exports**: PDF working paper with cover page & watermark, 7-tab Excel workbook, structured Word document
- **Demo Mode**: Built-in sample data for demonstrations

---

## Setup

### 1. Install Python dependencies

```bash
pip install -r finsight/requirements.txt
```

### 2. Install and configure Ollama (for AI commentary)

AI commentary is generated locally using [Ollama](https://ollama.com). This ensures client data never leaves your machine.

**a) Install Ollama**

Download from [https://ollama.com/download](https://ollama.com/download) and follow the installer for your OS.

**b) Pull a model** (choose one — see RAM requirements below)

```bash
ollama pull llama3.2      # Recommended — 3B params, ~4 GB RAM
ollama pull mistral       # 7B params, ~8 GB RAM
ollama pull qwen2.5       # 7B params, ~8 GB RAM
ollama pull llama3.1      # 8B params, ~8 GB RAM
```

**c) Start Ollama** (if not auto-started)

```bash
ollama serve
```

Ollama typically auto-starts on macOS. On Linux/Windows, run `ollama serve` in a terminal.

**d) Verify** — you should see a green status indicator in the FinSight sidebar once Ollama is running.

### 3. Run FinSight

```bash
cd finsight && streamlit run app.py
```

The app opens at [http://localhost:8501](http://localhost:8501).

---

## Project Structure

```
finsight/
├── app.py                    # Main Streamlit app (8-tab UI)
├── parser/
│   ├── xero_parser.py        # Xero CSV/Excel parsing with smart period detection
│   └── pdf_parser.py         # PDF parsing with pdfplumber
├── metrics/
│   └── calculator.py         # 20+ financial metric calculations
├── benchmarks/
│   ├── ato_fetcher.py        # ATO benchmark query logic
│   └── ato_benchmarks.json   # Offline fallback benchmark data (22 industries)
├── commentary/
│   └── claude_commentary.py  # Ollama local LLM commentary generation
├── exports/
│   ├── pdf_export.py         # Polished PDF with cover, watermark, sections
│   ├── excel_export.py       # 7-tab Excel workbook
│   └── word_export.py        # Structured Word document
├── utils/
│   └── formatters.py         # Shared number formatting utilities
├── assets/
│   └── logo_placeholder.png
├── requirements.txt
└── README.md
```

---

## Usage

1. Launch the app and complete the **Session Setup** in the sidebar (client name, industry, FY end, firm name)
2. Select your source type:
   - **Xero CSV/Excel**: Upload P&L and Balance Sheet exports from Xero
   - **PDF Financial Statements**: Upload a PDF (extraction is approximate — always review)
   - **Demo Mode**: Explore with built-in sample data (no file needed)
3. Click **Run Analysis** — detected period labels are shown; a warning appears if positional fallback was used
4. Navigate the 8 tabs: Overview | Profitability | Liquidity | Efficiency | Leverage | Benchmarks | Commentary | Export
5. In the **Commentary** tab, select an Ollama model and click **Generate AI Commentary**
6. Edit the commentary as needed, then **Export** to PDF, Excel, or Word

---

## Exports

| Format | Contents |
|--------|----------|
| **PDF** | Cover page with client details & INTERNAL USE ONLY watermark, Exec Summary, detailed metric sections, ATO benchmarks, AI commentary, Appendix |
| **Excel** | 7 sheets: Cover, Executive Summary, Detailed Metrics, P&L Data, Balance Sheet Data, Charts, Commentary |
| **Word** | Cover page, Exec Summary, metric sections, ATO benchmarks, AI commentary, blank Accountant's Notes page |

All exports are marked **INTERNAL USE ONLY**.

---

## Notes

- ATO benchmarks are embedded as a fallback JSON; 22 industry categories with typical ranges as % of turnover
- ATO benchmark data is updated annually and may lag by one financial year
- PDF parsing is approximate — always review and confirm parsed figures in the confirmation step
- Period detection supports: `FY2024`, `FY24`, `2023/24`, `Jul 2023 – Jun 2024`, `Year ended 30 June 2024`, `Current Year / Prior Year`
- Number formatting: `$1,234,567` (currency), `42.3%` (ratios), `1.85x` (ratios), `47 days` (days)
- Metric calculations are not modified by Ollama selection; only the commentary text is LLM-generated

---

## Parsing Accuracy — Best Practice

For best results when uploading Xero CSV/Excel or PDF financial statements:

### Include explicit subtotal rows
Financial statements should ideally include explicit **subtotal/total rows** for each section:
- `Total Revenue` or `Total Income` (not just individual revenue line items)
- `Total Cost of Sales` or `Total Direct Costs` (not just individual COGS lines)
- `Gross Profit` as an explicit row
- `Total Operating Expenses`

When these rows are present, the parser will use them directly. Without them, the parser sums individual line items under the detected section, which is less reliable.

### Inventory — always use the Balance Sheet figure
Inventory for ratio calculations (Current Ratio, Quick Ratio, Inventory Days) is sourced **exclusively from the Balance Sheet current assets section**. The "Closing Stock" line in the P&L COGS section is a cost adjustment and will not be used. If inventory is not detected on the Balance Sheet, the tool will note that Quick Ratio = Current Ratio and Inventory Days cannot be calculated.

### Note references
If your statements include note reference numbers in a column adjacent to account names, the parser will attempt to detect and exclude these columns automatically. The Data Quality panel shows which columns were identified as reference columns.

### EBIT & EBITDA
EBIT and EBITDA are always calculated from components (Net Profit + Tax + Interest + D&A) rather than relying on an explicit line in the statements. The component breakdown is shown in the Profitability tab.

### Data Integrity Checks
After parsing, the tool runs 6 automated self-checks before displaying results:
1. **P&L Balance** — Revenue − COGS − OpEx − Interest − Tax = Net Profit
2. **Gross Profit Consistency** — Parsed GP = Revenue − COGS
3. **Balance Sheet Equation** — Total Assets = Total Liabilities + Equity
4. **Equity Movement** — Closing Equity ≈ Opening Equity + Net Profit (±capital movements)
5. **Current Assets Subtotal** — Sum of identified CA components ≤ Total Current Assets
6. **Revenue Reasonableness** — Flag if revenue changed >50% YoY

FAIL results (red) must be acknowledged before viewing analysis. WARN results (amber) allow proceeding but display a persistent banner.
