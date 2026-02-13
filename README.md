# email_tool
# PST Email Searcher — Legal eDiscovery Tool

A Python tool built to search through Outlook .pst email archives for legally relevant correspondence. Designed specifically for a federal healthcare fraud defense case, but adaptable to any case requiring targeted email discovery.

## What It Does

This tool takes a massive .pst file (the kind that chokes Outlook) and efficiently:

1. **Extracts** all emails without needing Outlook installed
2. **Filters** emails by a list of specific companies/contacts you provide
3. **Scores** each email for relevance using keyword matching across categories:
   - Doctor authorization language (highest weight)
   - Termination/firing language (highest weight)
   - Compliance issue language (medium weight)
   - Marketing relationship language (supporting context)
4. **Optionally** uses Claude AI to read and analyze emails for deeper relevance scoring
5. **Exports** results to Excel with multiple sheets (summary, high-priority, full content, statistics)

## Quick Start

### 1. Install Python (if you don't have it)

Download from [python.org](https://www.python.org/downloads/) — version 3.8 or higher.

During installation, **check the box that says "Add Python to PATH"**. This is important.

### 2. Download This Tool

Save all the files from this folder to a location on your computer. For example:
```
C:\Users\YourName\Documents\pst_search_tool\
```

### 3. Install Dependencies

Open a terminal (Command Prompt on Windows, Terminal on Mac):

```bash
cd C:\Users\YourName\Documents\pst_search_tool
pip install -r requirements.txt
```

**Note on pypff:** This library can sometimes be tricky to install on Windows. If you get errors:

- **Windows Option A:** Try `pip install pypff` by itself first
- **Windows Option B:** If that fails, install via conda: `conda install -c conda-forge pypff`
- **Windows Option C:** If all else fails, use the mbox conversion method (see "Alternative: mbox Method" below)
- **Mac/Linux:** Usually installs fine with pip

### 4. Prepare Your Company List

Edit the `companies.txt` file and replace the example entries with the actual marketing companies from Jane's list. One company per line:

```
ABC Marketing Group
XYZ Health Solutions LLC
Premier Medical Marketing
@healthleads.com
```

You can use company names, email domains, or full email addresses. The tool is smart about matching — it strips "LLC", "Inc", etc. and checks multiple variations.

### 5. Run the Search

```bash
python pst_email_searcher.py --pst path/to/client_emails.pst --companies companies.txt
```

That's it. The tool will:
- Parse the PST file (may take a few minutes for large files)
- Search every email against your company list and keywords
- Export results to the `./results/` folder

### 6. Review Results

Open `./results/email_search_results.xlsx`:

| Sheet | What's In It |
|-------|-------------|
| **Results_Summary** | All matching emails with scores, matched companies, and keywords |
| **High_Priority** | Only the highest-scoring emails — review these first |
| **Full_Content** | Full email bodies for detailed review |
| **Statistics** | Summary stats about the search results |

Also check `./results/search_report.txt` for a human-readable overview.

## Advanced Usage

### AI-Powered Analysis (Optional but Powerful)

The tool can optionally use Claude (Anthropic's AI) to actually **read** each email and assess its relevance — not just keyword matching, but understanding context. This is essentially what expensive eDiscovery platforms charge thousands for.

To enable this:

1. Get an API key from [console.anthropic.com](https://console.anthropic.com/)
2. Set your API key:
   ```bash
   # Windows
   set ANTHROPIC_API_KEY=your-key-here
   
   # Mac/Linux
   export ANTHROPIC_API_KEY=your-key-here
   ```
3. Run with the `--ai-score` flag:
   ```bash
   python pst_email_searcher.py --pst client_emails.pst --companies companies.txt --ai-score
   ```

**Cost note:** Claude API pricing is very reasonable. For ~1000 emails, you're looking at maybe $2-5 total depending on email length. Way cheaper than any eDiscovery platform.

### Using mbox Format (Alternative to PST)

If you can't get pypff installed, you can convert the PST to mbox format first:

**On Mac/Linux:**
```bash
# Install readpst
# Mac: brew install libpst
# Ubuntu/Debian: sudo apt install pst-utils

# Convert PST to mbox
readpst -o ./converted_emails/ -r client_emails.pst

# Then run the search on the converted files
python pst_email_searcher.py --mbox ./converted_emails/ --companies companies.txt
```

**On Windows:**
- Download a PST to mbox converter (free options available online)
- Convert the file
- Run with `--mbox` flag pointing to the output

### All Command-Line Options

```
python pst_email_searcher.py --help

Options:
  --pst PATH           Path to the .pst file
  --mbox PATH          Path to mbox file or directory (alternative to --pst)
  --companies PATH     Path to company list file (required)
  --output DIR         Output directory (default: ./results)
  --ai-score           Enable AI relevance scoring
  --api-key KEY        Anthropic API key (or use ANTHROPIC_API_KEY env var)
  --format FORMAT      Output format: excel, csv, or both (default: both)
  --max-results N      Maximum number of results to export
```

## Customizing Keywords

The keyword lists are defined at the top of `pst_email_searcher.py` in the `KEYWORD_CATEGORIES` dictionary. You can modify these for different case types:

```python
KEYWORD_CATEGORIES = {
    "doctor_authorization": {
        "weight": 3,  # how important this category is (1-3)
        "terms": [
            "doctor authorization",
            "physician order",
            # add your own terms here...
        ]
    },
    # add new categories as needed...
}
```

## Troubleshooting

### "pypff not found" or installation errors
→ Use the mbox conversion method instead (see above). It works exactly the same, just requires an extra conversion step.

### PST file won't open / corruption errors
→ Try running Outlook's built-in repair tool (`scanpst.exe`) on the PST file first. On Windows, it's usually at:
- `C:\Program Files\Microsoft Office\root\OfficeXX\SCANPST.EXE`
- `C:\Program Files (x86)\Microsoft Office\root\OfficeXX\SCANPST.EXE`

### No results found
→ Double-check your `companies.txt` file for typos. Try adding email domains in addition to company names. The tool matches flexibly, but it needs something to match against.

### Memory errors with very large PST files (5GB+)
→ Convert to mbox first (splits into smaller files), then process with `--mbox`.

### Excel file won't open
→ Make sure you have `openpyxl` installed (`pip install openpyxl`). Or use `--format csv` for a simpler output.

## Security & Confidentiality Notes

- **This tool runs entirely on your local machine.** No email data is sent anywhere unless you explicitly enable the `--ai-score` option.
- **If using AI scoring:** Email content is sent to Anthropic's API for analysis. Review Anthropic's data retention policies if this is a concern for privileged communications.
- **The company list and search results should be treated as work product** and handled according to applicable privilege rules.

## File Structure

```
pst_search_tool/
├── pst_email_searcher.py    # Main tool
├── companies.txt            # Your list of target companies (edit this)
├── requirements.txt         # Python dependencies
├── README.md                # This file
└── results/                 # Created when you run the tool
    ├── email_search_results.xlsx
    ├── email_search_results.csv
    └── search_report.txt
```
