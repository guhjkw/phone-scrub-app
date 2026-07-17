# Phone Scrub

A self-serve Streamlit app that cleans skip-trace Excel files by removing invalid,
disconnected, and non-working US phone numbers via the
[IPQualityScore Phone Validation API](https://www.ipqualityscore.com/documentation/phone-number-validation-api/overview).

## What it does

1. **Structural pre-filter (free)** — rejects NANP-invalid numbers (bad area code / exchange) and Pager entries before any paid API call.
2. **Live IPQS validation** — calls the IPQS API for every number that passes step 1. Deduplicates: the same number appearing across many rows/columns is validated only once.
3. **Type filter (optional)** — UI toggles to also drop VOIP and/or Landline numbers for mobile-only output.
4. **In-place rewrite** — all ~270 columns and non-phone data are preserved exactly. A Summary sheet is added to the output workbook.

## Local setup

```bash
# 1. Clone and enter the repo
git clone <your-repo-url>
cd phone-scrub-app

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add your IPQS API key
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Then edit .streamlit/secrets.toml and replace "your_key_here" with your real key

# 4. Run
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Deploy to Streamlit Community Cloud

1. Push this repo to GitHub (the `.streamlit/secrets.toml` file is gitignored, so the key stays local).
2. Go to [share.streamlit.io](https://share.streamlit.io) and click **New app**.
3. Select your repository, branch (`main`), and set the main file path to `app.py`.
4. Under **Advanced settings → Secrets**, paste:

   ```toml
   IPQS_API_KEY = "your_real_key_here"
   ```

5. Click **Deploy**. No other changes are needed — the repo runs as-is on Streamlit Cloud.

## Input file format

- `.xlsx`, single sheet, up to ~270 columns, a few hundred rows.
- Phone data lives in any column whose header contains the word **Phone** (case-insensitive). Column position and exact spelling vary — the app detects them automatically.
- Each phone cell may hold multiple numbers separated by blank lines, each optionally tagged with a type:

  ```
  (661) 600-2347 (Wireless)

  (661) 297-0544 (Landline)
  ```

- Supported type labels: `Wireless`, `Landline`, `Pager`, `OtherPhone`. Cells may also contain `N/A` or be empty.

## Output

- The same workbook with invalid/disconnected numbers removed. Surviving numbers retain their original `(xxx) xxx-xxxx (Type)` format.
- A **Summary** sheet listing counts for each removal category.
- Download button appears on screen when processing is complete.
