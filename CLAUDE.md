# CLAUDE.md — Project Context

## What this project is
A self-serve Streamlit web app that cleans skip-trace Excel files by removing
invalid, disconnected, and non-working US phone numbers. Built for a freelance
client deliverable under the brand Pathfinder Automation Solutions.

## The client's requirement (source of truth)
"Identify numbers that are disconnected, invalid, or no longer in service.
Preserve valid and active numbers. Keep all other columns and data unchanged."

## Non-negotiables (do not simplify these away)
- Detect phone columns dynamically by header containing "Phone" — never hardcode
  column positions. Files have ~270 columns and the layout varies.
- Preserve the full 270-column layout byte-for-byte. Edit phone cells in place
  with openpyxl. Do NOT round-trip through a DataFrame.
- Deduplicate API lookups: validate each unique number once, cache, then apply.
- Removal rule: remove if valid==false OR active==false. Never remove on
  success==false (that is an API error / exhausted credits — keep and flag).
- If a lookup fails or times out, keep the number and flag it "unverified".
  Never silently delete a number because an API call failed.
- The IPQS API key is read from st.secrets["IPQS_API_KEY"]. It is NEVER
  hardcoded and NEVER committed. .gitignore must exclude .streamlit/secrets.toml.

## Tech
- Python, Streamlit, openpyxl, requests.
- IPQS Phone Validation API:
  https://www.ipqualityscore.com/api/json/phone/{API_KEY}/{number}?country[]=US
- Runs locally (streamlit run app.py) and deploys to Streamlit Community Cloud
  from GitHub with no code changes.

## Workflow preferences
- Keep it to one or two files. Simplicity is the maintenance strategy.
- Deliver complete working code, not placeholders.
- Test against the client's 10-row sample before anything else.
