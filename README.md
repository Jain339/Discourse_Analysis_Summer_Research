# Discourse_Analysis_Summer_Research
Sharvi Scraper (Info)
Caraway & Jain -- Phase 2: data collection
==========================================
 
Collects official sustainability-related communications (Jan 2024 -> present)
from the websites listed in Caraway_Jain_Data_Classification_Sheet.xlsx,
keeps the ones that mention the codewords, and writes:
 
  outputs/Caraway_Jain_Articles.xlsx   links, sorted and split month-wise
  outputs/articles.csv                 same table as plain CSV
  outputs/corpus.jsonl                 full text + metadata (for Steps 2-3)
 
The spreadsheet is the single source of truth: add a URL or a keyword there
and the next run picks it up. Site-specific tweaks live in SOURCE_OVERRIDES.
 
Typical use
-----------
  pip install -r requirements.txt
  python caraway_scraper.py --probe                 # check every site first
  python caraway_scraper.py --limit-per-source 20   # quick trial run
  python caraway_scraper.py                         # full run (hours; resumable)
  python caraway_scraper.py --offline               # rebuild outputs from cache
 
In Jupyter/Colab:  from caraway_scraper import main; main(["--probe"])
 
How it works
------------
1. DISCOVER candidate article URLs per source (sitemap, listing pages with
   pagination, the GOV.UK search API, or Hong Kong's daily index pages).
2. FETCH each candidate politely (robots.txt, per-domain delay, SQLite cache,
   so an interrupted run resumes where it stopped).
3. EXTRACT title, date and main text with trafilatura.
4. FILTER to the date range and to pages that match at least one codeword.
5. WRITE the outputs.