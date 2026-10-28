# Caraway & Jain: Phase 2 data collection

Collects official sustainability communications from **1 Jan 2024 to 31 May 2026** for
USA, Canada, UK, EU, New Zealand, South Korea, China, India, Singapore and Hong Kong.
It runs **one country at a time**, then combines the countries into one month-sorted workbook.

## Files

| File | What it is |
|---|---|
| `caraway_scraper.py` | The scraper. Don't paste it into notebook cells; import it. |
| `Caraway_Runner.ipynb` | One cell per country, then Combine, then save to GitHub |
| `Caraway_Jain_Data_Classification_Sheet.xlsx` (or `..._General.csv` + `..._Codewords.csv`) | Sites and keywords. Must sit next to the script. |
| `manual_urls/` | Optional lists of links collected by hand (see section 4) |
| `outputs/countries/<Country>/` | One folder per finished country: `articles.csv`, `corpus.jsonl`, `source_log.csv`, `<Country>_Articles.xlsx` |
| `outputs/Caraway_Jain_Articles_ALL.xlsx` | Made by **Combine**: all articles, one tab per month, monthly counts per country (empty months shaded), Source Log |

## 1. Normal use (Codespaces or your laptop)

Open `Caraway_Runner.ipynb` and run it top to bottom, or use the command line:

```bash
python caraway_scraper.py --list                    # countries and their sources
python caraway_scraper.py --country Canada --probe  # optional health check
python caraway_scraper.py --country Canada          # scrape one country
python caraway_scraper.py --combine                 # merge all finished countries
```

* Re-running a country overwrites only that country's folder. Pages are cached in
  `cache/<Country>.sqlite`, so an interrupted run resumes where it stopped.
* Countries can be run on different machines and days. Commit the
  `outputs/countries/<Country>/` folder, pull it on the other machine, then run `--combine`.
* Trial run: add `--limit-per-source 20`. A trial run only takes each site's **newest**
  20 items, so it will always look sparse. Don't judge coverage from a trial run.

## 2. Why sites block us, and what the code now does about it

Codespaces runs in a Microsoft data centre. Many government sites (all three US sites,
Canada's PM, NZ DPMC, India MoEF) refuse requests from data-centre IP addresses.
Others answer with a "checking your browser" page instead of the article. The code now:

1. **Detects blocking.** It recognises "Access denied" and "checking your browser" pages
   and records them in the Source Log's **Blocked** column instead of calling them
   "no keywords". Blocked pages are never cached, so the next run tries them again.
2. **Falls back to the Internet Archive.** When an article is blocked, its saved copy on
   web.archive.org is used. **Fetched Via = wayback** in the output shows this.
3. **Discovers blocked sites through the archive.** If a site's listing page is blocked,
   the code asks the Internet Archive for every page it saved under that site's news
   folder in the study period. EPA and the State Department are set up this way by default.
4. **Uses official APIs where they exist.** These are the Federal Register (US executive
   orders, memoranda and proclamations, which your Step 1 lists first) and GOV.UK.

These fallbacks help, but archive coverage is incomplete. **The biggest single
improvement is step 3 below: run the blocked countries on your own computer.**

## 3. Steps for overcoming blocking, in order

**Step A. Find out which sites are blocked.** Run `--probe` for each country, or look at
the Source Log tab after a run. "BLOCKED" or a high **Blocked** count marks the sites to fix.

**Step B. Run those countries on your own computer** (home or university network).
Sites that refuse Codespaces usually accept a normal household connection. On Windows (PowerShell):

```powershell
git clone https://github.com/<your-account>/Discourse_Analysis_Summer_Research.git
cd Discourse_Analysis_Summer_Research
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python caraway_scraper.py --country USA --probe
python caraway_scraper.py --country USA
git add outputs/countries/USA
git commit -m "USA scrape"
git push
```

Then in Codespaces: `git pull`, and run Combine. Keep the laptop awake and plugged in
during long runs. If it sleeps, run the same command again and it resumes.

**Step C. Use a real browser for JavaScript or still-blocked sites (EU press corner, Singapore SGPC).**
On your own computer, add `--js`. To use your installed Google Chrome, also add `--browser chrome`:

```powershell
python caraway_scraper.py --country EU --js --browser chrome
```

With `--js`, pages are rendered in a real browser, and blocked pages are retried in it before
falling back to the archive. `--headed` shows the window, which is useful to see what a site does.
If you don't have Chrome, run `playwright install chromium` once and leave out `--browser chrome`.
This mode doesn't work in Codespaces, because it's still a data-centre IP.

**Step D. Slow down.** Some sites block for speed, not location. `--delay 5` waits 5 seconds
between requests to the same site (default 2).

**Step E. Collect links manually for sites that are only forms** (India PIB). See section 4.

**Don't** use rotating proxies, VPN-hopping, "stealth" plugins or CAPTCHA-solving services
to get around a site that refuses you. They breach most sites' terms of use, can get your
university network banned, and are hard to defend in a methods section. The script also
obeys each site's `robots.txt`. Your methods section can say: "collected with a
rate-limited scraper that respected robots.txt; where sites refused automated requests,
archived copies from the Internet Archive Wayback Machine were used."

## 4. Manual link lists (and where Claude in Chrome fits)

Any CSV in `manual_urls/` named after a country (e.g. `manual_urls/India.csv`) is added
to that country's run. The code fetches, dates, keyword-filters and classifies those links
like everything else. Columns: `url` (required), and optionally `date`, `title`, `source_type`.
See `manual_urls/_TEMPLATE.csv`.

**Claude in Chrome** is a browser extension that lets Claude operate *your* Chrome tab by tab.
It is not a bulk scraper: it can't be called from this Python code or from Git, and paging
through thousands of releases with it would be slow. Where it is useful:

* **Collecting links from form-based sites.** Example: on PIB, choose Ministry of Environment
  and a month, and copy every release link into `manual_urls/India.csv`. Repeat for each month.
  Then commit the CSV (`git add manual_urls/India.csv`, commit, push) and run `--country India`.
* **Showing a site's structure.** Ask it what the article URLs and the "next page" links look
  like on a site that shows `0 article links`, then put that into `SOURCE_OVERRIDES`
  (section 1 of the script) as `link_pattern` / `page_param`.

## 5. Fixing a source that finds 0 links

In `caraway_scraper.py`, section 1, `SOURCE_OVERRIDES` holds one entry per site.
The fields that matter most:

* `link_pattern`: a regex every article URL matches, e.g. `r"beehive\.govt\.nz/release/[^/?#]+$"`.
* `page_param` / `page_start`: how pages are numbered, e.g. `?page=0,1,2` gives `page_param="page", page_start=0`.
  If you leave these out, the code tries common patterns itself.
* `wayback_prefix`: where the site's articles live (e.g. `"www.epa.gov/newsreleases/"`).
  This enables Internet Archive discovery when the site blocks you.
* `render_js=True`: the site needs a browser, so run it with `--js`.

## 6. Notes for the analysis

* **Hong Kong's** daily index covers *every* HK government release. Only keyword-matching
  titles are kept, but HK can still produce many more documents than other countries.
  Consider capping or weighting per jurisdiction.
* **Strong Match = No** rows matched only broad words such as "energy" or "green". Filter
  on Strong Match = Yes for a cleaner corpus.
* **Date Source = first archive capture** is approximate, and is used only when no
  publication date was found on the page.
