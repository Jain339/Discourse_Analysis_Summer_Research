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

Methods (keras API- use deep neural classifiers)
Step 1: Data Collection
==========================================
The study will examine official sustainability-related communications published between January 2024 and May 2026  from selected Global North and Global South jurisdictions. The current countries considered include the United States, Canada, the United Kingdom, the European Union, New Zealand and South Korea, together with China, India, Singapore and Hong Kong. The primary corpus will be constructed through web scraping of official government websites, prioritizing executive orders, ministerial statements, regulatory announcements, government speeches, sustainability policies, and official press releases. Selected journalistic organizations may also be collected separately to aid with corpus construction.
Step 2: Corpus Construction and Preprocessing
==========================================
Clean and organize extracted data into a structured text corpus. This will involve removing duplicate documents and other irrelevant content. Each document will then be segmented into paragraph-level units, which will serve as the primary unit of analysis. Metadata including the country, institution, publication date, source and document type will be retained for every paragraph.
Step 3: Exploratory Natural Language Processing
==========================================
Exploratory NLP will be conducted on the corpus, identifying broad linguistic patterns. This stage will include keyword frequency analysis, Term Frequency–Inverse Document Frequency (TF-IDF) analysis to identify words that are particularly distinctive within specific countries or time periods, and sentiment and uncertainty analysis to examine overall changes in tone across the corpus. 
Step 4: Frame Construction and Manual Coding
==========================================
A frame codebook will then be developed to classify the dominant sustainability narratives present within government communications. Two possible approaches include: First, construct frames around sustainability governance models and reporting frameworks, examining references to institutions such as ISSB, IFRS, SASB, GRI and CSRD, together with the governance philosophies they represent. The second would employ broader thematic frames, including regulatory leadership, economic competitiveness, energy security, national sovereignty, development equity and international cooperation. Each paragraph may receive one or more frame labels depending on the ideas communicated.
A representative sample of the corpus will be manually coded by both researchers. The manually coded dataset will then serve as the labelled training data for the machine learning model.
Step 5: Stance Identification
==========================================
Following frame identification, each paragraph will be assigned a stance toward the identified frame or governance model. Stances will be categorized as supportive, neutral or mixed, or opposed, allowing the analysis to distinguish between governments that actively promote, criticize or simply acknowledge sustainability governance approaches and reporting standards.
Step 6: Machine Learning Classification
The manually coded dataset will be used to train a supervised machine learning model to classify the remaining corpus automatically. The primary approach currently being considered is Sentence-BERT embeddings combined with Logistic Regression, whereby each paragraph is first converted into a numerical representation of its contextual meaning before classification. 
Step 7: Model Validation
==========================================
Model predictions will be compared against the manually coded dataset to evaluate classification performance. Standard evaluation metrics, including precision, recall, F1-score and confusion matrices, will be used to assess the accuracy of the automated classifications before applying the selected model to the complete corpus.
Step 8: Comparative Analysis
The final stage will compare the classified corpus across countries, regions and time periods. The analysis will examine changes in
•	Countries' alignment with international sustainability reporting standards and governance frameworks (ISSB, IFRS, CSRD, GRI, SASB, etc.) following US policy retrenchment. 
•	Shifts in dominant sustainability governance frames and policy stances across countries and regions before and after the second US withdrawal from the Paris Agreement. 
•	Patterns of convergence, divergence, and influence to examine whether countries increasingly align with regional sustainability governance models and identify potential emerging centres of sustainability leadership.
