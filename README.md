# Day 69 — Multi-Page Web Scraping (Pagination Handling)

Part of the **84-Day Python & Advanced Excel Mastery Roadmap**
Phase 5: Web Data & Scraping

---

## 🎯 Objective
Master pagination handling — scraping data across multiple pages of a website using loops, delays, and error handling. Export unified results to Excel.

## 🧠 Concepts Covered
- Pagination patterns (URL parameter, path-based, offset)
- Looping through page numbers with `range()`
- Rate limiting via `time.sleep()`
- Custom HTTP `User-Agent` headers
- Auto-detecting the last page (stop condition)
- Error handling with `try-except` for network failures
- Exporting scraped data to Excel with `pandas`

## 🛠️ Tech Stack
- Python 3.10+
- `requests` — HTTP calls
- `beautifulsoup4` — HTML parsing
- `pandas` + `openpyxl` — Excel export

## 📦 Installation
```bash
pip install requests beautifulsoup4 pandas openpyxl
