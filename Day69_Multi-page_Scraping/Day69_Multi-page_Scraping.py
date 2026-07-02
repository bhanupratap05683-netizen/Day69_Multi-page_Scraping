"""
================================================================
DAY 69: MULTI-PAGE WEB SCRAPING (PAGINATION HANDLING)
84-Day Python & Advanced Excel Mastery Roadmap
Author: Bhanu Pratap Singh
================================================================
Topics Covered:
    1. Basic page loop (range)
    2. Scraping across multiple pages
    3. Rate limiting with time.sleep()
    4. Custom headers (User-Agent)
    5. Auto-detect last page
    6. Error handling in scraping
    7. Export multi-page data to Excel
================================================================
Requirements:
    pip install requests beautifulsoup4 pandas openpyxl
================================================================
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time


# ---------- HEADERS (Look like a real browser) ----------
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


# =================================================================
# EXERCISE 1: BASIC PAGE LOOP
# =================================================================
def exercise_1_page_urls():
    """Print URLs for pages 1-5 using a loop."""
    print("\n" + "=" * 60)
    print("EXERCISE 1: Building paginated URLs")
    print("=" * 60)
    base_url = "https://example.com/stocks?page="
    for page in range(1, 6):
        print(f"Page {page}: {base_url}{page}")


# =================================================================
# EXERCISE 2 + 3: SCRAPE QUOTES + ADD DELAYS
# =================================================================
def scrape_quotes_fixed_pages(max_pages=5):
    """Scrape quotes.toscrape.com — pages 1 to max_pages."""
    print("\n" + "=" * 60)
    print(f"EXERCISE 2+3: Scraping {max_pages} pages of quotes")
    print("=" * 60)

    all_quotes = []                       # Accumulator list
    base_url = "http://quotes.toscrape.com/page/{}/"

    for page_num in range(1, max_pages + 1):
        url = base_url.format(page_num)
        print(f"\n📄 Scraping Page {page_num}: {url}")

        try:
            response = requests.get(url, headers=HEADERS, timeout=10)
            response.raise_for_status()   # Raise if 4xx/5xx error
        except requests.exceptions.RequestException as e:
            print(f"❌ Error fetching {url}: {e}")
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        quotes = soup.find_all("div", class_="quote")

        for q in quotes:
            text = q.find("span", class_="text").get_text(strip=True)
            author = q.find("small", class_="author").get_text(strip=True)
            all_quotes.append({"page": page_num, "quote": text, "author": author})

        print(f"   ✅ Got {len(quotes)} quotes from page {page_num}")

        # RATE LIMITING — Be polite to the server
        time.sleep(1)

    print(f"\n🎯 TOTAL QUOTES SCRAPED: {len(all_quotes)}")
    return all_quotes


# =================================================================
# EXERCISE 4: AUTO-DETECT LAST PAGE
# =================================================================
def scrape_quotes_auto_detect():
    """Keep scraping until a page returns 0 quotes."""
    print("\n" + "=" * 60)
    print("EXERCISE 4: Auto-detecting last page")
    print("=" * 60)

    all_quotes = []
    page_num = 1
    base_url = "http://quotes.toscrape.com/page/{}/"

    while True:
        url = base_url.format(page_num)
        print(f"📄 Page {page_num} ...", end=" ")

        try:
            response = requests.get(url, headers=HEADERS, timeout=10)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"❌ Error: {e}")
            break

        soup = BeautifulSoup(response.text, "html.parser")
        quotes = soup.find_all("div", class_="quote")

        # STOP CONDITION: no quotes on this page
        if not quotes:
            print("No more quotes. Stopping.")
            break

        for q in quotes:
            all_quotes.append({
                "page": page_num,
                "quote": q.find("span", class_="text").get_text(strip=True),
                "author": q.find("small", class_="author").get_text(strip=True),
            })

        print(f"got {len(quotes)} quotes ✅")
        page_num += 1
        time.sleep(1)

    print(f"\n🎯 Total pages scraped: {page_num - 1}")
    print(f"🎯 Total quotes collected: {len(all_quotes)}")
    return all_quotes


# =================================================================
# EXERCISE 5: EXPORT TO EXCEL
# =================================================================
def export_to_excel(data, filename="quotes_output.xlsx"):
    """Save list of dicts to Excel using pandas."""
    print("\n" + "=" * 60)
    print(f"EXERCISE 5: Exporting to {filename}")
    print("=" * 60)

    if not data:
        print("⚠️ No data to export.")
        return

    df = pd.DataFrame(data)
    df.to_excel(filename, index=False, sheet_name="Quotes")
    print(f"✅ Saved {len(df)} rows to {filename}")


# =================================================================
# EXERCISE 7 (CHALLENGE): SCRAPE BOOKS (Title + Price)
# =================================================================
def scrape_books_multipage(max_pages=3):
    """Scrape books.toscrape.com — title + price across pages."""
    print("\n" + "=" * 60)
    print(f"CHALLENGE: Scraping {max_pages} pages of books")
    print("=" * 60)

    all_books = []
    base_url = "http://books.toscrape.com/catalogue/page-{}.html"

    for page_num in range(1, max_pages + 1):
        url = base_url.format(page_num)
        print(f"\n📚 Page {page_num}: {url}")

        try:
            response = requests.get(url, headers=HEADERS, timeout=10)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"❌ Error: {e}")
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        books = soup.find_all("article", class_="product_pod")

        for book in books:
            title = book.h3.a["title"]
            price = book.find("p", class_="price_color").get_text(strip=True)
            all_books.append({"page": page_num, "title": title, "price": price})

        print(f"   ✅ Got {len(books)} books")
        time.sleep(1)

    # Save to Excel
    df = pd.DataFrame(all_books)
    df.to_excel("books_output.xlsx", index=False, sheet_name="Books")
    print(f"\n🎯 Total books: {len(all_books)} → saved to books_output.xlsx")
    return all_books


# =================================================================
# MAIN DRIVER
# =================================================================
def main():
    print("\n" + "🐍" * 30)
    print("DAY 69 — MULTI-PAGE WEB SCRAPING")
    print("🐍" * 30)

    # Ex 1
    exercise_1_page_urls()

    # Ex 2 + 3
    quotes_fixed = scrape_quotes_fixed_pages(max_pages=3)

    # Ex 4
    quotes_all = scrape_quotes_auto_detect()

    # Ex 5
    export_to_excel(quotes_all, "quotes_output.xlsx")

    # Ex 7
    scrape_books_multipage(max_pages=3)

    print("\n" + "✅" * 30)
    print("DAY 69 COMPLETE — All exercises executed")
    print("✅" * 30 + "\n")


if __name__ == "__main__":
    main()