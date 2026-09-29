
import requests
from bs4 import BeautifulSoup
from collections import Counter
import re

url = "https://advertisingweek.com/aw360/news/yes-that-youtube-ad-really-was-45-min-\
long-the-rise-of-long-form-youtube-ads/8825"

# A descriptive User-Agent helps avoid being blocked
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 \
(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

response = requests.get(url, headers=headers, timeout=10)

# Raise an error if the request failed
response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

# Find the element with class "post-content"
post_contents = soup.find_all(class_="entry-content alignfull wp-block-post-content\
 is-layout-constrained wp-block-post-content-is-layout-constrained")

if not post_contents:
    print("No elements with class 'post-content' found.")
    exit()

full_text = " ".join(
    element.get_text(separator=" ", strip=True)
    for element in post_contents
)


complaint_keywords = [
    "long",
    "unskippable",
    "intrusive",
    "annoying",
    "too many",
    "excessive",
    "forced",
    "pre-roll",
    "mid-roll",
    "interrupt",
    "frustrating",
    "ruined",
    "unbearable",
    "non-skippable",
    "ad overload",
    "ad fatigue",
    "double ads",
    "back-to-back",
    "repetitive",
    "irrelevant",
    "targeting",
    "privacy",
]

# Normalize text to lowercase for matching
text_lower = full_text.lower()

# Count occurrences of each complaint keyword (as a phrase)
keyword_counts = {}

for keyword in complaint_keywords:
    # Use regex to count whole phrases (handles spacing variations)
    pattern = re.compile(re.escape(keyword.lower()))
    count = len(pattern.findall(text_lower))
    if count > 0:
        keyword_counts[keyword] = count

# Sort by frequency (highest first)
sorted_keywords = sorted(keyword_counts.items(), key=lambda x: x[1], reverse=True)

print("Complaint Keyword Frequency:")


# Simple stop word list (you can expand this or use NLTK)
stop_words = {
    "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for",
    "of", "with", "by", "from", "is", "are", "was", "were", "be", "been",
    "it", "its", "this", "that", "these", "those", "i", "you", "he", "she",
    "we", "they", "my", "your", "his", "her", "our", "their", "as", "so",
    "if", "then", "than", "too", "very", "just", "can", "will", "would",
    "could", "should", "may", "might", "must", "has", "have", "had", "do",
    "does", "did", "not", "no", "nor", "only", "own", "same", "such",
    "more", "most", "other", "some", "any", "each", "every", "all", "both",
    "few", "many", "much", "several", "what", "which", "who", "whom",
    "when", "where", "why", "how", "there", "here", "up", "down", "out",
    "off", "over", "under", "again", "further", "once", "about", "into",
    "through", "during", "before", "after", "above", "below", "between",
    "because", "while", "although", "though", "since", "until", "unless",
    "also", "however", "therefore", "thus", "hence", "still", "yet",
    "even", "well", "back", "way", "thing", "things", "get", "got",
    "make", "made", "take", "took", "see", "saw", "know", "knew", "think",
    "thought", "say", "said", "go", "went", "come", "came", "want",
    "wanted", "use", "used", "find", "found", "give", "gave", "tell",
    "told", "work", "worked", "call", "called", "try", "tried", "ask",
    "asked", "need", "needed", "feel", "felt", "become", "became",
    "leave", "left", "put", "mean", "meant", "keep", "kept", "let",
    "begin", "began", "seem", "seemed", "help", "helped", "talk",
    "talked", "turn", "turned", "start", "started", "show", "showed",
    "hear", "heard", "play", "played", "run", "ran", "move", "moved",
    "live", "lived", "believe", "believed", "hold", "held", "bring",
    "brought", "happen", "happened", "write", "wrote", "provide",
    "provided", "sit", "sat", "stand", "stood", "lose", "lost", "pay",
    "paid", "meet", "met", "include", "included", "continue",
    "continued", "set", "learn", "learned", "change", "changed", "lead",
    "led", "understand", "understood", "watch", "watched", "follow",
    "followed", "stop", "stopped", "create", "created", "speak",
    "spoke", "read", "allow", "allowed", "add", "added", "spend",
    "spent", "grow", "grew", "open", "opened", "walk", "walked",
    "win", "won", "offer", "offered", "remember", "remembered",
    "love", "loved", "consider", "considered", "appear", "appeared",
    "buy", "bought", "wait", "waited", "serve", "served", "die",
    "died", "send", "sent", "expect", "expected", "build", "built",
    "stay", "stayed", "fall", "fell", "cut", "reach", "reached",
    "kill", "killed", "remain", "remained",
}

# Extract words (letters only, 3+ characters)
words = re.findall(r"\b[a-z]{3,}\b", text_lower)

# Filter out stop words
filtered_words = [w for w in words if w not in stop_words]

# Count word frequencies
word_counts = Counter(filtered_words)

print("\nTop 20 Most Frequent Words:")
print("-" * 40)


import pandas as pd
import os
import pandas as pd
from openpyxl import load_workbook
EXCEL_FILE = "top_20_word_frequency.xlsx"

def safe_sheet_name(base, existing_sheets):
    """Create a valid, unique Excel sheet name (max 31 chars, no invalid chars)."""
    # Remove characters Excel does not allow in sheet names
    for ch in r'\/?*[]:':
        base = base.replace(ch, '_')
    base = base[:31] or "Sheet"
    name = base
    i = 1
    while name in existing_sheets:
        suffix = f"_{i}"
        name = base[:31 - len(suffix)] + suffix
        i += 1
    return name

def append_sheet(df, sheet_base):
    """Append a DataFrame as a new sheet to EXCEL_FILE."""
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE, read_only=True)
        existing = wb.sheetnames
        wb.close()
        mode = 'a'
        kwargs = {'if_sheet_exists': 'new'}
    else:
        existing = []
        mode = 'w'
        kwargs = {}

    sheet_name = safe_sheet_name(sheet_base, existing)

    with pd.ExcelWriter(
        EXCEL_FILE,
        engine='openpyxl',
        mode=mode,
        **kwargs
    ) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Added sheet: {sheet_name}")

top_20 = word_counts.most_common(20)
df = pd.DataFrame(top_20, columns=["Word", "Frequency"])
df.insert(0, "URL", url)

    # Choose a short, unique sheet name.
    # Option A: numbered sheets
sheet_base = "Page_AdvertisingWeek1"


append_sheet(df, sheet_base)

EXCEL_FILE = "complaint_keyword_frequency.xlsx"
top_20 = Counter(keyword_counts).most_common(20)

    # Build DataFrame
df = pd.DataFrame(top_20, columns=["Keyword", "Frequency"])
df.insert(0, "URL", url)

    # Create a unique sheet name – use a counter or a short slug
sheet_base = f"Complaints_AdvertisingWeek1"
    # Optional: use domain + timestamp
    # from urllib.parse import urlparse
    # from datetime import datetime
    # sheet_base = f"{urlparse(url).netloc}_{datetime.now().strftime('%H%M%S')}"

append_sheet(df, sheet_base)
