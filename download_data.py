"""
Downloads the TTC subway delay files from Toronto Open Data into data/raw/.
"""
from pathlib import Path
from urllib.request import urlretrieve

BASE = ("https://ckan0.cf.opendata.inter.prod-toronto.ca/dataset/"
        "996cfe8d-fb35-40ce-b569-698d51fc683b/resource/")
FILES = {
    "delays_2024.xlsx": "2ee1a65c-da06-4ad1-bdfb-b1a57701e46a/download/ttc-subway-delay-2024.xlsx",
    "delays_since_2025.csv": "0b6e5c52-e993-46d6-8d74-8602ee224457/download/ttc-subway-delay-data-since-2025.csv",
    "delay_codes.csv": "b2d8f5e0-0997-46b5-8abd-caa685a0290b/download/code-descriptions.csv",
}

raw = Path("data/raw")
raw.mkdir(parents=True, exist_ok=True)
for name, path in FILES.items():
    print(f"Downloading {name}...")
    urlretrieve(BASE + path, raw / name)
print("Done.")
