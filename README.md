# TTC Subway Delay Analysis

Analyzing TTC subway delay incidents from the City of Toronto's open data, to find which lines, causes and months drive delays.

Data: [City of Toronto Open Data – TTC Subway Delay Data](https://open.toronto.ca/dataset/ttc-subway-delay-data/)

## Setup

```
pip install -r requirements.txt
python download_data.py
python clean_data.py
python build_dashboard.py
```

This downloads the 2024 and 2025-onward delay files plus the delay code descriptions into `data/raw/`.

`clean_data.py` removes duplicates, fixes inconsistent line and station names, adds the delay code descriptions and flags 0 minute and very long delays. The cleaned file goes to `data/clean/`.

`build_dashboard.py` makes an Excel dashboard in `output/` with delays by line, cause and month, plus a sheet on train equipment failures.
