# TTC Subway Delay Analysis

Analyzing TTC subway delay incidents from the City of Toronto's open data, to find which lines, causes and months drive delays.

Data: [City of Toronto Open Data – TTC Subway Delay Data](https://open.toronto.ca/dataset/ttc-subway-delay-data/)

## Setup

```
pip install -r requirements.txt
python download_data.py
```

This downloads the 2024 and 2025-onward delay files plus the delay code descriptions into `data/raw/`.
