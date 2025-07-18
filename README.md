[![Download EXE](https://img.shields.io/github/v/release/AngryMunky/mtg-collection-enricher?label=Download%20EXE&logo=windows&style=for-the-badge)](https://github.com/AngryMunky/mtg-collection-enricher/releases/latest/download/mtgscan.exe)


# MTG Collection Enricher GUI  — v 2.75

A Windows desktop tool that enriches a ManaBox export with full Scryfall data and
creates a filter-ready Excel workbook (one sheet, drop-down filters on every
column).  
Features include:

* Random MTG card image that cycles every 10 s  
* Bulk Scryfall data auto-download / update (~500 MB, cached locally)  
* Split **Card / Type** columns, **Set Name**, auto-width columns  
* Status log with progress bar and **Version 2.75** label  
* “Choose save location” dialog each run; log shows full output path  
* One-click EXE build via PyInstaller

| Column order |
|--------------|
| Card Name • Color • Rarity • Mana Value • **Card** • **Type** • Set Name • Foil • Quantity • Collector Number • ManaBox ID • Scryfall ID |

---

## 🖥️ Quick Start (from source)

```bash
git clone https://github.com/<your-username>/mtg-collection-enricher.git
cd mtg-collection-enricher
pip install -r requirements.txt      # pillow, pandas, openpyxl, requests
python mtgscan.py

