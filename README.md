# XLSX to Json for Django
The program convert xlsx to json for Django's test case.

# Demo
Json file export in same directory as xlsx file.

# Features
You can select file by GUI!

# Requirement
- et-xmlfile 1.1.0
- openpyxl 3.0.9

# Installation
---bash
git clone https://github.com/Cicely-s/xlsx_to_json.git
pip install -r requirements.txt
---

# Usage
```bash
cd xlsx_to_json
python main.py
```

# Note
The program convert Django fixture file below.
```fixture.json
[
    {
        "models": "SheetName",
        "pk": 1,
        "fields": {
            "A1_val": 3,
            "B1_val": 1489,
            "C1_val": "2021-09-20T00:00:00",
            "D1_val": 36.192,
            "E1_val": 927,
            "F1_val": 404192,
            "G1_val": 40419,
            "H1_val": 0,
            "I1_val": 444611,
            "J1_val": "1489_0000003_bill.PDF",
            "K1_val": false
        }
    },
    {
        "models": "SheetName",
        "pk": 2,
        "fields": {
            "A1_val": 75,
            "B1_val": 1489,
            "C1_val": "2021-09-20T00:00:00",
            "D1_val": 44.717,
            "E1_val": 133,
            "F1_val": 375262,
            "G1_val": 37526,
            "H1_val": 13772,
            "I1_val": 426560,
            "J1_val": "1489_0000075_bill.PDF",
            "K1_val": false
        }
    }
]
```

# Author
- Cicely-s

# License
"XLSX to Json for Django" is under [MIT License](https://en.wikipedia.org/wiki/MIT_License)
