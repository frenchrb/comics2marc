Script to transform comic book metadata in spreadsheet format into MARC records

## Requirements
Created and tested with Python 3.5. Requires xlrd and pymarc. 

To set up environment from file: ```conda env create –f=environment.yml```

## Input file
Subfield delimiters in spreadsheet must be '$' with no spaces before or after the subfield.
```Zelenetz, Alan,$1https://www.wikidata.org/wiki/Q4708123```

## Usage
```python comics2marc.py input_spreadsheet.xlsx```
MARC file is saved in the same directory.

## Contact
Rebecca B. French - <https://github.com/frenchrb>
