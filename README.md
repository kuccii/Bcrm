# Business Scanner Agent

Scans all businesses from any directory website (like RwandaYP) and exports to Excel.

## Installation

```bash
cd business-scanner
npm install
```

## Usage

```bash
node index.js <source-url> [output-file]
```

## Example

```bash
node index.js https://www.rwandayp.com/location/Kigali kigali_businesses.xlsx
```

## Output

Excel file with columns:
- Business Name
- Has Website (Yes/No)

## Results

- Kigali: 2,835 businesses scanned from 143 pages
- File: kigali.xlsx

## Files

- `index.js` - Main scanner script
- `kigali.xlsx` - Sample output from RwandaYP Kigali