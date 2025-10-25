

A Python script to scrape data from the NAFDAC Greenbook website (https://greenbook.nafdac.gov.ng/) 



## Requirements
- Python 3.x
- Chrome browser
- Required Python packages:
  - selenium
  - webdriver_manager
  - openpyxl

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/nafdac-greenbook-scraper.git
cd nafdac-greenbook-scraper
```

2. Install required packages:
```bash
pip install selenium webdriver_manager openpyxl
```

## Usage

Simply run the script:
```bash
python run.py
```

The script will create an Excel file named `nafdac_greenbook.xlsx` containing all the scraped data.

## Data Structure

The following data is collected for each product:
- Product Name
- Active Ingredient
- Dosage Form
- Product Category
- NAFDAC Registration Number
- Applicant
- Manufacturer
- Approval Date

