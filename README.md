# Amazon.ae Vacuum Cleaner Scraper

Automated script to collect Amazon.ae vacuum cleaner data with seller information (Fulfilled by, Sold by, Shipper/Seller) for filtered ASINs.

## Features

✅ **CSV-Based ASIN Loading** - Process only your pre-filtered ASINs  
✅ **3-Column Seller Extraction** - Captures Fulfilled by, Sold by, and Shipper/Seller  
✅ **API Integration** - Uses Rainforest API for product data  
✅ **Web Scraping** - Selenium-based extraction for reliable seller info  
✅ **Excel Output** - Professional formatting with summary and details  
✅ **Scheduled Execution** - GitHub Actions workflow for automated runs  

## Prerequisites

- Python 3.8+
- Google Chrome browser
- ChromeDriver (auto-detected or manual path)
- Rainforest API key (free tier available)

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/amazon-vacuum-scraper.git
cd amazon-vacuum-scraper
```

### 2. Create Virtual Environment
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Setup Configuration
```bash
# Copy the example config
cp .env.example .env

# Edit .env with your settings
# Required: RAINFOREST_API_KEY
```

### 5. Add Your CSV File
Place your `ASIN.csv` file in the project root directory. Format:
```
Brand,(Multiple Items)

Row Labels,
B00DFBDALC
B01EX9D3UW
B079V2CCKX
...
```

## Usage

### Local Execution
```bash
python amazon_vacuum_scraper_improved.py
```

### With Custom Config File Path
```bash
CSV_FILE_PATH=/path/to/your/ASIN.csv python amazon_vacuum_scraper_improved.py
```

### Output
- **File**: `Amazon_Vacuum_Cleaners_Filtered.xlsx`
- **Sheets**:
  - **Summary** - Collection statistics
  - **All Products** - Complete product list with 3-column seller info

## Configuration

All settings can be configured via `.env` file:

| Variable | Description | Default |
|----------|-------------|---------|
| `RAINFOREST_API_KEY` | Your Rainforest API key | - |
| `AMAZON_DOMAIN` | Amazon domain | `amazon.ae` |
| `CSV_FILE_PATH` | Path to ASIN CSV file | `ASIN.csv` |
| `OUTPUT_DIR` | Output directory for Excel file | `./` |
| `OUTPUT_FILENAME` | Output Excel filename | `Amazon_Vacuum_Cleaners_Filtered.xlsx` |
| `REQUEST_DELAY` | Delay between API requests (seconds) | `1` |
| `WAIT_TIMEOUT` | Selenium wait timeout (seconds) | `10` |

## Scheduled Execution (GitHub Actions)

### 1. Create GitHub Secrets
Go to your repository → Settings → Secrets and add:
- `RAINFOREST_API_KEY` - Your API key
- `CSV_URL` - (Optional) URL to download CSV file

### 2. Workflow Already Configured
The `.github/workflows/schedule.yml` is included and will:
- Run daily at 00:00 UTC
- Collect product data
- Upload results as artifact
- (Optional) Push results to repository

### 3. Modify Schedule
Edit `.github/workflows/schedule.yml` to change execution time:
```yaml
schedule:
  - cron: '0 0 * * *'  # UTC time - change as needed
```

## API Key Setup

1. Visit [Rainforest API](https://www.rainforestapi.com/)
2. Sign up for a free account
3. Copy your API key
4. Add to `.env` file:
   ```
   RAINFOREST_API_KEY=your_key_here
   ```

## Troubleshooting

### Chrome/ChromeDriver Issues
```bash
# Install ChromeDriver (macOS)
brew install chromedriver

# Or specify path in .env
CHROME_DRIVER_PATH=/usr/local/bin/chromedriver
```

### API Rate Limiting
Increase `REQUEST_DELAY` in `.env`:
```
REQUEST_DELAY=2  # 2 seconds between requests
```

### Selenium Timeout
Increase `WAIT_TIMEOUT` in `.env`:
```
WAIT_TIMEOUT=15  # 15 seconds to load page
```

### CSV Not Found
Ensure `ASIN.csv` is in the project root and check `CSV_FILE_PATH` in `.env`

## Output Format

Excel file includes:

**Summary Sheet**
- Total products collected
- Total ASINs processed
- Collection date/time
- Average rating

**All Products Sheet**
| Column | Description |
|--------|-------------|
| ASIN | Amazon product ID |
| Title | Product name |
| Price | Product price |
| Currency | Currency (AED) |
| **Fulfilled by** | Who fulfills the order |
| **Sold by** | Merchant/seller name |
| **Shipper/Seller** | Shipper info (when others unavailable) |
| Rating | Customer rating |
| Reviews | Number of reviews |
| URL | Amazon product URL |

## Development

### Project Structure
```
amazon-vacuum-scraper/
├── amazon_vacuum_scraper_improved.py  # Main script
├── config.py                          # Configuration management
├── requirements.txt                   # Python dependencies
├── ASIN.csv                          # Your filtered ASINs
├── .env                              # Local configuration (gitignored)
├── .env.example                      # Configuration template
├── .gitignore                        # Git ignore rules
├── README.md                         # This file
└── .github/workflows/
    └── schedule.yml                  # GitHub Actions workflow
```

### Running Tests
```bash
# (Add your test suite here)
python -m pytest tests/
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

MIT License - See LICENSE file for details

## Support

For issues and questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Review the Troubleshooting section

## Changelog

### v2.0
- Added CSV-based ASIN loading
- Enhanced 3-column seller extraction
- Improved error handling and logging
- Added GitHub Actions workflow
- Better configuration management

### v1.0
- Initial release with brand-based search

## Future Enhancements

- [ ] Database integration for historical tracking
- [ ] Email notifications on completion
- [ ] Discord/Slack notifications
- [ ] Multi-domain support (amazon.com, amazon.co.uk, etc.)
- [ ] Advanced filtering and analytics
- [ ] Web dashboard for results visualization

---

**Last Updated**: August 2026  
**Maintainer**: Your Name
