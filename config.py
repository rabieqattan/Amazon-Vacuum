"""
Configuration file for Amazon Vacuum Scraper
Load settings from environment variables or .env file
"""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# API Configuration
RAINFOREST_API_KEY = os.getenv('RAINFOREST_API_KEY', 'AECB78A27B374B34A2F037C3A6E1AA3B')
AMAZON_DOMAIN = os.getenv('AMAZON_DOMAIN', 'amazon.ae')

# File Paths
CSV_FILE_PATH = os.getenv('CSV_FILE_PATH', 'ASIN.csv')
OUTPUT_DIR = os.getenv('OUTPUT_DIR', './')
OUTPUT_FILENAME = os.getenv('OUTPUT_FILENAME', 'Amazon_Vacuum_Cleaners_Filtered.xlsx')

# Selenium Configuration
CHROME_DRIVER_PATH = os.getenv('CHROME_DRIVER_PATH', None)  # None = auto-detect
HEADLESS_MODE = os.getenv('HEADLESS_MODE', 'True').lower() == 'true'
WAIT_TIMEOUT = int(os.getenv('WAIT_TIMEOUT', '10'))

# Request Configuration
REQUEST_DELAY = int(os.getenv('REQUEST_DELAY', '1'))  # Seconds between requests
CONNECTION_TIMEOUT = int(os.getenv('CONNECTION_TIMEOUT', '30'))  # Seconds

# Logging Configuration
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
LOG_FILE = os.getenv('LOG_FILE', 'scraper.log')

# Script Configuration
MAX_RETRIES = int(os.getenv('MAX_RETRIES', '3'))
SKIP_FAILED_ASINS = os.getenv('SKIP_FAILED_ASINS', 'False').lower() == 'true'
