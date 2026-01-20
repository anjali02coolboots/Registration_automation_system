import os

TECHGIG_USERNAME = os.getenv('TECHGIG_USERNAME', '')
TECHGIG_PASSWORD = os.getenv('TECHGIG_PASSWORD', '')

LOGIN_URL = "https://www.techgig.com/mis/link.php"
STATS_URL = "https://www.techgig.com/mis/mis_tg_reg_stats.php"

OUTPUT_DIR = "exports"
NAV_TIMEOUT_MS = 60000
HEADLESS = True   # set False to see browser