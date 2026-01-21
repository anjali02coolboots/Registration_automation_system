import os

# ================= TECHGIG CREDENTIALS =================
# Priority: Environment variables > Hardcoded values
# This allows GitHub Actions to use secrets

TECHGIG_USERNAME = os.getenv('TECHGIG_USERNAME', 'your_username_here')
TECHGIG_PASSWORD = os.getenv('TECHGIG_PASSWORD', 'your_password_here')

# ================= TECHGIG URLS =================
LOGIN_URL = "https://www.techgig.com/mis/link.php"
STATS_URL = "https://www.techgig.com/mis/mis_tg_reg_stats.php"

# ================= SETTINGS =================
OUTPUT_DIR = "exports"
HEADLESS = os.getenv('HEADLESS', 'True').lower() == 'true'  # Default to headless
NAV_TIMEOUT_MS = int(os.getenv('NAV_TIMEOUT_MS', '60000'))  # 60 seconds default

# ================= VALIDATION =================
if TECHGIG_USERNAME == 'your_username_here' or TECHGIG_PASSWORD == 'your_password_here':
    if not os.getenv('GITHUB_ACTIONS'):
        print("\n⚠️  WARNING: TechGig credentials not configured!")
        print("Please update credentials_store.py or set environment variables:")
        print("  - TECHGIG_USERNAME")
        print("  - TECHGIG_PASSWORD")
        print()
    else:
        # In GitHub Actions, credentials MUST come from secrets
        if not os.getenv('TECHGIG_USERNAME') or not os.getenv('TECHGIG_PASSWORD'):
            raise RuntimeError(
                "TechGig credentials not found in environment variables.\n"
                "Please add TECHGIG_USERNAME and TECHGIG_PASSWORD to GitHub Secrets."
            )