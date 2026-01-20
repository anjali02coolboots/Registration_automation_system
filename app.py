"""
AUTOMATED REGISTRATION REPORT ORCHESTRATOR
==========================================
This script orchestrates the entire automated registration reporting workflow:
1. Scrapes TechGig registration data for yesterday
2. Processes the data with source lookup and appends to template
3. Generates formatted Excel template sheet
4. Sends report via Gmail

All heavy lifting is done in sub-scripts for clean, maintainable code.
"""

import os
import sys
import subprocess

# ================= CONFIGURATION =================
EXCEL_TEMPLATE = 'Registration_Template.xlsx'
RECIPIENT_EMAIL = "ankit.k@coolbootsmedia.com"
EMAIL_SUBJECT = "📊 Automated Report - Registration Template"


# ================= MAIN ORCHESTRATION =================
def main():
    print("="*60)
    print("AUTOMATED REGISTRATION REPORT WORKFLOW")
    print("="*60)
    
    try:
        # STEP 1: Scrape TechGig Registration Data (Yesterday's Data)
        print("\n[STEP 1/4] Running TechGig scraper...")
        print("Downloading yesterday's registration data...")
        result = subprocess.run([sys.executable, "techgig_scraper.py"], check=True)
        print("✅ Registration data downloaded")
        
        # STEP 2: Process Data (Lookup + Append to Template)
        print("\n[STEP 2/4] Processing registration data...")
        print("Performing source lookup and appending to template...")
        result = subprocess.run([sys.executable, "data_processor.py"], check=True)
        print(f"✅ Data processed and appended to {EXCEL_TEMPLATE}")
        
        # STEP 3: Generate Formatted Excel Template Sheet
        print("\n[STEP 3/4] Generating formatted template sheet...")
        result = subprocess.run([sys.executable, "generate_template.py"], check=True)
        final_template_path = os.path.abspath(EXCEL_TEMPLATE)
        print(f"✅ Formatted template ready: {final_template_path}")
        
        # STEP 4: Send via Gmail
        print("\n[STEP 4/4] Sending via Gmail...")
        result = subprocess.run([sys.executable, "gmail_sender.py"], check=True)
        
        print("\n" + "="*60)
        print("🎉 AUTOMATION COMPLETE!")
        print("="*60)
        print(f"📊 Report generated: {final_template_path}")
        print(f"📧 Sent via Gmail to: {RECIPIENT_EMAIL}")
        
    except subprocess.CalledProcessError as e:
        print("\n" + "="*60)
        print("❌ AUTOMATION FAILED!")
        print("="*60)
        print(f"Error: Script failed with exit code {e.returncode}")
        sys.exit(1)
    except Exception as e:
        print("\n" + "="*60)
        print("❌ AUTOMATION FAILED!")
        print("="*60)
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()


# ================= SCRIPT SUMMARY =================
"""
### **Sub-Scripts:**
1. **techgig_scraper.py** - Downloads yesterday's registration data from TechGig
2. **data_processor.py** - Performs source lookup and appends to template
3. **generate_template.py** - Generates formatted Excel template sheet
4. **gmail_sender.py** - Sends report via Gmail

### **File Structure:**
Report 4/
├── app.py                              # THIS FILE - Orchestrator
├── techgig_scraper.py                  # Script 1
├── data_processor.py                   # Script 2
├── generate_template.py                # Script 3
├── gmail_sender.py                     # Script 4 (Gmail sender)
├── credentials_store.py                # TechGig credentials
├── credentials.json                    # Gmail API credentials
├── token.pickle                        # Gmail API token (auto-generated)
├── Source_TG_Latest.xlsx               # Lookup file
├── Registration_Template.xlsx          # Main template
└── exports/                            # Generated files

### **Prerequisites:**
- Gmail API credentials.json file in the same directory
- First run will require browser authentication for Gmail
- Excel file must exist before running

### **Usage:**
python app.py

This will run all 4 scripts in sequence automatically!
"""