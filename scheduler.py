import schedule
import time
import subprocess
import os

def run_alert():
    print("Running daily alert...")
    subprocess.run(["python", "daily_alert.py"])

def run_cayman():
    print("Running weekly Cayman Islands gazette scrape...")
    subprocess.run(["python", "cayman_scraper.py"])

# Run every day at 8am
schedule.every().day.at("08:00").do(run_alert)

# Cayman Gazette is published fortnightly, so a weekly check is enough —
# refresh_cayman() no-ops (no Claude API cost) if there's nothing new.
schedule.every().monday.at("09:00").do(run_cayman)

# Also run once on startup to test
run_alert()
run_cayman()

print("Scheduler started - running daily alert at 8am, Cayman scrape weekly on Mondays at 9am")
while True:
    schedule.run_pending()
    time.sleep(60)
