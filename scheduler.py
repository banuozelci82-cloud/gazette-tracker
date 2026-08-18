import schedule
import time
import subprocess
import os

def run_alert():
    print("Running daily alert...")
    subprocess.run(["python", "daily_alert.py"])

# Run every day at 8am
schedule.every().day.at("08:00").do(run_alert)

# Also run once on startup to test
run_alert()

print("Scheduler started - running daily at 8am")
while True:
    schedule.run_pending()
    time.sleep(60)
