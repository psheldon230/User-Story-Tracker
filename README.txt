NCIM Story Tracker
==================

Live URL: https://user-story-tracker.onrender.com
(Works from any browser including VDI - may take ~30 seconds to load if it hasn't been used recently)


HOW TO USE
----------

There are 3 tabs depending on your setup:


1. FULL SYNC (same machine, no VDI)
   - In Jira, run your filter and export as CSV
   - Download your Excel tracker from SharePoint
   - Upload both files and click Sync Stories
   - Download the updated Excel and upload it back to SharePoint


2. VDI: PARSE CSV (do this step on VDI)
   - In Jira, run your filter and export as CSV
   - Upload the CSV and click Parse & Copy JSON
   - Click Copy to copy the JSON text
   - Switch to your non-VDI machine and open the tool


3. SYNC JSON WITH TRACKER (do this step outside of VDI)
   - Paste the JSON text copied from VDI
   - Download your Excel tracker from SharePoint and upload it
   - Click Sync Stories
   - Download the updated Excel and upload it back to SharePoint


RUNNING LOCALLY
---------------
Requires Python and Git installed.

1. Clone the repo:
   git clone https://github.com/psheldon230/user-story-tracker.git

2. Navigate to the folder:
   cd user-story-tracker

3. Install dependencies:
   pip install -r requirements.txt

4. Double-click run.bat to start the app (also pulls latest changes)

To get future updates:
   cd user-story-tracker
   git pull
