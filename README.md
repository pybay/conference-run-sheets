# conference-run-sheets
This Python repo reads the PyBay conference schedule and speaker info from Sessionize, and generates concise run sheets - timelines that coordinate talk times, speakers and transitions, keeping everyone in sync and the conference on schedule. Used by Room Captains/Moderators, AV Teams, and event staff.

***

## Conference Run Sheet Creation Utility

This Python utility provides a straightforward way to read conference schedule and speaker info from downloaded content from Sessionize, the PyBay talk proposal and speaker management system (as of 2025).

Sessionize contains all the details needed to run the conference, but is not available to all event staff and does not provide an easy way to create concise "Run Sheets" that give a concise timeline view of the talks and speakers for the conference.

**Built for the PyBay conference:** Given a PyBay event staff who has access to both Sessionize and the PyBay Google Drive for the event year, this utility allows the person to read in two spreadsheets from Sessionize, and then create the run sheets in a local Excel file that can be uploaded to the PyBay Google Drive for sharing, printing and distribution to the team.

### Key Features:
* **Generate Speaker Intro Notes:** Specifically includes select speaker intro notes as available in Sessionize: e.g. timeslot, talk title, pronunciation of speaker name, preferred intro bullets, etc.
* **Save Organizer Time:** Reduces manual workload of processing and formatting spreadsheets.
* **Reads Excel downloads:**  Reads the `Schedule` and  `Flattened Accepted Sessions` Excel files as input, after they have been manually downloaded from Sessionize.
* **Automates Data Engineering:** Processes existing Sessionize data into multiple user-friendly formats, appropriate for the AV Team, Room Captain/Moderators - including summary and detail views.
* **LLM free:** Uses only Python, and common supporting libraries like Pandas, xlsxwriter, etc. for free, deterministic, low stress processing.  Google Sheets retains much Excel formatting when uploaded. 

### How It Works:<br>
```text
(PyBay Organizer)                    (Sessionize)              (PyBay Team - local drive or Google)
 ----------------------------------------------------------------------------------------------------------------------
[Pre-Event Setup]
 
 1. Install This Repo ----------------------------------------> [ This Python Repo ]
 
 2. Login to Sessionize -------> Download Core Files:
                                    • Schedule.xlsx
                                    • *Flattened_Accepted_Sessions.xlsx
                                                    |
                                                    |
                                                    ├─--------> Save to project root
                                                                       |
                                                                       V
 3. Run main.py ---------> [ main.py ] -----------------------> Generated Excel Workbook:
                                                                       ├─ robertson_summary
                                                                       ├─ robertson_detail (+ speaker contacts)
                                                                       ├─ fisher_summary
                                                                       ├─ fisher_detail (+ speaker contacts)
                                                                       ├─ workshop_summary (if any)
                                                                       └─ workshop_detail (if any)
                                                                        |
                                                                        |
                                                                        V
 4. Review & Edit (minimal) -----------------------------------> [ Edit if needed ]
                                                                        |
                                                                        |
                                                                        V
 5. Upload to Google Drive ------------------------------------> PyBay Google Drive
                                                             (e.g., PyBay 2025 folder)
                                                                        |
                                                                        |
                                                                        V
 6. Share Google Drive Link --------------------> PyBay Logistics Chair + PyBay Volunteer Chair
                                                                        |
                                                                        |
                                                                        V
                                                          [ Team Review & Last-Minute Edits ]
                                                             (e.g., speaker cancellations)
 ----------------------------------------------------------------------------------------------------------------------
                                                                  [Day of Event]
 
                                                            [ PyBay Volunteer Chair ]
                                                                        |
                                                                        V
                                                             Print Hard copies & Distribute:
                                                                        |
                                            +---------------------------+---------------------------+
                                            |                           |                           |
                                            V                           V                           V
                                      **AV Team**              **Room Captains**        **PyBay Logistics Chair**
                                  All Summary Sheets          All Summary Sheets           All Summary Sheets
                                                               + Detail Sheets              + Detail Sheets
                                                                 
```

## Setup & Usage

### 0. Verify or Install Python >=3.12

### 1. Clone this Repo

```
git clone https://github.com/pybay/conference-run-sheets.git
cd conference-run-sheets
```

### 2. Create a Virtual Environment & Install Project Dependencies

Of course, you do this automatically!

```
python3 -m venv venv
source venv/bin/activate
python3 -m pip install -r requirements.txt
```

### 3. Download Required Input Files

Assumes you are using Sessionize for the conference management of speakers and talks.  <br>
Download both:
- the `Schedule` Excel File (e.g. `pybay2025 schedulelist - exported 2025-10-12.xlsx`)
- the `Flattened Accepted Sessions` Excel file, which has all info on one tab, and easier to process (e.g. `pybay2025 flattened accepted sessions - exported 2025-10-13.xlsx`)


### 4. Review the Fields Actually Available in each Excel File

Sessionize provides both standard field names that cannot be edited (e.g., `Session ID`, `Title`, `Owner`, `Room`) and captures all speaker responses from the Call For Proposals (CFP), including any CUSTOM questions added by the PyBay organizers.

**Custom questions are highly valuable** because they let you collect conference-day information from speakers in advance during the CFP process. This information can be displayed directly on the Run Sheets, making event management much smoother.

**Useful custom fields to collect include:**
- First/last name pronunciation
- Preferred intro bullets (how they want to be introduced)
- Pronouns
- Whether this is their first conference talk
- Preferred contact method for day-of communications
- Any special AV or accessibility needs

> **Note:** The actual field names in Sessionize may differ from these examples. See [README-SESSIONIZE-FIELDS.md](README-SESSIONIZE-FIELDS.md) for the specific field names used in the PyBay 2025 Sessionize setup.

**⚠️ Important:** This script will fail or produce unexpected output if Sessionize field names change or are reworded. If you modify custom questions in Sessionize, you must update `main.py` to match the current schema before running the script.

### 5. Update the `main.py` script in this repo to match schema/columns available in current data<br>
If you have updated CUSTOM questions, that will change the field names; if you have removed custom questions, those must be removed from the scripts, if you have ADDED custom questions to Sessionize - AND you need the info for the Run Sheets - you must add them to the script.  If  you don't need the speaker responses for the Run Sheets, don't worry about it.<br><br>

### 6. Run `main.py` script from project root
It produces one output Excel file named `pybay_YYYY-run_sheets.YYYY.MM.DD.xlsx` at project root.  This file has multiple tabs for each use case (e.g. Summary Sheets, etc.).  Open and inspect them, correct any issues by either modifying the repo code (e.g. for system issues like missing/renamed columns).  If there are minor edits in content - it is likely coming from Sessionize Speaker input, recommend not worrying about it.  If there are major changes, e.g. a speaker cancels well in advance of the conference - update the conference schedule in Sessionize FIRST (so our public schedule on the website is current), THEN rerun this spreadsheet as needed. <br><br>

```
source ./venv/bin/activate
python -m main.py
```
### 7. Upload the generated Excel file to Google Sheets - to share with the team per diagram above.

