PROCUREMENT APPROVAL AUTOMATION
1. Install Python 3.11+ on Windows and tick "Add Python to PATH".
2. Open Command Prompt in this folder.
3. Run: python -m pip install -r requirements.txt. The run_report.bat launcher also checks for and installs missing project dependencies.
4. Export your procurement data as CSV (with column headers) into input\.
5. Double-click run_report.bat or run: python procurement_report.py. If running the script directly and you see a missing pandas error, install dependencies with the command in step 3 from this folder.
6. Find the Excel workbook in output\.

Expected columns: a header containing "Requisition Date" and a header containing "Assigned Date".
Input date format is US MM-DD-YYYY or MM/DD/YYYY. Output display is DD-MM-YYYY.
For reproducible results, run: python procurement_report.py --as-of 2026-10-14
NOTE: Approval days are CALENDAR days, not business days. The report flags rows, not unique PRs.
The Raw CSV sheet preserves the original values for audit.
