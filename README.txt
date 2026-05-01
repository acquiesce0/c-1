============================================================
  BRINE DENSITY LOOKUP - INSTALLATION & USER GUIDE
============================================================

This guide walks you through installing the app on a Windows
computer that has nothing installed yet. Follow the steps in
order. It takes about 10 minutes.

What this app does:
  You type a material name (for example "C-3") and a density
  (for example "1.295"). The app looks through every Excel
  spreadsheet in the "years" folder and shows every matching
  record (within +/- 0.0005 of the density), grouped by year.


------------------------------------------------------------
  STEP 1 - INSTALL PYTHON (one-time setup)
------------------------------------------------------------

1.1  Open a web browser and go to:
        https://www.python.org/downloads/

1.2  Click the big yellow button that says
        "Download Python 3.x.x"
     (3.x.x will be the latest version - that is fine.)

1.3  When the file finishes downloading, double-click it to
     open the installer.

1.4  *** VERY IMPORTANT ***
     At the bottom of the first installer screen, you will
     see a checkbox that says:

        [ ] Add python.exe to PATH

     TICK THAT BOX so it looks like:

        [x] Add python.exe to PATH

     If you skip this, the app will not work.

1.5  Click "Install Now" and wait for it to finish (about
     2 minutes). When it says "Setup was successful",
     click "Close".


------------------------------------------------------------
  STEP 2 - INSTALL THE EXCEL LIBRARY (one-time setup)
------------------------------------------------------------

The app needs a small helper library called "openpyxl" to
read Excel files.

2.1  Press the Windows key + R together. A small "Run"
     window opens.

2.2  Type:
        cmd
     and press Enter. A black Command Prompt window opens.

2.3  In the black window, type exactly:
        pip install openpyxl
     and press Enter.

2.4  Wait. You will see lines of text scrolling. When it
     stops and shows a new prompt, the install is done.
     You can close the black window.


------------------------------------------------------------
  STEP 3 - COPY THE APP FOLDER TO THE COMPUTER
------------------------------------------------------------

3.1  Copy the entire "materl" folder (the one that contains
     this README.txt, app.py, run.bat, and the "years"
     folder) onto the computer. The Desktop is a good place.

     The final layout should look like this:

        materl\
           README.txt
           app.py
           run.bat
           years\
              2023\   (Excel files inside)
              2024\   (Excel files inside)
              2025\   (Excel files inside)
              2026\   (Excel files inside)


------------------------------------------------------------
  STEP 4 - START THE APP
------------------------------------------------------------

4.1  Open the "materl" folder.

4.2  Double-click the file named:
        run.bat

4.3  A window titled "Brine Density Lookup" opens. The
     status bar at the top says "Loading data..." for a
     few seconds, then changes to something like:
        "Loaded 5071 records from years: 2023, 2024, 2025, 2026."

     The app is ready.


------------------------------------------------------------
  STEP 5 - HOW TO USE IT
------------------------------------------------------------

5.1  In the "Material name" box, type a location name.
     Examples:
        C-3        SP-1        R.B        PC-2/C-3

     The match is case-insensitive and treats "/" as a
     separator. Typing "C-3" finds both "C-3" and "PC-2/C-3".

5.2  In the "Density" box, type a number with a decimal
     point. Example:
        1.295

5.3  Press the Enter key, or click "Search".

5.4  The table fills with every matching record, grouped
     by year. Each row shows:
        Year, Month, Day, Material, Density,
        Mg+2, Ca+2, K+ A.A., Na+ A.A., Source File

     The status bar at the bottom shows totals like:
        "Found 16 match(es) - 2023: 2, 2024: 9, 2025: 5"


------------------------------------------------------------
  STEP 6 - ADDING NEW DATA LATER
------------------------------------------------------------

6.1  To add a new month: drop the new spreadsheet into the
     correct year folder inside "years\". The file name must
     follow the pattern:
        Pan MM-YYYY.xlsx
     For example: Pan 05-2026.xlsx

6.2  To add a brand new year: create a new folder under
     "years\" named after the year (for example "years\2027"),
     then put the spreadsheets inside it.

6.3  IMPORTANT: the spreadsheet layout must match the 2025
     files (each tab = one day, location names in the header
     row, "Density" / "Mg+2" / "Ca+2" / "K+ A.A." / "Na+ A.A."
     in column A).

6.4  In the app, click the "Reload data" button at the top.
     The new files will be picked up.


------------------------------------------------------------
  TROUBLESHOOTING
------------------------------------------------------------

PROBLEM: When I double-click run.bat, a black window flashes
         open and closes immediately.
FIX:     Python is not installed, or "Add python.exe to PATH"
         was not ticked during install. Repeat Step 1 and
         make sure to tick that box.

PROBLEM: The black window stays open and shows
         "ModuleNotFoundError: No module named 'openpyxl'".
FIX:     Repeat Step 2 (pip install openpyxl).

PROBLEM: The app opens but says "No 'years' folder found".
FIX:     Make sure the "years" folder is in the same folder
         as app.py and run.bat. See the layout in Step 3.

PROBLEM: The app opens but a spreadsheet's data is missing.
FIX:     The file name probably does not match the pattern
         "Pan MM-YYYY.xlsx". Rename it (for example rename
         "January 2027.xlsx" to "Pan 01-2027.xlsx") and click
         "Reload data".
