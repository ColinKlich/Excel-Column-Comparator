# Excel Column Comparator
This app compares two excel documents and uses a custom written "fuzzy matching" algorithm to compare a similar column in each document.
Matches are exported to a new document along with a similarity index from 0-1 (This feature is currently disabled in production).

![](columncompare.png)

## Makes use of these librarys:
- Pandas
- Jellyfish
- Difflib
- Tqdm
- Beaupy
- Pyinstaller

## Goals
- This program attempts to fix excel's built it fuzzy matching module
- The specific process that this program was built to address takes 6 hrs/month. 
    - Versions 1-3 bring that time down to 1 hr. 
    - Versison 4 and beyond should bring that time down to 5 minutes.
- This program is designed to be multi-functional and address similar processes with similar time-saving requirements
- Code is compiled into an .exe file to run on any windows machine.

## Upcoming Features
- Ability to select columns from source lists and output them to the 'output.xlsx' file (v4) DONE
- Exports 'Not Found' list along with 'Found' list (v6) DONE
- Version 7: FINAL - speed improved 300% compared to V6. 'not found' list is merged with input to show all original data.