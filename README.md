# Automatic Mass-Invitation Program
Dynamic mass-invitation program with multi-event assignments. Ideal for routine onboarding. Invitation button generates a sheet by bundling all assignments and their respective email addresses and goes through invitation process using calendar patch. Invitation process is optimized, limiting API calls per event ID.

**The Settings Sheet** handles any amount of labels/assignments. Below each label, Google Calendar invite URLs should be listed.

**The Main Sheet** can include any data, with the labels in a specified column (C in the preview below) and the email addresses in another specified column (L and M in the preview below). When the program is complete, all successfully invited rows will have a checkmark, and irrelevant rows are skipped.

**The Sidebar** is flexibly written, allowing the user to specify the columns for the labels and email(s). If more than one email column is referenced, separate by column letter and comma, e.g. "L, M". The program iterates through all rows by default, but a _selected range_ feature is provided in the dropdown.


*Note: Program and formulas are set to start on row 6 in the Main Sheet. Labels are set for row 5 and invite URLs are set for row 6 in the Settings Sheet.*

![Main Sheet Preview](https://raw.githubusercontent.com/sajadmh/Automatic-Mass-Invitation-Program/main/Main%20Sheet%20Preview.png)

![Settings Sheet Preview](https://raw.githubusercontent.com/sajadmh/Automatic-Mass-Invitation-Program/main/Settings%20Preview.png)
