## An Alert App for Canteen Queensland
---

### Overview
This program is designed to regularly send alert email to CanTeen Queenland staff.
The email contains a report on volunteers whose licenses going to expire in a certain
amount of time. The program also checks if a volunteer has been inactive for more than 2 years.

### Workflow
- Create log file(.txt)
- Create report file(.xlsx)
- Read volunteer list from file(.csv)
- Loop through the list and write the report
- Send the email with the report attached

### Volunteer Data
The program reads volunteer data from a .csv file and examines each volunteer's license
status, as well as his/her last activity date.

The data is extracted from the organisation's SQL database, and the .csv file is automatically generated every 2 weeks.

### Log
Upon execution, the program creates a log file(.txt) and stores it in the directory ./Log.
If the program doesn't function as expected, the log can be useful for debuggin.

### Deployment
The application is deployed on the organisation's server and set to run every 2 weeks after the .csv file is generated.

### Files Required
Three items are required before running the program, namely qld_old_members.csv, as well as mail.html and MailList.txt under the application's directory.

### Acknowledgement
All sensitive information has been taken off the code.
The application serves for a specific task and does not apply to any general usage.
This repository is for code demonstration only.