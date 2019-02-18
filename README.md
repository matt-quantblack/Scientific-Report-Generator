# Scientific-Report-Generator
Cloud based scientific report generator for comparative test data including ANOVA statistical testing. Uses Google drive for data storage, Google sheets for data entry and Microsoft Word for the report templates.

## Installation

Runs on Python 2.7 or greater

requires:
googleapiclient
google.oauth2
google.auth
statsmodels
scipy
re
statistics
docx

## Usage

- Run 'SRG.py start' to run the background service, ideally as a parallel process eg. linux 'SRG.py start &'
- 'SRG.py stop' will stop any background running process
- Create a google account for the report generating robot
- Create a ReportTemplate.docx and save in a team drive shared with the report robot account or share the file with report_robot account
- Create a google sheets document with a details page and each samples result on each tab. Save on team drive or share with report_robot Use SampleDataEntry.gsheet as an example format: https://docs.google.com/spreadsheets/d/1m13NCVqPrTx9qbB7icf6AsNZBsEnYkk42R_J94hqbx0/edit?usp=sharing
- When ready to build the report change filename to start with PROCESS 
- Change will be picked up, processed and report link will be emailed to the ShareWith email address
- Save the report in the desired folder


## License
[MIT](https://choosealicense.com/licenses/mit/)
