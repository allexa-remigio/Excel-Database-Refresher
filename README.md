# Master Excel Dashboard Refresher (MEDBR)
Master Database Dashboards

## Authors
Allexa Remigio

## Status
Completed

## Description
Given multiple, daily-updated Excel Worksheets containing various patient databases on Microsoft Excel (online version), I was tasked with creating an auto-updated master patient dashboard to compile all the important patient information required for the administrative team to schedule patients. In order to help the team further, the worksheet that houses the master patient worksheet is also able to be modified by the staff, so this script also features a way to keep track of all the edits that the administrative team makes for the external database management.

Since each database is formatted differently for internal reasons, the MEDBR only extracts the essential patient data required for scheduling purposes:
-Date Referred
-Initial Visit Date
-Patient Demographics (Name, Date of Birth, MRN, Location, Zip Code)
-Assigned Doctor/Nurse Practitioner (or otherwise Medical Provider)
-Last Visit Date
-Anticipated Next Visit Date
-Additional Notes

After the information is extracted, it is compiled into a mastersheet of all patients seen by one of our medical providers.

To keep up with all the changes in the separete patient databases, this script is run daily once all the previous appointment information has been udpated via Power Automate. 

## Location of File
The MEDBR file works in the same folder as the worksheets containing the patient databases. 

## Usage/Options
Upload the script to Excel Automate. Automate > Script > Run Script
`run main function`

Optional: Create a button to run it straight from the worksheet.

## Output:

Worksheet | Definition 
 ------------ | ------------- 
Changes | the worksheet that keeps track of all the edits the administrative staff creates
Master | the worksheet that houses the compiled patient information
