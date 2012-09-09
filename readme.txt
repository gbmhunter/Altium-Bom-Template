Designed to work with Altium's BOM creation feature.

Two features added, 'Calculate Pricing', and 'Resize Rows'

WHEN EDITING MACRO TEMPLATE FILE
- Make sure to save the macros as part of the sheet, not the workbook.
- Do not rename the sheet (leave as "Sheet1")
- Always save with the extension ".xlt" (excel template file), as Altium requires this

The macros could of been designed to run as soon as Altium fills in the cells, but the advantages of running the macros on a user's button press is as follows:
1) If something goes wrong when the macros are running (e.g. null or out-of-bounds exception due to inforrect data), Altium will not lock-up
2) User can perfrom small BOM changes before running the macro to fix incomplete fields (although it is recommended to enter it correctly into Altium in the first place).