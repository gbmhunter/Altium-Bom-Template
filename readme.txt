Designed to work with Altium's automatic BOM creation feature and Altium's supplier linked components.

This repo contains a macro enabled BOM template file, with two buttons, 'Calculate Pricing', and 'Resize Rows'

HOW TO GENERATE BOM:
- Add this as a template file for a BOM
- Make sure the correct parameters are to be included in the BOM. This BOM macro requires supplier linked components. 
   The parameters are:
   - Designator
   - Description (supplier linked param)
   - Manufacturer (supplier linked param)
   - Manufacturer Part Number (supplier linked param)
   - Supplier 1 (supplier linked param)
   - Supplier Part Number 1 (supplier linked param)
   - Pricing 1 (supplier linked param, this one needs special consideration, see below)   
   - Quantity
- Add a project parameter called "Reference" and give it a recognisable project id as it's value
- Run Altium BOM generation
- Click the "Calculate Pricing" button and then the "Resize Rows" button in the generated BOM file (Excel spreadsheet).
- Done!

Pricing 1 Parameter Format:

<quantity1>=<price1>, <quantity1>=<price1>, ... (<currency>)
e.g.
"1=1.23, 10=0.97, 25=0.65 (USD)"
This is the same format as that generated when supplier paramters are added to components. Fields can be manually
changed/added as required.
Note that the currency has to be the same for all components (the macros do not perform currency conversions!).

WHEN EDITING MACRO TEMPLATE FILE
- Make sure to save the macros as part of the sheet, not the workbook.
- Do not rename the sheet (leave as "Sheet1")
- Always save with the extension ".xlt" (excel template file), as Altium requires this

The macros could of been designed to run as soon as Altium fills in the cells, but the advantages of running the macros on a user's button press is as follows:
1) If something goes wrong when the macros are running (e.g. null or out-of-bounds exception due to inforrect data), Altium will not lock-up
2) User can perfrom small BOM changes before running the macro to fix incomplete fields (although it is recommended to enter it correctly into Altium in the first place).