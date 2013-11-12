==============================================================
Altium BOM Template
==============================================================

- Author: gbmhunter <gbmhunter@gmail.com> (http://www.cladlab.com)
- Created: 2012/06/12
- Last Modified: 2013/11/12
- Version: v3.0.1.0
- Company: CladLabs
- Project: Free Code Libraries	.
- Language: Excel
- Compiler: n/a
- uC Model: n/a
- Computer Architecture: n/a
- Operating System: n/a
- Documentation Format: n/a
- License: GPLv3

Description
===========

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
- Leave all columns unhiden. If columns are hidden, Atlium will not auto-fill them with component parameters.

The macros could of been designed to run as soon as Altium fills in the cells, but the advantages of running the macros on a user's button press is as follows:
1) If something goes wrong when the macros are running (e.g. null or out-of-bounds exception due to inforrect data), Altium will not lock-up
2) User can perfrom small BOM changes before running the macro to fix incomplete fields (although it is recommended to enter it correctly into Altium in the first place).

External Dependencies
=====================

Requires Altium to be of any use.

Issues
======

See GitHub Issues.
	
Changelog
=========

========= ========== ============================================================================================================
Version   Date       Comment
========= ========== ============================================================================================================
v3.0.1.0  2013/11/12 Fixed issue in where Supplier Sub-total price was taking into account the production quantity, where it should really be only the cost per PCB (quantity of a component on a single PCB*price of that component).
v3.0.0.0  2013/11/12 Renamed 'Altium BOM Template.xlt' to 'Altium BOM Template v1.0.xlt' and added 'Altium BOM Template v2.0 (vault compatible).xlt' which is compatible with vault-stored supplier information. v2.0 only has basic functionality, need to fix supplier sub-totals and add support for partial components.
v2.0.1.0  2013/11/12 Deleted .hgignore file left over from Mercurial repo.
v2.0.0.0  2013/11/08 Converted Mercurial repo to Git repo. Updated readme.txt to README.txt and formatted to match other project README's. Added changelog information, taken from commit comments.
v1.8.0.0  2012/11/18 Added ability to hide/unhide supplier 1 details. Moved quantity column to before supplier info.
v1.7.1.0  2012/11/18 Unhid all columns, as Altium won't auto-fill hidden columns.
v1.7.0.0  2012/11/18 Added resize function apply to all columns (except pricing).
v1.6.0.0  2012/11/18 Added shading to Assembly columns. Hid assembly parameters by default.
v1.5.0.0  2012/11/14 Reverse Hide/Unhide button text to reflect first operation. Hid Supplier 2 component parameters by default.
v1.4.0.0  2012/11/14 Added button to hide/unhide Supplier 2 component information.
v1.3.4.1  2012/11/13 Updated readme.txt.
v1.3.4.0  2012/11/13 Fixed bug: Extra row at end of components was being shaded.
v1.3.3.0  2012/11/13 Fixed bug: Last column was not being shaded. Added CONST for num of columns.
v1.3.2.0  2012/11/13 Reference field added as first column. Uses project-wide parameter 'Reference'.
v1.3.1.0  2012/11/13 Merged two cells so that the 'Bill of Materials' text didn't prevent the first column from resizing (smaller) when the auto-resize command was used.
v1.3.0.0  2012/09/09 Component rows are now shaded when 'Resize Rows' is pressed.
v1.2.1.0  2012/09/09 Removed test.txt.
v1.2.0.0  2012/06/12 Added BOM template file with macro code.
v1.1.0.0  2012/06/12 Test with proper username.
v1.0.0.0  2012/06/12 Initial commit.
========= ========== ============================================================================================================