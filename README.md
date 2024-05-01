# Supply 2.0

__Supply 2.0__ is a fully functional tool designed to serve the uniform distribution needs for all air cadet squadrons in Canada. It is designed to run on the excel app on the desktop. The below document will briefly describe its functionalities.

**Initiation**: after the excel sheet is downloaded onto the desktop. You must right click on the files and click properties and then allow macros. (More info) Both Files also must be placed into a folder named “Supply Excel”.

## Components
**Menu**: houses a list of all cadets, their uniform status, and buttons, which are linked to individual cadet sheets(seen below) through hyperlinks. The buttons are:
- Update status: goes through each sheet to update status of cadets
- New Cadet Uniform Form: Generates a new excel sheet with all necessary uniform parts based of body measurements automatically through algorithms
- Old Cadet exchange: Meant for cadets with no file in excel. Creates new file to find uniforms to exchange
- Manual Backup: creates backup on the file on the desktop in a folder called supply 2.0. Also triggered automatically when the document is closed.

**Pickup Sheet**: creates a list of all uniforms with the status of ready to be picked up. This is especially helpful for uniform distribution day.
- Generate button: refreshes table for all uniform parts ready to pick up
- Complete button: update all uniform statuses to ready to complete.

**Mastersheet**: overview of all uniform parts and their statuses. This page is designed with filtering and sorting in mind. Uniforms can be sorted from small to large and you can filter the names to only look for cadets who need uniforms.
- Generate button: refreshes table to recreate uniform logs
- Toggle button: triggers status change in the name cell. Different colors reflect whether uniforms are in process or complete.

**Individual Cadet Sheet**: Features personal information of each cadet, body measurements, their uniforms parts and their uniform status. 
- Exchange Item Button: click on the uniform parts that you need to exchange and toggle the button. Enter the new body measurements that turn blue and press submit. 
- Change Inventory for selected NSN: changes inventory in inventory file once you select the NSN by selecting the box. 
- Update In Stock Status: check if UNP(unprocessed items) are in stock
- Mark as SOS: puts sos on the file to show that it is SOS in the menu
- SOS button: deleted excel sheet and corresponding row in menu

**Note For Statuses**: UNP = unprocessed, pick up = uniforms packaged and ready to pick up


## Important Notes:
- Import sheet is used for updates in the system. Ignore for now
- Do not tamper with the template sheet, unless to change squadron name
- All saves of files will be stored on a folder on the desktop named “Supply 2.0”




