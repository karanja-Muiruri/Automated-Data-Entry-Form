# Excel VBA Dynamic Data Entry Form
## Table of Contents
  -  Project Overview
  -  Tools Used
  -  Project Workflow

## Project Overview
This project involves designing an Excel data entry dynamic form in VBA window using text boxes, option buttons, combo boxes, command buttons and list boxes. The data entered is saved in a database thus proving efficiency in data collection and safety.  

## Tools Used
Microsoft Excel (365) [Download here](https://microsoft.com)

## Project Workflow

#### Launch Form 
- Create a new Excel file and Save As 'Your prefferred name' and also ensure Save as type is 'Excel Macro-Enable Workbook(*.xlsm)
- Rename sheet1 as Home and on the View Tab, uncheck gridlines.
- Click Insert Tab, click shapes, and select Rounded rectangle.
- Draw the Rounded rectange on the middle of the sheet(Home).
- Select the drawn Rounded rectangle, go to Format Tab, select Shape Effects, choose Preset, and finally select Preset 4.
- Right click on the drawn Rounded Rectangle, select Edit Text, enter 'Launch Form'.
- Select the drawn Rounded rectangle, go to Home Tab, centre the 'Lauch Form' wording and change the Calibri(Body) text size to 40.
    - ![001](https://github.com/user-attachments/assets/27401198-4185-491e-8cc4-d013074f9f39)

- Add a new sheet on this same workbook and rename it(Sheet2) to Database.
- Add the following data as the Headers.
- Select a range(A1:I22), apply all borders, headers customization, and uncheck gridlines.
    - ![002](https://github.com/user-attachments/assets/5b2fbf1b-6c50-4e6d-ba32-360469317d01)

#### Designing Form in VBA Window
- Go to Developer Tab, click Visual Basic, resulting into a pop-up named Microsoft Visual Basic Applications(MVBA).
- On the MVBA pop-up, click Insert, select UserForm.
- On the left hand-side, is a Properties-UserForm1, locate Height, change from 180 to 325, locate Width, change from 240 to 578.
- Also on the Properties-UserForm1, locate (Name), change from UserForm1 to frmform.
- Change Caption as well to Automated Data Entry Form.   
    - ![003](https://github.com/user-attachments/assets/cc8aece9-ecbe-4bc6-97e8-8ada6ca93c8a)
- On the Toolbox, select Frame, and draw two frames in the MVBA pop-up.
- Change Caption of Frame1 to "Enter Details", change BorderStyle to "1- fmBorderStyleSingle', and choose BorderColor of your liking from palette.
- Change Caption of Frame2 to 'Database', change BorderStyle to "1- fmBorderStyleSingle', and choose BorderColor of your liking from palette.
    - ![004](https://github.com/user-attachments/assets/59e3c0ca-96e1-4ba6-b3cd-a74c656faedc)














