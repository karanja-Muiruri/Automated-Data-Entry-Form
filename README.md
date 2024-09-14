# Excel VBA Dynamic Data Entry Form
## Table of Contents
  -  Project Overview
  -  Tools Used
  -  Project Workflow

## Project Overview
This project involves designing an Excel data entry dynamic form in VBA window using text boxes, option buttons, combo boxes, command buttons and list boxes. The data entered is saved in a database thus proving efficiency in data collection and safety.  
    - ![007](https://github.com/user-attachments/assets/1d3b7859-12fc-43a5-b717-61ea3745499e)


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
- On the Toolbox, select Label, and create three labels on the Enter Details frame. Rename their Captions to Employee ID, Employee Name, Gender respectively.
- On the Toolbox, select TextBox, add TextBoxes adjacent to both Emp ID and Emp Name Labels, and rename the TextBoxes (Name) TextBox1 to txtID, txtName respectively.
- On the Toolbox, select OptionButton, create two OptionButtons adjacent to the Gender Label, and rename the OptionButtons (Name) from OptionButton1 to optMale, optFemale respectively. Also edit their Captions to Male and Female.
- Do the same for Department, City and Country. For Department, replace the TextBox with a ComboBox.
- Rename the Department ComboBox (Name) from ComboBox1 to cmbDepartment, City (name) to txtCity, Country (Name) to txtCountry.
- CommandButton1 (Name) to cmdSave, Accelarator set as 's', Backcolor 'Blue', Caption change to 'Save'.
- CommandButton2 (Name) to cmdReset, Accelarator set as 'r', Backcolor 'Red', Caption change to 'Reset'.
- Change the TabIndex, starting from Emp ID's TextBox as 1 to Reset CommandButton as 9.
- On the ToolBox, select CommandButton, create two Command Buttons just below the Country's Label and TextBox.
    - ![005](https://github.com/user-attachments/assets/003cae74-a774-4ae2-a03e-71d8f9715077)

- On the ToolBox, select ListBox, draw it on the Database Frame. Rename (Name) from ListBox1 to lstDatabase, Assign TabIndex as 10.
    - ![006](https://github.com/user-attachments/assets/f539e369-ee6c-42dd-b931-51386d5b7ec7)


#### VBA Code for Sub Procedures and Event Handling
- On the MVBA pop-up, Click Insert Tab and select Module.
- On the Module, using Sub Reset() to reset the form controls and initialize form;

        Sub Reset()
                Dim iRow As Long
                iRow = [Counta(Database!A:A)] # This code calculates the last row number in column A of the Database sheet where there is data.
                With frmform
                    .txtID.Value = ""         #  clears the values of textboxes (txtID, txtName) & unchecks the radio buttons (optMale, optFemale) on frmform.
                    .txtName.Value = ""
                    .optMale.Value = False
                    .optFemale.Value = False
                    
                    .cmbDepartment.Clear      # clears any existing items in the ComboBox cmbDepartment & then adds four new items to it: "HR", "Operation", "Training", & "Quality".
                    .cmbDepartment.AddItem "HR"
                    .cmbDepartment.AddItem "Operation"
                    .cmbDepartment.AddItem "Training"
                    .cmbDepartment.AddItem "Quality"
                    
                    .txtCity.Value = ""      # clears the values in the textboxes txtCity and txtCountry.
                    .txtCountry.Value = ""
                    
                    .lstDatabase.ColumnCount = 9    # sets up the lstDatabase ListBox with 9 columns, enables column headers, & specifies the widths for each column.
                    .lstDatabase.ColumnHeads = True
                    
                    .lstDatabase.ColumnWidths = "30,60,75,40,60,45,55,70,70"
                    
                    If iRow > 1 Then      # If there are more than 1 row of data (iRow > 1), it sets the RowSource to include rows from A2 to the last row I where data exists.
                        .lstDatabase.RowSource = "Database!A2:I" & iRow
                    Else
                        .lstDatabase.RowSource = "Database!A2:I2"
                        
                    End If
        
                End With
    
        End Sub  


- After the above code, on MVBA, click the Debug Tab, select CompileVBAProject.

#### Sub Submit(), The Submit subroutine is designed to gather data from a user form and write it into the next available row in a worksheet named "Database".


        Sub Submit()

              Dim sh As Worksheet      # declares two variables: sh for the worksheet object and iRow for the row number.
              Dim iRow As Long
              
              Set sh = ThisWorkbook.Sheets("Database")   # sets the variable sh to refer to the "Database" worksheet in the current workbook.
              
              iRow = [Counta(Database!A:A)] + 1    # calculates the next available row by counting the number of non-empty cells in column A of the "Database" sheet and adding 1 to it.
              
              With sh                              # writes data to the worksheet "Database" in the next available row:
                  
                  .Cells(iRow, 1) = iRow - 1       # Sets the first cell in the new row to the row number minus 1.
                  .Cells(iRow, 2) = frmform.txtID.Value      # Writes the value from the txtID textbox on the form to column 2.
                  .Cells(iRow, 3) = frmform.txtName.Value    
                  .Cells(iRow, 4) = IIf(frmform.optFemale.Value = True, "Female", "Male")  # Writes "Female" if the optFemale radio button is selected, otherwise writes "Male".
                  .Cells(iRow, 5) = frmform.cmbDepartment.Value
                  .Cells(iRow, 6) = frmform.cmbDepartment.Value
                  .Cells(iRow, 7) = frmform.txtCountry.Value
                  .Cells(iRow, 8) = Application.UserName
                  .Cells(iRow, 9) = [Text(Now(), "DD-MM-YYYY HH:MM:SS")]   # Writes the current date and time in the format "DD-MM-YYYY HH:MM" to column 9.
                  
              End With
              
        End Sub



    
        
        Sub Show_Form()

            frmform.Show      # This subroutine is used to open the user form so that users can interact with it. Typically, you would use this procedure to present a form that allows users to input data, make selections, or interact with other controls.
    
        End Sub


       
- ![008](https://github.com/user-attachments/assets/b63b6df6-df27-4b66-acb9-a3bcea3e23f7)


-  After the above code, on MVBA, click the Debug Tab, select CompileVBAProject.

-  On the top-left of the MVBA, under Forms Folder, locate frmform and click it, resulting into Automated Data Entry Form Version 1.0 popping up.
-  Double click on the Reset CommandButton, and write the following code;


          Private Sub cmdReset_Click()
    
                  Dim msgValue As VbMsgBoxResult                  # declares a variable msgValue of type VbMsgBoxResult, used to store the result of the message box that will be displayed to the user.
                  
                  msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")               # displays a message box to the user. If msgValue = vbNo Then Exit Sub   # checks if the user's response was "No" (vbNo). If it was, the Exit Sub statement terminates the cmdReset_Click subroutine early, meaning that the form will not be reset.
                  
                  Call Reset                                       # If the user chose "Yes", the Reset subroutine is called. This subroutine is responsible for clearing and setting up the form (as defined in the Reset subroutine you provided earlier).
                    
          End Sub



-  On the top-left of the MVBA, under Forms Folder, locate frmform and click it, resulting into Automated Data Entry Form Version 1.0 popping up.
-  Double click on the Save CommandButton, and write the following code;


          Private Sub cmdSave_Click()
    
                  Dim msgValue As VbMsgBoxResult
                  
                  msgValue = MsgBox("Do you want to save the data?", vbYesNo + vbInformation, "Confirmation")
                  
                  If msgValue = vbNo Then Exit Sub
                  
                  Call Submit
                  Call Reset
    
          End Sub


-  After the above code, on MVBA, click the Debug Tab, select CompileVBAProject.
-  On the top-left of the MVBA, under Forms Folder, locate frmform and click it, resulting into Automated Data Entry Form Version 1.0 popping up.

- Top right, on the Click drop-down, select Initialize and write the following code.


          Private Sub UserForm_Initialize()                      # declares the UserForm_Initialize event handler. This subroutine is automatically called by Excel VBA when the user form is initialized, which typically occurs when the form is first opened.
    
                   Call Reset                                    # Inside the UserForm_Initialize event handler, the Reset subroutine is called. The Reset subroutine is responsible for resetting the user form to its default state. This includes clearing input fields, setting default values, and configuring any controls on the form as described in the Reset subroutine you provided earlier.
    
          End Sub


-  After the above code, on MVBA, click the Debug Tab, select CompileVBAProject.


- Top left, below the File Tab, is a 'View Microsoft Excel Tab', click it. Right click the 'Launch Form', select Assign Macro, finally select Show Form, click OK.

  
