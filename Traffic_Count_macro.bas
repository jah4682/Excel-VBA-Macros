Attribute VB_Name = "Traffic_Count"
Sub TrafficCount()

'###Declare Variables###

'Variables for looping through excel files
Dim MyFolder As String, MyFile As String, StationName As Variant

'Variables for identifying which column to insert traffic counts
Dim MyColumnMaster As Variant, MyColumnListA As Variant, MyColumnListB As Variant, MyColumnListC As Variant, LastColumn As Variant

'traffic count number
Dim MyTrafficCount As Variant

'store the Station number cell's location
Dim rgStationLoc As Range, StationLoc As Variant

'to build new cell address
Dim Row As Variant, NewCellAddress As String



'###Input Dialog Box###
'Ask user which column they want the traffic count to be inserted into
'MyColumnMaster = InputBox("Enter the Column letter you want to input the traffic counts for sheet ""Master-All Stations""", "traffic counts", "Enter Column Letter")
'MyColumnListA = InputBox("Enter the Column letter you want to input the traffic counts for sheet ""List A Every Year""", "traffic counts", "Enter Column Letter")
'MyColumnListB = InputBox("Enter the Column letter you want to input the traffic counts for sheet ""List B Even Years""", "traffic counts", "Enter Column Letter")
'MyColumnListC = InputBox("Enter the Column letter you want to input the traffic counts for sheet ""List C Odd Years""", "traffic counts", "Enter Column Letter")



'###Opens a file dialog box for user to select a folder###
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .Show
    MyFolder = .SelectedItems(1)
    Err.Clear
End With



'This section will loop through and open each file in the selected folder
'and then close that file before opening the next file
 MyFile = Dir(MyFolder & "\", vbReadOnly)   '8001.xls\
 
 '###Loop###
 Do While MyFile <> ""
     Workbooks.Open Filename:=MyFolder & "\" & MyFile, UpdateLinks:=False
     ''''''''''''ENTER YOUR CODE HERE TO DO SOMETHING'''''''''
     
     'test where traffic count is located
     If IsEmpty(Workbooks(MyFile).Worksheets(1).Range("D106").Value) Then
        'Copy traffic count in cell B103 in the raw traffic count excel file
        MyTrafficCount = Workbooks(MyFile).Worksheets(1).Range("B103").Value
     Else
        'Copy traffic count in cell D106 in the raw traffic count excel file
        MyTrafficCount = Workbooks(MyFile).Worksheets(1).Range("D106").Value
     End If
     
     'get the four left digits from the file name and assign to StationName ex: 8001
     'this will be the value we will be searching for
     StationName = Left(MyFile, 4)
     
     
     'find corresponding station in List A, B, or C
     
     '''List A - Every Year Counts'''
     If Not Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List A - Every Year Counts").Range("B2:B201").Find(StationName) Is Nothing Then

         'get station cell address and assign to variable
         Set rgStationLoc = Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List A - Every Year Counts").Range("B2:B201").Find(StationName)
         
         'Alternate way to find last column
         With Workbooks("MASTER LIST_copy_macro").Worksheets("List A - Every Year Counts")
             LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
         End With

         'get station's cell row. convert to string. trim leading spaces
         'Row = LTrim(Str(rgStationLoc.Row))
         'Alternate
         Row = rgStationLoc.Row

         'build new cell location to insert traffic count into
         'NewCellAddress = "$" + MyColumnListA + "$" + Row
         'Alternate
         NewCellAddress = Cells(Row, LastColumn).Address

         'Paste copied traffic count into cell location
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List A - Every Year Counts").Range(NewCellAddress).Value = MyTrafficCount

         'format number
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List A - Every Year Counts").Range(NewCellAddress).NumberFormat = "#,##"

     '''List B - Even Years'''
     ElseIf Not Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List B - Even Years").Range("B2:B201").Find(StationName) Is Nothing Then

        'get station cell address and assign to variable
         Set rgStationLoc = Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List B - Even Years").Range("B2:B201").Find(StationName)
         
         'Alternate way to find last column
         With Workbooks("MASTER LIST_copy_macro").Worksheets("List B - Even Years")
             LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
         End With

         'get station's row. convert to string. trim leading spaces
         'Row = LTrim(Str(rgStationLoc.Row))
         'Alternate
         Row = rgStationLoc.Row

         'build new cell location to insert traffic count into
         'NewCellAddress = "$" + MyColumnListB + "$" + Row
         'Alternate
         NewCellAddress = Cells(Row, LastColumn).Address

         'Paste copied traffic count into cell location
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List B - Even Years").Range(NewCellAddress).Value = MyTrafficCount

         'format number
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List B - Even Years").Range(NewCellAddress).NumberFormat = "#,##"

     '''List C - Odd Years'''
     ElseIf Not Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List C - Odd Years").Range("B2:B191").Find(StationName) Is Nothing Then

         'get station cell address and assign to variable
         Set rgStationLoc = Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List C - Odd Years").Range("B2:B191").Find(StationName)
         
         'Alternate way to find last column
         With Workbooks("MASTER LIST_copy_macro").Worksheets("List C - Odd Years")
             LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
         End With

         'get station's row. convert to string. trim leading spaces
         'Row = LTrim(Str(rgStationLoc.Row))
         'Alternate
         Row = rgStationLoc.Row

         'build new cell location to insert traffic count into
         'NewCellAddress = "$" + MyColumnListC + "$" + Row
         'Alternate
         NewCellAddress = Cells(Row, LastColumn).Address

         'Paste copied traffic count into cell location
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List C - Odd Years").Range(NewCellAddress).Value = MyTrafficCount

         'format number
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("List C - Odd Years").Range(NewCellAddress).NumberFormat = "#,##"

     Else
        MsgBox "Station " + StationName + " was not found"
     End If
    
    
     '''Master-All Sheet'''
     'find corresponding station in Master-All Sheet
     With Workbooks("MASTER LIST_copy_macro").Worksheets("Master-All Stations").Range("B2:B591")
        'get station cell address and assign to variable
        Set rgStationLoc = .Find(StationName)
     End With
         
    'Alternate way to find last column
    With Workbooks("MASTER LIST_copy_macro").Worksheets("Master-All Stations")
         LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    
         'get station's row. convert to string. trim leading spaces
         'Row = LTrim(Str(rgStationLoc.Row))
         Row = rgStationLoc.Row
         
         'build new cell location to insert traffic count into
         'NewCellAddress = "$" + MyColumnMaster + "$" + Row
         
         'Alternate way to build new cell address
         NewCellAddress = Cells(Row, LastColumn).Address
     
         'Paste copied traffic count into cell location
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("Master-All Stations").Range(NewCellAddress).Value = MyTrafficCount
         
         'format number
         Workbooks("MASTER LIST_copy_macro.xlsm").Worksheets("Master-All Stations").Range(NewCellAddress).NumberFormat = "#,##"

     
     Workbooks(MyFile).Close SaveChanges:=False
     MyFile = Dir
     
 Loop



End Sub
