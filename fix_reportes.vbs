Option Explicit 
'************************************************************* 
' You need to set a reference to the Microsoft Excel Object 
' Library to use these Excel Objects. 
'************************************************************* 
' Private xlApp As Excel.Application  ' Excel Application Object 
' Private xlBook As Excel.Workbook    ' Excel Workbook Object 

'************************************************************* 
' Gets the contents of an Excel Worksheet's cell. 
' 
' xlWorksheet: Name of a worksheet in an Excel File, for example, 
'              "Sheet1" 
' xlCellName:  Name of a Cell (Row and Column), for example, 
'              "A1" or "B222". 
' xlFileName:  Name of an Excel File, for example, "C:TestTesting.xls" 
'************************************************************* 
REM the Excel Application
Dim xlApp_in
REM the path to the excel file to be read
Dim xlFileName_in
REM the File System Object
Dim objFSO
REM the Start Folder Object
Dim objStartFolder
REM the Folder Object
Dim objFolder
REM the Files Collection
Dim colFiles
REM the File Object
Dim objFile


' Create the Excel App Objects
Set xlApp_in = CreateObject("Excel.Application")

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\Users\jose.ibarra\Documents\Josue\Trabajo\TopaDengue\Reportes\test\"
Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For each objFile in colFiles
   xlFileName_in = objStartFolder & objFile.Name
   ' Wscript.Echo "Parsing " & xlFileName_in & " ..." & vbCRLF
  
   Replace_Manzanas objStartFolder, objFSO, xlApp_in, xlFileName_in
Next

xlApp_in.Quit
Set xlApp_in = Nothing
WScript.Echo "Final Success!"

Private Sub Replace_Manzanas(objStartFolder, objFSO, xlApp_in, xlFileName_in):
   REM the Excel Book
   Dim xlBook_in
   REM The New File Name
   Dim xlFileName_in_new
   REM the Worsheet Name
   Dim xlWorksheet_in_name
   REM the coordinates of the Cell to edit
   Dim xlCellName_in
   REM the Contents to set the Cell
   Dim xlCellContents

   REM how many worksheets are in the current excel file
   Dim inWorkSheetCount
   Dim counter
   REM the worksheet we are currently getting data from
   Dim currentWorksheet_in

   REM variables that store the relevant data of the csv
   Dim csv_cod
   Dim csv_cod_new
   Dim manzana_cod
   Dim manzana_cod_new
   Dim num_flia_orig
   Dim num_flia_new

   REM aux variables
   Dim sustraendo

   ' Script.Echo "Entered Replace_Manzanas" & vbCRLF
   ' Create the Excel Workbook Object. 
   Set xlBook_in = xlApp_in.Workbooks.Open(xlFileName_in)

   REM How many worksheets are in this Excel documents
   inWorkSheetCount = xlApp_in.Worksheets.Count

   REM Loop through each worksheet
   For counter = 1 to inWorkSheetCount
      Set currentWorksheet_in = xlApp_in.ActiveWorkbook.Worksheets(counter)
      xlWorksheet_in_name = currentWorksheet_in.Name

      ' WScript.Echo "-----------------------------------------------"
      ' WScript.Echo "Reading data from worksheet No. " & counter & ": " & xlWorksheet_in_name & vbCRLF


      If xlWorksheet_in_name = "Sheet1" Then
         csv_cod = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, "B1")
         csv_cod = Split(csv_cod, "-")
         manzana_cod = csv_cod(0)
         manzana_cod_new = manzana_cod
         If manzana_cod = "25005" Then
            manzana_cod_new = "15005"
         End If
         If manzana_cod = "25450" Then
            manzana_cod_new = "15450"
         End If
         If manzana_cod = "15880(ilegible)" Then
            manzana_cod_new = "15420"
         End If
         If manzana_cod = "15430" Then
            manzana_cod_new = "15420"
         End If
         If manzana_cod = "15760" Then
            manzana_cod_new = "15670"
         End If
         csv_cod(0) = manzana_cod_new
         csv_cod_new = Join(csv_cod,"-")

         If manzana_cod = "29050" Then
            num_flia_orig = csv_cod(1)
            num_flia_orig = CInt(num_flia_orig)
            manzana_cod_new = csv_cod(0)
            sustraendo = 0
            If num_flia_orig >= 5 and num_flia_orig <= 8 Then
               sustraendo = 4
               manzana_cod_new = "29130"
            End If
            If num_flia_orig >= 9 and num_flia_orig <= 12 Then
               sustraendo = 8
               manzana_cod_new = "29030"
            End If
            If num_flia_orig >= 13 and num_flia_orig <= 16 Then
               sustraendo = 12
               manzana_cod_new = "29000"
            End If
            num_flia_new = num_flia_orig - sustraendo
            csv_cod(0) = manzana_cod_new
            csv_cod(1) = num_flia_new
            csv_cod_new = Join(csv_cod,"-")

         End If

         If manzana_cod = "15670" Then
            num_flia_orig = csv_cod(1)
            num_flia_orig = CInt(num_flia_orig)
            manzana_cod_new = csv_cod(0)
            sustraendo = 0
            If num_flia_orig >= 5 and num_flia_orig <= 6 Then
               sustraendo = 4
               manzana_cod_new = "15540"
            End If
            If num_flia_orig >= 7 and num_flia_orig <= 10 Then
               sustraendo = 2
               manzana_cod_new = manzana_cod
            End If
            num_flia_new = num_flia_orig - sustraendo
            csv_cod(0) = manzana_cod_new
            csv_cod(1) = num_flia_new
            csv_cod_new = Join(csv_cod,"-")

         End If

         SetExcel xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, "B1", csv_cod_new
      End If
      Set currentWorksheet_in = Nothing 'Copy'
   Next
   ' Save changes
   xlBook_in.Save
   ' Save changes and close the spreadsheet 
   ' xlBook_in.Save

   ' Close the spreadsheet
   xlBook_in.Close(false)
   Set xlBook_in = Nothing 

   xlFileName_in_new = objStartFolder & csv_cod_new & ".xlsm"
   ' WScript.Echo "Renaming " & xlFileName_in & " to " & xlFileName_in_new & vbCRLF
   objFSO.MoveFile xlFileName_in, xlFileName_in_new
End Sub

Private Function GetExcel(xlApp, xlBook, xlFileName, xlWorksheet, xlCellName) 
   Dim strCellContents 
   
   ' Get the Cell Contents 
   strCellContents =     xlBook.worksheets(xlWorksheet).range(xlCellName).Value
   
   GetExcel = strCellContents
   
   Exit Function 

End Function 

'************************************************************* 
' Sets the contents of an Excel Worksheet's cell. 
' 
' xlWorksheet: Name of a worksheet in an Excel File, for example, 
'              "Sheet1" 
' xlCellName:  Name of a Cell (Row and Column), for example, 
'              "A1" or "B222". 
' xlFileName:  Name of an Excel File, for example, "C:TestTesting.xls" 
' xlCellContents:  What you want to place into the Cell. 
'************************************************************* 
Private Sub SetExcel(xlApp, xlBook, xlFileName, xlWorksheet, xlCellName, xlCellContents)    
   ' Set the value of the Cell 
   xlBook.worksheets(xlWorksheet).range(xlCellName).Value = xlCellContents 

   Exit Sub
End Sub