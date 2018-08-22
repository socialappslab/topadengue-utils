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
REM the Excel Application
Dim xlApp_out
REM the Excel Book
Dim xlBook_out
REM the path to the excel file to be read
Dim xlFileName_in
REM the path to the excel file to be write
Dim xlFileName_out
REM the worksheet we are currently getting data from
Dim currentWorksheet_out
REM the Worsheet Name
Dim xlWorksheet_out_name
REM the number of columns in the current worksheet that have data in them
Dim usedColumnsOutCount
REM the number of rows in the current worksheet that have data in them
Dim usedRowsOutCount
REM the topmost row in the current worksheet that has data in it
Dim top
REM the leftmost row in the current worksheet that has data in it
Dim left
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
Set xlApp_out = CreateObject("Excel.Application")

xlFileName_out = "C:\Users\jose.ibarra\Documents\Josue\Trabajo\TopaDengue\Reportes\test1.xlsx"
' Create the Excel Workbook Object. 
Set xlBook_out = xlApp_out.Workbooks.Open(xlFileName_out)

Set currentWorksheet_out = xlApp_out.ActiveWorkbook.Worksheets(1)
xlWorksheet_out_name = currentWorksheet_out.Name
' WScript.Echo "Out-Worksheet name: " & xlWorksheet_out_name & vbCRLF

REM What is the topmost row in the spreadsheet that has data in it
top = currentWorksheet_out.UsedRange.Row
REM What is the leftmost column in the spreadsheet that has data in it
left = currentWorksheet_out.UsedRange.Column
REM how many rows are used in the current worksheet
usedRowsOutCount = currentWorksheet_out.UsedRange.Rows.Count
REM how many columns are used in the current worksheet
usedColumnsOutCount = currentWorksheet_out.UsedRange.Columns.Count

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\Users\jose.ibarra\Documents\Josue\Trabajo\TopaDengue\Reportes\test\"
Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For each objFile in colFiles
   xlFileName_in = objStartFolder & objFile.Name
   Wscript.Echo "Parsing " & xlFileName_in & " ..." & vbCRLF

   REM What is the topmost row in the spreadsheet that has data in it
   top = currentWorksheet_out.UsedRange.Row
   REM What is the leftmost column in the spreadsheet that has data in it
   left = currentWorksheet_out.UsedRange.Column
   REM how many rows are used in the current worksheet
   usedRowsOutCount = currentWorksheet_out.UsedRange.Rows.Count
   REM how many columns are used in the current worksheet
   usedColumnsOutCount = currentWorksheet_out.UsedRange.Columns.Count
   
   WriteUnifiedData xlApp_in, xlApp_out, xlBook_out, xlFileName_in, xlFileName_out, currentWorksheet_out, xlWorksheet_out_name, top, left, usedRowsOutCount, usedColumnsOutCount
Next

' Save changes and close the spreadsheet 
xlBook_out.Save
' Close the spreadsheet
xlBook_out.Close
Set xlBook_out = Nothing 
xlApp_in.Quit
xlApp_out.Quit
Set xlApp_in = Nothing
Set xlApp_out = Nothing
WScript.Echo "Final Success!"

Private Sub WriteUnifiedData(xlApp_in, xlApp_out, xlBook_out, xlFileName_in, xlFileName_out, currentWorksheet_out, xlWorksheet_out_name, top, left, usedRowsOutCount, usedColumnsOutCount):
   REM the Excel Book
   Dim xlBook_in
   REM the Worsheet Name
   Dim xlWorksheet_in_name
   REM the coordinates of the Cell to edit
   Dim xlCellName_in
   Dim xlCellName_out
   REM the Contents to set the Cell
   Dim xlCellContents

   REM how many worksheets are in the current excel file
   Dim inWorkSheetCount
   Dim counter
   REM the worksheet we are currently getting data from
   Dim currentWorksheet_in


   Dim row_in
   Dim row_in_string
   Dim column
   Dim Cells
   REM the number of the first data-free row
   Dim first_unused_row_out
   REM the current row and column of the current worksheet we are reading
   Dim curCol_in
   Dim curRow_in
   Dim curCol_out
   Dim curRow_out
   Dim curRow_out_string
   Dim curCol_tmp
   REM the value of the current row and column of the current worksheet we are reading
   Dim word

   REM variables that store the relevant data of the csv
   Dim fecha_visita
   Dim csv_cod
   Dim manzana_cod
   Dim casa_cod
   Dim flia_jefe
   Dim tipo_criadero_dc 'Tipo de criadero segun los codigos de DengueChat
   Dim abatizado
   Dim larvas_bool 'Larvas? Si/No
   Dim tipo_criadero_senepa 'Tipo de criadero segun los codigos del SENEPA
   Dim cant_cont 'Cantidad de contenedores

   ' Script.Echo "Entered WriteUnifiedData" & vbCRLF
   ' Create the Excel Workbook Object. 
   Set xlBook_in = xlApp_in.Workbooks.Open(xlFileName_in)

   REM How many worksheets are in this Excel documents
   inWorkSheetCount = xlApp_in.Worksheets.Count

   first_unused_row_out = top + usedRowsOutCount
   curRow_out = first_unused_row_out
   curRow_out_string = CStr(curRow_out)
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
         casa_cod = csv_cod(1)
         flia_jefe = csv_cod(2)
         ' WScript.Echo "Codigo de Manzana: " & manzana_cod & vbCRLF
         ' WScript.Echo "Codigo de Casa: " & casa_cod & vbCRLF
         ' WScript.Echo "Jefe de Flia: " & flia_jefe & vbCRLF

         row_in = 12
         row_in_string = CStr(row_in)
         curCol_tmp = "A"
         fecha_visita = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
         Do While fecha_visita <> ""
            curCol_tmp = "A"
            xlCellName_out = curCol_tmp & curRow_out_string
            ' WScript.Echo "Cell_Out: " & xlCellName_out & ". Cell_in: " & (curCol_tmp & row_in_string) & "." & vbCRLF
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, fecha_visita

            curCol_tmp = "C"
            tipo_criadero_dc = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))            
            curCol_tmp = "B"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, tipo_criadero_dc

            curCol_tmp = "G"
            abatizado = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))            
            curCol_tmp = "M"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, abatizado

            curCol_tmp = "G"
            larvas_bool = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))            
            curCol_tmp = "E"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, larvas_bool

            curCol_tmp = "M"
            tipo_criadero_senepa = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))            
            curCol_tmp = "C"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, tipo_criadero_senepa

            curCol_tmp = "N"
            cant_cont = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))            
            curCol_tmp = "D"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, cant_cont

            curCol_tmp = "G"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, manzana_cod

            curCol_tmp = "H"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, casa_cod

            curCol_tmp = "I"
            xlCellName_out = curCol_tmp & curRow_out_string
            SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, flia_jefe

            ' WScript.Echo "Fecha de visita: " & fecha_visita & vbCRLF
            ' WScript.Echo "Tipo de criadero segun los codigos de DengueChat: " & tipo_criadero_dc & vbCRLF
            ' WScript.Echo "Abatizado?: " & abatizado & vbCRLF
            ' WScript.Echo "Larvas?: " & larvas_bool & vbCRLF
            ' WScript.Echo "Tipo de criadero segun los codigos del SENEPA: " & tipo_criadero_senepa & vbCRLF
            ' WScript.Echo "Cantidad de contenedores: " & cant_cont & vbCRLF

            row_in = row_in + 1
            row_in_string = CStr(row_in)
            curRow_out = curRow_out + 1
            curRow_out_string = CStr(curRow_out)

            curCol_tmp = "A"
            fecha_visita = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
         Loop

      End If
      
      If xlWorksheet_in_name = "Comentarios_Tipos_Criaderos" Then
         
      End If

      If xlWorksheet_in_name = "Larvicida" Then
         
      End If

      ' curRow_out = curRow_out + 1
      ' ' Set Cells = currentWorksheet.Cells
      ' xlCellName_in = "B1"
      ' xlCellName_out = "B1"

      ' xlCellContents = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, xlCellName_in)
      ' WScript.Echo xlCellContents
      ' SetExcel xlApp_out, xlBook_out, xlFileName_out, xlWorksheet_out_name, xlCellName_out, xlCellContents
      ' WScript.Echo "Success!"

      ' REM Loop through each row in the worksheet 
      ' For row = 0 to (usedRowsCount-1)
         
      '    REM Loop through each column in the worksheet 
      '    For column = 0 to usedColumnsCount-1
      '       REM only look at rows that are in the "used" range
      '       curRow = row+top
      '       REM only look at columns that are in the "used" range
      '       curCol = column+left
      '       REM get the value/word that is in the cell 
      '       word = Cells(curRow,curCol).Value
      '       REM display the column on the screen
      '       If curRow = 1 and curCol = 2 Then
      '          WScript.Echo (word)
      '       End If
               
      '    Next
      ' Next


      REM We are done with the current worksheet, release the memory
      Set currentWorksheet_in = Nothing 'Copy'
   Next
   ' Save changes
   xlBook_out.Save
   ' Save changes and close the spreadsheet 
   ' xlBook_in.Save

   ' Close the spreadsheet
   xlBook_in.Close
   Set xlBook_in = Nothing 
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