' Given a folder, open each file in it, parse it
' , extract the data, and store it into the DB.

' The procedure followed the first time this script was executed is as follows:
' 1) Downloaded KML data from MyMaps.
' 2) Used an Online tool to convert it into an excel-compatible file.
' 3) Fixed some format details, in order to leave it like the example excel file,
' and had to manually add the 'data-locations-type' column.
' 4) Created a directory and stored all the files to be parsed there.
' 5) Changed the value of 'objStartFolder' to the path of the directory
' created in 4) (has to end with a '\').
' 6) Configured the ODBC properly to connect the desired DB.
' 7) Run the script.
' 
' Obs:
' - The values of city_id and neighborhood_id are hardcoded to 
' 9 and 52, respectively. That corresponds to city 'Asuncion'
' and neighborhood 'San Cayetano'.
'  The correct and complete approach would be
' store the locations-to-neighborhood mapping
' (or the manzanas-to-neighborhoods one, if manzana_cod
' is extracted from location_code) somewhere, and consult it to
' automatically map the location to its corresponding neighborhood.
'  For now, all locations to be parsed must be from a same neighborhood,
' and the 'neighborhood_id' variable must be set to its corresponding id.
' This must be done every time locations from a different neighborhood
' are to be parsed/stored.

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
REM Since it is a folder, it has to end with a backslacsh (\)
objStartFolder = "C:\Users\usuario\Documents\Trabajo\TopaDengue\Mapeo_Casas\casas_maps_xlsx\"
Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For each objFile in colFiles
   xlFileName_in = objStartFolder & objFile.Name
   ' Wscript.Echo "Parsing " & xlFileName_in & " ..." & vbCRLF
   
   Read_store_into_Db xlApp_in, xlFileName_in
Next

xlApp_in.Quit
Set xlApp_in = Nothing
WScript.Echo "Final Success!"

Private Sub Read_store_into_Db(xlApp_in, xlFileName_in):
   REM the Excel Book
   Dim xlBook_in
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


   Dim row_in
   Dim row_in_string
   Dim curCol_tmp

   REM variables that store the relevant data of the csv
   Dim location_cod
   Dim location_type 'Tipo de 
   Dim lat
   Dim lon
   Dim created_at_timestamp
   Dim updated_at_timestamp
   Dim neighborhood_id
   Dim source
   Dim city_block_id
   Dim city_id

   REM Connection Object
   Dim ConnObj
   REM Recordset Object
   Dim Recordset
   REM Database Command String
   Dim DBCommStr


   ' Script.Echo "Entered Read_store_into_Db" & vbCRLF
   ' Create the Excel Workbook Object. 
   Set xlBook_in = xlApp_in.Workbooks.Open(xlFileName_in)

   Set ConnObj = CreateObject("ADODB.Connection")
   ConnObj.Open "DSN=PostgreSQL30"

   REM How many worksheets are in this Excel documents
   inWorkSheetCount = xlApp_in.Worksheets.Count

   REM Loop through each worksheet
   For counter = 1 to inWorkSheetCount
      Set currentWorksheet_in = xlApp_in.ActiveWorkbook.Worksheets(counter)
      xlWorksheet_in_name = currentWorksheet_in.Name

      ' WScript.Echo "-----------------------------------------------"
      ' WScript.Echo "Reading data from worksheet No. " & counter & ": " & xlWorksheet_in_name & vbCRLF

      If xlWorksheet_in_name = "Sheet 1" Then
         row_in = 2
         row_in_string = CStr(row_in)

         curCol_tmp = "A"
         location_cod = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
         curCol_tmp = "C"
         location_type = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
         curCol_tmp = "I"
         lat = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
         curCol_tmp = "J"
         lon = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))

         Do While (location_cod <> "") OR (lat <> "") OR (lon <> "")
            Do ' For achieving C++'s-Continue-like behaviour, had to add an extra loop that always evaluates to False
               If (location_cod = "") OR (lat = "") OR (lon = "") Then
                  Exit Do ' Works as the Continue statement
               End If

               created_at_timestamp = "current_timestamp"
               updated_at_timestamp = "current_timestamp"
               neighborhood_id = 52 ' San Cayetano
               source = "ODK"
               city_block_id = "DEFAULT"
               city_id = 9 ' Asuncion
               ' WScript.Echo "Location code: " & location_cod & vbCRLF
               ' WScript.Echo "Tipo de location: " & location_type & vbCRLF
               ' WScript.Echo "Latitude: " & lat & vbCRLF
               ' WScript.Echo "Longitude?: " & lon & vbCRLF
               DBCommStr = "INSERT INTO locations "
               DBCommStr = DBCommStr & "(address, latitude, longitude, created_at, updated_at, neighborhood_id, source, city_block_id, city_id, location_type)"
               DBCommStr = DBCommStr & "VALUES ("
               DBCommStr = DBCommStr & "'" & location_cod & "'" & ", "
               DBCommStr = DBCommStr & lat & ", "
               DBCommStr = DBCommStr & lon & ", "
               If created_at_timestamp <> "" Then
                  DBCommStr = DBCommStr & created_at_timestamp & ", "
               Else
                  DBCommStr = DBCommStr & "DEFAULT, "
               End If
               If updated_at_timestamp <> "" Then
                  DBCommStr = DBCommStr & updated_at_timestamp & ", "
               Else
                  DBCommStr = DBCommStr & "DEFAULT, "
               End If
               If neighborhood_id <> "" Then
                  DBCommStr = DBCommStr & neighborhood_id & ", "
               Else
                  DBCommStr = DBCommStr & "DEFAULT, "
               End If
               If source <> "" Then
                  DBCommStr = DBCommStr & source & ", "
               Else
                  DBCommStr = DBCommStr & "DEFAULT, "
               End If
               If city_block_id <> "" Then
                  DBCommStr = DBCommStr & city_block_id & ", "
               Else
                  DBCommStr = DBCommStr & "DEFAULT, "
               End If
               If city_id <> "" Then
                  DBCommStr = DBCommStr & city_id & ", "
               Else
                  DBCommStr = DBCommStr & "DEFAULT, "
               End If
               If location_type <> "" Then
                  DBCommStr = DBCommStr & "'" & location_type & "'" & ")"
               Else
                  DBCommStr = DBCommStr & "DEFAULT)"
               End If
               
               ' WScript.Echo DBCommStr & vbCRLF

               Set Recordset = ConnObj.Execute(DBCommStr)

            Loop While False

            row_in = row_in + 1
            row_in_string = CStr(row_in)

            curCol_tmp = "A"
            location_cod = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
            curCol_tmp = "C"
            location_type = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
            curCol_tmp = "I"
            lat = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
            curCol_tmp = "J"
            lon = GetExcel(xlApp_in, xlBook_in, xlFileName_in, xlWorksheet_in_name, (curCol_tmp & row_in_string))
         Loop
      End If
      REM We are done with the current worksheet, release the memory
      Set currentWorksheet_in = Nothing 'Copy'
   Next
   ' Save changes and close the spreadsheet 
   ' xlBook_in.Save

   ' Close the spreadsheet
   xlBook_in.Close(false)
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