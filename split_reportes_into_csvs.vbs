' Given the dataset of the recipients records, in the form of a spreadsheet,
'  parse each of the records, and create 1 csv for each house, storing all the
'  information in it, in the appropiate format.
' 
' Procedure:
'  1) Download the spreadsheet that contains the records to be split
'   (By the time it can be found in
'    https://docs.google.com/spreadsheets/d/1skfxDEsuZ0IvnADj179Ux_L2vdNO5jOwBW2mx2K8YYU/edit?usp=sharing)
'  2) Set parameters
'   (xlFileName_in, StartFolder_out, xlFileName_out, initial_row_in, limit_row_in)
'  3) Run the Script

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
Dim xlApp_out
REM the Excel Application
Dim xlApp_in
REM the Excel Book
Dim xlBook_in
REM the Excel Book
Dim xlBook_out
REM the path to the excel file to be read
Dim xlFileName_out
REM the path to the excel file to be write
Dim xlFileName_in
REM the worksheet we are currently getting data from
Dim currentWorksheet_in
REM the worksheet we are currently getting data from
Dim currentWorksheet_out
REM the Worsheet Name
Dim xlWorksheet_in_name
REM the Worsheet Name
Dim xlWorksheet_out_name
REM the row number corresponding to the first row that we are taking into account when reading the joined file
Dim initial_row_in
REM for debugging. the row number corresponding to the last row that we are taking into account when reading the joined file
Dim limit_row_in
REM Absolute path to the folder where the split xlsx will be stored
Dim StartFolder_out


' Create the Excel App Objects
Set xlApp_out = CreateObject("Excel.Application")
Set xlApp_in = CreateObject("Excel.Application")

' Path to the unified dataset
xlFileName_in = "C:Users\...\Encuestas Entomologicas Profesionales Simplificado.xlsx"
' Path to folder where all the ouput csvs will be stored
StartFolder_out = "C:\Users\...\output_csvs\"
' Path to Template of a csv (in DengueChat format. Can be found in the repository)
xlFileName_out = "C:\Users\...\forma_csv.xlsx"
' Create the Excel Workbook Object. 
Set xlBook_in = xlApp_in.Workbooks.Open(xlFileName_in)
Set xlBook_out = xlApp_out.Workbooks.Open(xlFileName_out)

Set currentWorksheet_in = xlApp_in.ActiveWorkbook.Worksheets(1)
xlWorksheet_in_name = currentWorksheet_in.Name
' WScript.Echo "In-Worksheet name: " & xlWorksheet_in_name & vbCRLF

' Row numbers that delimit the range of rows to be considered
'  (i.e. parsed and split into csvs). The limits are included.
' By default is set to [2693, 5449], corresponding to the records corresponding
' to "Relevamiento Julio 2018"

initial_row_in = 2693
limit_row_in = 5449

REM id_predio (manz-nrocasa-apell)
Dim id_predio
REM the actual excel row number of the current recipient
Dim current_recipient_row_out

Dim xlCellName_in_id_predio
Dim last_id_predio
Dim current_row_in
Dim current_row_in_string
Dim current_row_out_string
Dim current_column_tmp
Dim xlCellName_out
Dim fecha_visita
Dim tipo_criadero_dc
Dim abatizado
Dim larvas_bool
Dim larvicida_utilizado
Dim permiso_entrada
Dim descr_cont
Dim tipo_criadero_senepa
Dim cant_cont
Dim xlFileName_out_new

current_row_in = initial_row_in
current_row_in_string = CStr(current_row_in)

current_column_tmp = "R"
id_predio = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))
' WScript.Echo "id_predio (" & (current_column_tmp & current_row_in_string) & "): " & id_predio & vbCRLF
Do While id_predio <> "" AND current_row_in < limit_row_in
   Set currentWorksheet_out = xlApp_out.ActiveWorkbook.Worksheets(1)
   xlWorksheet_out_name = currentWorksheet_out.Name
   ' WScript.Echo "xlWorksheet_out_name: " & xlWorksheet_out_name & vbCRLF

   xlCellName_out = "B1"
   ' WScript.Echo "Cell_Out: " & xlCellName_out & ". Cell_in: " & (current_column_tmp & current_row_in_string) & "." & vbCRLF
   SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, id_predio

   current_recipient_row_out = 12
   current_row_out_string = CStr(current_recipient_row_out)
   last_id_predio = id_predio
   Do While id_predio = last_id_predio AND current_row_in < limit_row_in
      ' "copy" the desired values into the "clipboard"
      ' and "Paste" the desired values into the destination cells
      current_column_tmp = "C"
      fecha_visita = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))
      current_column_tmp = "A"
      xlCellName_out = current_column_tmp & current_row_out_string
      ' WScript.Echo "Cell_Out: " & xlCellName_out & ". Cell_in: " & (current_column_tmp & current_row_in_string) & "." & vbCRLF
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, fecha_visita

      current_column_tmp = "T"
      tipo_criadero_dc = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))            
      current_column_tmp = "C"
      xlCellName_out = current_column_tmp & current_row_out_string
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, tipo_criadero_dc

      current_column_tmp = "O"
      abatizado = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))
      ' If abatizado = "" Then abatizado = 0        
      current_column_tmp = "F"
      xlCellName_out = current_column_tmp & current_row_out_string
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, abatizado

      current_column_tmp = "K"
      larvas_bool = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))            
      current_column_tmp = "G"
      xlCellName_out = current_column_tmp & current_row_out_string
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, larvas_bool

      current_column_tmp = "M"
      descr_cont = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))            
      current_column_tmp = "L"
      xlCellName_out = current_column_tmp & current_row_out_string
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, descr_cont

      current_column_tmp = "I"
      tipo_criadero_senepa = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))            
      current_column_tmp = "M"
      xlCellName_out = current_column_tmp & current_row_out_string
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, tipo_criadero_senepa

      current_column_tmp = "J"
      cant_cont = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))            
      current_column_tmp = "N"
      xlCellName_out = current_column_tmp & current_row_out_string
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, cant_cont

      current_recipient_row_out = current_recipient_row_out + 1
      current_row_out_string = CStr(current_recipient_row_out)
      current_row_in = current_row_in + 1
      current_row_in_string = CStr(current_row_in)

      current_column_tmp = "R"
      id_predio = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))
   Loop
   ' Save the current file as a separate xlsx, with the corresponding name (id_predio)
   REM Check for the case that there are no recipients records
   If tipo_criadero_dc = "" AND tipo_criadero_senepa = "" AND cant_cont = "" Then
      permiso_entrada = 1 ' To be defined. How do we know it's still 1?
      xlCellName_out = "C2"
      ' WScript.Echo "Cell_Out: " & xlCellName_out & ". Cell_in: " & (current_column_tmp & current_row_in_string) & "." & vbCRLF
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, permiso_entrada
   Else If tipo_criadero_dc <> "" AND tipo_criadero_senepa <> "" AND cant_cont <> "" Then
      permiso_entrada = 1
      xlCellName_out = "C2"
      ' WScript.Echo "Cell_Out: " & xlCellName_out & ". Cell_in: " & (current_column_tmp & current_row_in_string) & "." & vbCRLF
      SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, permiso_entrada
   Else
      WScript.Echo "Error." & vbCRLF _
       & " tipo_criadero_dc = " & tipo_criadero_dc & vbCRLF _
       & "tipo_criadero_senepa = " & tipo_criadero_senepa & vbCRLF _
       & "cant_cont = " & cant_cont & vbCRLF
   End If

   REM It would be better if I could directly open the sheet by name
   Set currentWorksheet_out = xlApp_out.ActiveWorkbook.Worksheets(3)
   xlWorksheet_out_name = currentWorksheet_out.Name
   ' WScript.Echo "xlWorksheet_out_name: " & xlWorksheet_out_name & vbCRLF
   If xlWorksheet_out_name <> "Larvicida" Then
      WScript.Echo "Ojo! La Hoja numero 3 del archivo " & xlFileName_out & " no se llama 'Larvicida'" & vbCRLF
      Exit Do
   End If

   current_column_tmp = "N"
   REM Como todas los registros de un mismo predio tienen el mismo valor para larvicida_utilizado
   REM se utiliza el del ultimo
   REM (se decrementa en 1 porque actualmente esta apuntando al que vino despues del ultimo recipiente)
   current_row_in_string = CStr(current_row_in - 1)
   larvicida_utilizado = GetExcel(xlBook_in, xlWorksheet_in_name, (current_column_tmp & current_row_in_string))
   xlCellName_out = "C2"
   ' WScript.Echo "Cell_Out: " & xlCellName_out & ". Cell_in: " & (current_column_tmp & current_row_in_string) & "." & vbCRLF
   SetExcel xlBook_out, xlWorksheet_out_name, xlCellName_out, larvicida_utilizado

   xlApp_out.DisplayAlerts = false
   ' Save changes
   xlFileName_out_new = StartFolder_out & last_id_predio & ".xlsx"
   ' Como al guardar un archivo con un nombre no se acepta el caracter '?', se lo reemplaza con su significado, es decir, que es dudoso
   xlFileName_out_new = Replace(xlFileName_out_new, "?", "_dudoso_")
   ' WScript.Echo "Saving file: " & xlFileName_out_new & vbCRLF
   xlBook_out.SaveAs xlFileName_out_new, 51
   ' xlApp_out.DisplayAlerts = true
   ' Close the newly saved book and open the template one
   xlBook_out.Close
   Set xlBook_out = xlApp_out.Workbooks.Open(xlFileName_out)

   current_row_in_string = CStr(current_row_in)
Loop

xlBook_in.Close
Set xlBook_in = Nothing

' Close the spreadsheet
' xlBook_outs.Close
Set xlBook_out = Nothing 
xlApp_out.Quit

xlApp_in.Quit
Set xlApp_out = Nothing
Set xlApp_in = Nothing
WScript.Echo "Final Success!"

Private Function GetExcel(xlBook, xlWorksheet, xlCellName)
   Dim strCellContents 
   
   ' Get the Cell Contents 
   strCellContents =     xlBook.worksheets(xlWorksheet).range(xlCellName).Value
   ' WScript.Echo "xlBook.worksheets('" & xlWorksheet & "').range('" & xlCellName & "').Value" & vbCRLF
   ' WScript.Echo strCellContents & vbCRLF
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
Private Sub SetExcel(xlBook, xlWorksheet, xlCellName, xlCellContents)    
   ' Set the value of the Cell 
   xlBook.worksheets(xlWorksheet).range(xlCellName).Value = xlCellContents 

   Exit Sub
End Sub