Dim FieldNames() AS string

Begin Dialog dialogbox1 50,42,178,205,"Newdialog", .NewDialog
  Text 6,10,151,11, "Select the various FIELD headers below", .Text1
  DropListBox 5,40,61,11, FieldNames(), .DropListBox1
  DropListBox 5,76,60,11, FieldNames(), .DropListBox3
  Text 6,31,50,7, "Select the DATE", .Text1
  OKButton 14,146,50,14, "OK", .OKButton1
  CancelButton 99,149,50,13, "Cancel", .CancelButton1
  Text 6,104,55,7, "Select the BALANCE", .Text1
  DropListBox 5,115,61,11, FieldNames(), .DropListBox5
  DropListBox 84,40,61,11, FieldNames(), .DropListBox2
  DropListBox 86,75,61,11, FieldNames(), .DropListBox4
  Text 85,31,50,7, "Select the DETAILS", .Text1
  Text 92,67,50,7, "Select the DEBIT", .Text1
  Text 5,66,50,7, "Select the CREDIT", .Text1
End Dialog

Begin Dialog dialogbox2 50,46,173,151,"NewDialog", .NewDialog
  Text 15,42,132,11, "OR Enter Triggerwords Seperated by "","" Delimiter", .Text1
  OKButton 9,83,55,14, "OK", .OKButton1
  CancelButton 93,84,55,14, "Cancel", .CancelButton1
  PushButton 47,106,55,14, "Back", .PushButton1
  TextBox 6,56,151,14, .TextBox1
  PushButton 9,10,141,14, "Use builtin triggerwords", .PushButton2
End Dialog

' ========================================================================================
' IDEAScript:       		MergeRows_BankStatement.iss
' Author:          		Ayo Osafehinti
' Created On:      		October 2025
' Description:     		'This IDEAScript automates the process of merging multi-line bank statement entries
' 			into single, consolidated transaction records. It exports data from IDEA to Excel,
' 			dynamically injects and runs a VBA macro to group related lines based on transaction types 
' 			as trigger words for a new line (e.g., debit, counter credit, standing order), preserves the original 
' 			text format of date fields, and reimports the merged results back into IDEA as a new database.
' 			The script improves efficiency and consistency in cleaning and structuring
' 			bank transaction data for further analysis.
'___________________________________________________________


Option Explicit

Dim db As Object
Dim dbName As String
Dim table As Object
Dim task As Object
Dim thisField As Object
Dim field As Object
Dim i As Integer
Dim fieldCount As Integer
Dim columnCount As Integer
Dim dlg1 As dialogbox1, dlg2 As dialogbox2
Dim Button As Integer, Button1 As Integer
Dim DateVal As Integer, Credit As Integer, Debit As Integer, Balance As Integer, Detail As Integer, Response As String
Dim Transactiontype As Variant
Dim dateField As Integer, detailsField As Integer, creditField As Integer, debitField As Integer, balanceField As Integer
Dim SelectedDateField As String, SelectedDetailField As String, SelectedCreditField As String, SelectedDebitField As String, SelectedBalanceField As String
Dim counter As Integer, extension As String
Dim exportPath As String, macroPath As String, importPath As String
Dim xlApp As Object, wb As Object
Dim macroCode As String
Dim fso As Object, file As Object
Dim macroPathFile As String
Dim baseFileName As String
Dim dbFullPath As String
Dim baseExportPath As String
Dim logFilePath As String, timeStamp As String
Dim BuiltInTransactiontype As Variant

Sub Main
	timeStamp = Format(Now, "yyyymmdd_hhnnss")
	logFilePath = Client.WorkingDirectory & "Other.ILB\" & baseFileName & "_" & timeStamp & "_log.txt"
	
	On Error GoTo HandleError
	
	BuiltInTransactiontype = "debit, counter credit, standing order, direct debit, transfer, cash deposit, bill payment, card purchase, cash withdrawal" 'Using the Transactiontype to indicate start of a new line. (update here as required).
	
	Response = MsgBox ("No Guarantees. IDEAScript by Ayo",1+64)
	If Response = 2 Then 
		MsgBox "Macro terminated",48
		Exit Sub
	End If
	Set db = Client.CurrentDatabase
	Set table = db.TableDef
	fieldCount = table.Count
	If fieldCount = 0 Then
		MsgBox "No fields found!",48
		Exit Sub
	End If

	columnCount = 0
	For i = 1 To fieldCount
	Set thisField = table.GetFieldAt(i)
		If thisField.IsCharacter() Then
			columnCount = columnCount + 1
			ReDim Preserve FieldNames(1 To columnCount)
			FieldNames(columnCount) = thisField.Name
		End If
	Next i
	If columnCount = 0 Then
		MsgBox "You need at least 5 character value fields for this macro to work.",48
		Exit Sub
	End If

	Dialog_1:		
	Button = Dialog(dlg1)
	If Button = 0 Then
		MsgBox "Macro terminated",48
		Exit Sub
	Else
		DateVal = dialogbox1.DropListbox1 + 1 ' assumes OptionButtons returns 0-based index
		Detail = dialogbox1.DropListbox2 + 1
		Credit = dialogbox1.DropListbox3 + 1
		Debit = dialogbox1.DropListbox4 + 1
		Balance = dialogbox1.DropListbox5 + 1
		
		SelectedDateField = FieldNames(DateVal)
		SelectedDetailField = FieldNames(Detail)	
		SelectedCreditField = FieldNames(Credit)	
		SelectedDebitField = FieldNames(Debit)	
		SelectedBalanceField = FieldNames(Balance)	
				
		Call WriteToLog("Selected character fields in IDEA")
	End If
	If DateVal = Detail Or DateVal = Credit Or DateVal = Debit Or DateVal = Balance Or Detail = Credit Or Detail = Debit Or Detail = Balance Or Credit = Debit Or Credit = Balance  Or Debit = Balance Then
		MsgBox "You must select 5 different fields.",48
        		GoTo Dialog_1
	End If

	Button1 = Dialog(dlg2)
	If Button1 = 0 Then
		MsgBox "Macro terminated.",64
		Exit Sub
	End If
	If Button1 = 1 Then
		GoTo Dialog_1
	End If
	If Button1 = -1 Then
		Transactiontype = Trim(dialogbox2.TextBox1)

		Call WriteToLog("Transactiontype inputted in IDEA")
		
		Call Export		
		
		Call mergeRows

	End If	
	
	If Button1 = 2 Then
		Transactiontype = BuiltInTransactiontype

		Call WriteToLog("Transactiontype inputted in IDEA")
		
		Call Export		
		
		Call mergeRows

	End If
		
	Set xlApp = CreateObject("Excel.Application")
	xlApp.Visible = False
	Set wb = xlApp.Workbooks.Open(exportPath)
	wb.VBProject.VBComponents.Import macroPathFile
		
	Set fso = CreateObject("Scripting.FileSystemObject")
	macroPath = fso.GetParentFolderName(exportPath) & "\" & fso.GetBaseName(exportPath) & ".xlsm"
	wb.SaveAs macroPath, 52  ' Save as macro-enabled .xlsm
	Call WriteToLog("Saved as .xlsm: " & macroPath)
	wb.Close False
			
	Set wb = xlApp.Workbooks.Open(macroPath)
	wb.Save
	wb.Close False
	Set wb = xlApp.Workbooks.Open(macroPath)
	xlApp.Run "mergeRows"
	Call WriteToLog("Ran macro: mergeRows in " & macroPath)
	wb.Save
	wb.Close False
	xlApp.Quit
	Set wb = Nothing
	Set xlApp = Nothing
		
	Call Import
	
	Client.OpenDatabase (importPath & "-Merged_Transactions")

	Call MopUp
	
	MsgBox "All done! I just made your task faster.", 64, "Task Complete"
	

HandleError:
	If Err.Number <> 0 Then 
		MsgBox "An error occurred: " & Err.Description, 16, "Script Error" 	
		Exit Sub
	Call WriteToLog("ERROR: " & Err.Description)
	End If
	
	If Not db Is Nothing Then Set db = Nothing
	If Not table Is Nothing Then Set table = Nothing
	If Not Task Is Nothing Then Set Task = Nothing
	If Not field Is Nothing Then Set field = Nothing
		
End Sub

Sub MopUp()
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile macroPathFile, True 
End Sub
 		
Sub Export() 
	Set fso = CreateObject("Scripting.FileSystemObject")   		    		
        	Set db = Client.CurrentDatabase
    	Set task = db.ExportDatabase
	dbFullPath = db.Name
	baseFileName = fso.GetBaseName(dbFullPath)
	extension = ".xlsx"
    	baseExportPath = Client.WorkingDirectory & "Exports.ILB\" & baseFileName
    	exportPath = baseExportPath & "_" & counter & extension
    	counter = 1
    	Do While fso.FileExists(exportPath)
    		exportPath = baseExportPath & "_" & counter & extension
    		counter = counter + 1
    	Loop
    	task.IncludeAllFields
    	task.PerformTask exportPath, "database", "XLSX", 1, db.Count, ""
    	Call WriteToLog("Exported to Excel: " & exportPath)
End Sub


Sub Import()
	Set fso = CreateObject("Scripting.FileSystemObject")
	importPath = fso.GetBaseName(exportPath)
	Set task = Client.GetImportTask("ImportExcel")
	task.FileToImport = macroPath
	task.SheetToImport = "Merged_Transactions"
	task.OutputFilePrefix = importPath
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "FALSE"
	task.PerformTask
	Call WriteToLog("Imported file from: " & exportPath)
End Sub


Sub mergeRows()
	Call Inject_Macros
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set file = fso.CreateTextFile(macroPathFile, True)
	
	file.WriteLine "Sub mergeRows()"
	file.WriteLine "    	Dim wsSrc As Worksheet"
	file.WriteLine "	Dim wsOut As Worksheet"
	file.WriteLine "	Dim lastRow As Long"
	file.WriteLine "    	Dim i As Long"
	file.WriteLine "	Dim DateVal As String " 
	file.WriteLine "	Dim " & SelectedDetailField & " As String " 
	file.WriteLine "	Dim " & SelectedCreditField & " As String " 
	file.WriteLine "	Dim " & SelectedDebitField & " As String " 
	file.WriteLine "	Dim " & SelectedBalanceField & " As String " 
	file.WriteLine "	Dim tranxtype As Variant"		
	file.WriteLine "    	Dim istranxtype As Boolean"
	file.WriteLine "	Dim outRow As Long"
	file.WriteLine "	Dim word As Variant"	
	file.WriteLine "    	Dim newSheetName As String"
	file.WriteLine "	Set wsSrc = ActiveWorkbook.Sheets(1)"
	file.WriteLine "	newSheetName = ""Merged_Transactions"""	
	file.WriteLine "    	Set wsOut = ActiveWorkbook.Sheets.Add"
	file.WriteLine "	wsOut.Name = newSheetName"
	file.WriteLine "	lastRow = wsSrc.Cells(wsSrc.Rows.Count, " & detailsField & ").End(xlUp).Row"	
	file.WriteLine "    	tranxtype = Split(""" & Transactiontype & """, "","")"
	file.WriteLine "	wsOut.Cells(1, ""A"").Value = """ & SelectedDateField & """"
	file.WriteLine "	wsOut.Cells(1, ""B"").Value = """ & SelectedDetailField & """"
	file.WriteLine "	wsOut.Cells(1, ""C"").Value = """ & SelectedCreditField & """"
	file.WriteLine "	wsOut.Cells(1, ""D"").Value = """ & SelectedDebitField & """"
	file.WriteLine "	wsOut.Cells(1, ""E"").Value = """ & SelectedBalanceField &""""
	file.WriteLine "	outRow = 2"
	file.WriteLine "	DateVal = """""
	file.WriteLine "	" & SelectedDetailField & "  = """""
	file.WriteLine "	" & SelectedCreditField & " = """""
	file.WriteLine "	" & SelectedDebitField & "  = """""
	file.WriteLine "	" & SelectedBalanceField & " = """""	
	file.WriteLine "	For i = 2 To lastRow"
	file.WriteLine "		istranxtype = False"
	file.WriteLine "		For Each word In tranxtype"
	file.WriteLine "			If InStr(1, LCase(wsSrc.Cells(i, " & detailsField & ").Value), Trim(LCase(word))) > 0 Then"
	file.WriteLine "				istranxtype = True"	
	file.WriteLine "				Exit For"	
	file.WriteLine "			End If"
	file.WriteLine "		Next word"
	file.WriteLine "		If istranxtype Then"
	file.WriteLine "			If " & SelectedDetailField & " <> """" Then"
	file.WriteLine "				wsOut.Cells(outRow, ""A"").NumberFormat = ""@"""
	file.WriteLine "				wsOut.Cells(outRow, ""A"").Value = DateVal"	
	file.WriteLine "				wsOut.Cells(outRow, ""B"").Value = Trim(" & SelectedDetailField & ")"
	file.WriteLine "				wsOut.Cells(outRow, ""C"").Value =" & SelectedCreditField & ""
	file.WriteLine "				wsOut.Cells(outRow, ""D"").Value = " & SelectedDebitField & ""
	file.WriteLine "				wsOut.Cells(outRow, ""E"").Value = " & SelectedBalanceField &""
	file.WriteLine "				outRow = outRow + 1"
	file.WriteLine "			End If"
	file.WriteLine "			DateVal = wsSrc.Cells(i, " & dateField & ").Text"
	file.WriteLine "			" & SelectedDetailField & " = wsSrc.Cells(i, " & detailsField & ").Value"
	file.WriteLine "			" & SelectedCreditField & " = wsSrc.Cells(i, " & creditField & ").Value"
	file.WriteLine "			" & SelectedDebitField & " = wsSrc.Cells(i, " & debitField & ").Value"	
	file.WriteLine "			" & SelectedBalanceField &" = wsSrc.Cells(i, " & balanceField & ").Value"	
	file.WriteLine "		Else"
	file.WriteLine "			If Trim(wsSrc.Cells(i, " & dateField & ").Value) <> """" Then"
	file.WriteLine "				If DateVal = """" Then"
	file.WriteLine "					DateVal = wsSrc.Cells(i, " & dateField & ").Text"	
	file.WriteLine "				Else"
	file.WriteLine "					DateVal = DateVal &  wsSrc.Cells(i, " & dateField & ").Text"
	file.WriteLine "				End If"
	file.WriteLine "			End If"
	file.WriteLine "			If Trim(wsSrc.Cells(i, " & detailsField & ").Value) <> """" Then " & SelectedDetailField & " = " & SelectedDetailField & " & "" "" & wsSrc.Cells(i, " & detailsField & ").Value"
	file.WriteLine "			If Trim(wsSrc.Cells(i, " & creditField & ").Value) <> """" Then " & SelectedCreditField & " = wsSrc.Cells(i, " & creditField & ").Value"
	file.WriteLine "			If Trim(wsSrc.Cells(i, " & debitField & ").Value) <> """" Then " & SelectedDebitField & " = wsSrc.Cells(i, " & debitField & ").Value"
	file.WriteLine "			If Trim(wsSrc.Cells(i, " & balanceField & ").Value) <> """" Then " & SelectedBalanceField & " = wsSrc.Cells(i, " & balanceField & ").Value"
	file.WriteLine "		End If"	
	file.WriteLine "	Next i"		
	file.WriteLine "	If " & SelectedDetailField & " <> """" Then"
	file.WriteLine "		wsOut.Cells(outRow, ""A"").NumberFormat = ""@"""
	file.WriteLine "		wsOut.Cells(outRow, ""A"").Value = DateVal"
	file.WriteLine "		wsOut.Cells(outRow, ""B"").Value = Trim(" & SelectedDetailField & ")"
	file.WriteLine "		wsOut.Cells(outRow, ""C"").Value = " & SelectedCreditField & ""
	file.WriteLine "		wsOut.Cells(outRow, ""D"").Value = " & SelectedDebitField & ""
	file.WriteLine "		wsOut.Cells(outRow, ""E"").Value = " & SelectedBalanceField &""
	file.WriteLine "	End If"
	file.WriteLine "End Sub"		
	
	file.Close
	Set file = Nothing
	Set fso = Nothing
End Sub


Sub Inject_Macros
	macroPathFile = Client.WorkingDirectory & "Exports.ILB\mergeRows.bas"
	dateField = GetColumnNumberFromField(SelectedDateField)
	detailsField = GetColumnNumberFromField(SelectedDetailField)
	creditField = GetColumnNumberFromField(SelectedCreditField)
	debitField = GetColumnNumberFromField(SelectedDebitField)
	balanceField = GetColumnNumberFromField(SelectedBalanceField)
End Sub

Function GetColumnNumberFromField(fieldName As String) As Integer
	For i = 1 To table.Count
	Set field = table.GetFieldAt(i)
		If field.Name = fieldName Then
			GetColumnNumberFromField = i 
			Exit Function
		End If
	Next
End Function

Sub WriteToLog(message As String)
	Dim fsoLog As Object, logFile As Object
	Set fsoLog = CreateObject("Scripting.FileSystemObject")
	Set logFile = fsoLog.OpenTextFile(logFilePath, 8, True)
	logFile.WriteLine Now & " - " & message
	logFile.Close
	
	Set logFile = Nothing
	Set fsoLog = Nothing
End Sub








