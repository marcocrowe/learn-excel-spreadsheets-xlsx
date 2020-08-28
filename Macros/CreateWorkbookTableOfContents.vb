'
' Copyright (c) 2020 Mark Crowe <https://github.com/markcrowe-com>. All rights reserved.
'
Sub CreateWorkbookTableOfContents()
	'Optimize Code
	Application.DisplayAlerts = False
	Application.ScreenUpdating = False

	Dim tableOfConentsWorksheetName As String
	tableOfConentsWorksheetName = "Contents"

	'Delete Exisiting Table of Contents WorkSheet
	On Error Resume Next
	Worksheets(tableOfConentsWorksheetName).Activate
	On Error GoTo 0

	If ActiveSheet.Name = tableOfConentsWorksheetName Then
		Dim message As String
		message = "A worksheet named [" & tableOfConentsWorksheetName & "] has already been created, would you like to replace it?"

		myAnswer = MsgBox(message, vbYesNo)

		If myAnswer <> vbYes Then
			GoTo ExitSub
		Else
			Worksheets(tableOfConentsWorksheetName).Delete
		End If

	End If

	'Create Table of Contents WorkSheet
	Worksheets.Add Before:=Worksheets(1)
	Dim tableOfConentsWorksheet As Worksheet
    Set tableOfConentsWorksheet = ActiveSheet

    'Format Worksheet Title
    headingCellReference = "B2:C2"
	With tableOfConentsWorksheet
		.Name = tableOfConentsWorksheetName
		.Range(headingCellReference).Merge
		.Range(headingCellReference).Style = "Heading 1"
		.Range(headingCellReference) = "Table of Contents"
	End With

	Dim dataTableHeadingRowIndex, nameColumnIndex, numberColumnIndex As Long
	numberColumnIndex = 2 'Column B
	nameColumnIndex = 3 'Column C
	dataTableHeadingRowIndex = 4 'Row 4

	tableOfConentsWorksheet.Cells(dataTableHeadingRowIndex, numberColumnIndex).Value = "#"
	tableOfConentsWorksheet.Cells(dataTableHeadingRowIndex, nameColumnIndex).Value = "Worksheet"

	Dim worksheet As worksheet
	Dim worksheetNumber As Long
	For Each worksheet In ActiveWorkbook.Worksheets
		If worksheet.Name <> tableOfConentsWorksheetName Then
			worksheetNumber = worksheetNumber + 1
			With tableOfConentsWorksheet
				.Hyperlinks.Add.Cells(worksheetNumber + dataTableHeadingRowIndex, nameColumnIndex), "", SubAddress:="'" & worksheet.Name & "'!A1", TextToDisplay:=worksheet.Name
                .Cells(worksheetNumber + dataTableHeadingRowIndex, numberColumnIndex).Value = worksheetNumber
			End With
		End If
	Next worksheet


	tableOfConentsWorksheet.Activate
	tableOfConentsWorksheet.Columns(3).EntireColumn.AutoFit

	'Adjust Zoom and Remove Gridlines
	ActiveWindow.DisplayGridlines = False
	ActiveWindow.Zoom = 130

ExitSub:
	'Optimize Code
	Application.DisplayAlerts = True
	Application.ScreenUpdating = True
End Sub
