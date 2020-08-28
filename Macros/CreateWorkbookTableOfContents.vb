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

	Dim myArray As Variant

	'Create Array list with sheet names (excluding Contents)
	ReDim myArray(1 To Worksheets.Count - 1)
	Dim x As Long

	Dim sht As Worksheet
	For Each sht In ActiveWorkbook.Worksheets
		If sht.Name <> tableOfConentsWorksheetName Then
			myArray(x + 1) = sht.Name
			x = x + 1
		End If
	Next sht

	'Create Table of Contents
	For x = LBound(myArray) To UBound(myArray)
        Set sht = Worksheets(myArray(x))
        sht.Activate
		With tableOfConentsWorksheet
			.Hyperlinks.Add.Cells(x + 2, 3), "", SubAddress:="'" & sht.Name & "'!A1", TextToDisplay:=sht.Name
            .Cells(x + 2, 2).Value = x
		End With
	Next x

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
