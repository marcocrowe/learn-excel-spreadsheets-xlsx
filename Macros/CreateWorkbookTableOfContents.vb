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

    'Format Contents Sheet
	With tableOfConentsWorksheet
		.Name = tableOfConentsWorksheetName
		.Range("B1") = "Table of Contents"
		.Range("B1").Style = "Heading 1"
	End With

	Dim myArray As Variant

	'Create Array list with sheet names (excluding Contents)
	ReDim myArray(1 To Worksheets.Count - 1)
	Dim x As Long, y As Long

	Dim sht As Worksheet
	For Each sht In ActiveWorkbook.Worksheets
		If sht.Name <> tableOfConentsWorksheetName Then
			myArray(x + 1) = sht.Name
			x = x + 1
		End If
	Next sht

	'Alphabetize Sheet Names in Array List
	Dim shtName1 As String, shtName2 As String
	For x = LBound(myArray) To UBound(myArray)
		For y = x To UBound(myArray)
			If UCase(myArray(y)) < UCase(myArray(x)) Then
				shtName1 = myArray(x)
				shtName2 = myArray(y)
				myArray(x) = shtName2
				myArray(y) = shtName1
			End If
		Next y
	Next x

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

	'A Splash of Guru Formatting! [Optional]
	Columns("A:B").ColumnWidth = 3.86
	Range("B1").Font.Size = 18
	Range("B1:F1").Borders(xlEdgeBottom).Weight = xlThin

	With Range("B3:B" & x + 1)
		.Borders(xlInsideHorizontal).Color = RGB(255, 255, 255)
		.Borders(xlInsideHorizontal).Weight = xlMedium
		.HorizontalAlignment = xlCenter
		.VerticalAlignment = xlCenter
		.Font.Color = RGB(255, 255, 255)
		.Interior.Color = RGB(91, 155, 213)
	End With

	'Adjust Zoom and Remove Gridlines
	ActiveWindow.DisplayGridlines = False
	ActiveWindow.Zoom = 130

ExitSub:
	'Optimize Code
	Application.DisplayAlerts = True
	Application.ScreenUpdating = True
End Sub