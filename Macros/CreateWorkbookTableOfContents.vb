'
' Copyright (c) 2020 Mark Crowe <https://github.com/markcrowe-com>. All rights reserved.
'
Sub CreateWorkbookTableOfContents()
	'Optimize Code
	Application.DisplayAlerts = False
	Application.ScreenUpdating = False

	Dim ContentName As String
	ContentName = "Contents"

	'Delete Exisiting Table of Contents WorkSheet
	On Error Resume Next
	Worksheets(ContentName).Activate
	On Error GoTo 0

	If ActiveSheet.Name = ContentName Then
		myAnswer = MsgBox("A worksheet named [" & ContentName & "] has already been created, would you like to replace it?", vbYesNo)

		'Did user select No or Cancel?
		If myAnswer <> vbYes Then GoTo ExitSub

		'Delete old Contents Tab
		Worksheets(ContentName).Delete
	End If

	'Create New Contents Sheet
	Worksheets.Add Before:=Worksheets(1)

	'Set variable to Contents Sheet
	Dim Content_sht As Worksheet
    Set Content_sht = ActiveSheet

    'Format Contents Sheet
    With Content_sht
		.Name = ContentName
		.Range("B1") = "Table of Contents"
		.Range("B1").Font.Bold = True
	End With

	Dim myArray As Variant

	'Create Array list with sheet names (excluding Contents)
	ReDim myArray(1 To Worksheets.Count - 1)
	Dim x As Long, y As Long

	Dim sht As Worksheet
	For Each sht In ActiveWorkbook.Worksheets
		If sht.Name <> ContentName Then
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
		With Content_sht
			.Hyperlinks.Add.Cells(x + 2, 3), "", SubAddress:="'" & sht.Name & "'!A1", TextToDisplay:=sht.Name
            .Cells(x + 2, 2).Value = x
		End With
	Next x

	Content_sht.Activate
	Content_sht.Columns(3).EntireColumn.AutoFit

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