'
' Copyright (c) 2020 Mark Crowe <https://github.com/markcrowe-com>. All rights reserved.
'
Sub CreateWorkbookTableOfContents()
	'PURPOSE: Add a Table of Contents worksheets to easily navigate to any tab
	'SOURCE: www.TheSpreadsheetGuru.com

	Dim sht As Worksheet
	Dim Content_sht As Worksheet
	Dim myArray As Variant
	Dim x As Long, y As Long
	Dim shtName1 As String, shtName2 As String
	Dim ContentName As String

	'Inputs
	ContentName = "Contents"

	'Optimize Code
	Application.DisplayAlerts = False
	Application.ScreenUpdating = False

	'Delete Contents Sheet if it already exists
	On Error Resume Next
	Worksheets("Contents").Activate
	On Error GoTo 0

	If ActiveSheet.Name = ContentName Then
		myAnswer = MsgBox("A worksheet named [" & ContentName &
		  "] has already been created, would you like to replace it?", vbYesNo)

		'Did user select No or Cancel?
		If myAnswer <> vbYes Then GoTo ExitSub

		'Delete old Contents Tab
		Worksheets(ContentName).Delete
	End If

	'Create New Contents Sheet
	Worksheets.Add Before:=Worksheets(1)

'Set variable to Contents Sheet
  Set Content_sht = ActiveSheet

'Format Contents Sheet
  With Content_sht
		.Name = ContentName
		.Range("B1") = "Table of Contents"
		.Range("B1").Font.Bold = True
	End With

	'Create Array list with sheet names (excluding Contents)
	ReDim myArray(1 To Worksheets.Count - 1)

	For Each sht In ActiveWorkbook.Worksheets
		If sht.Name <> ContentName Then
			myArray(x + 1) = sht.Name
			x = x + 1
		End If
	Next sht

	'Alphabetize Sheet Names in Array List
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
			.Hyperlinks.Add.Cells(x + 2, 3), "", _
      SubAddress:="'" & sht.Name & "'!A1", _
      TextToDisplay:=sht.Name
      .Cells(x + 2, 2).Value = x
		End With
	Next x
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