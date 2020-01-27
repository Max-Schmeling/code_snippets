Function GetEffectiveUsedRange(Ws As Worksheet) As Range
    ' Returns the 'used value range' of the Range.UsedRange of the given Worksheet <Ws>.
    ' This means empty cells that are in the used range (e.g. when they are formatted) will be ignored.
    ' For Microsoft Excel by Max Schmeling

    Dim NewLastCol As Integer
    Dim NewLastRow As Long
    
    On Error GoTo InvalidUsedRange ' Expected error is no values in worksheet
    NewLastCol = Ws.UsedRange.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    NewLastRow = Ws.UsedRange.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    
    Set GetEffectiveUsedRange = Ws.UsedRange.Resize(NewLastRow - Ws.UsedRange.Row + 1, NewLastCol - Ws.UsedRange.Column + 1)
    Exit Function
    
InvalidUsedRange:
    Set GetEffectiveUsedRange = Ws.UsedRange ' return normal used range
End Function


Function IsWorksheetEmpty(Ws As Worksheet) As Boolean
    ' Returns True if the given Worksheet <Ws> is empty
    ' For Microsoft Excel by Max Schmeling

    If Ws.UsedRange.Address = Ws.Range("$A$1").Address And Ws.Range("$A$1").Text = "" Then
        IsWorksheetEmpty = True
    Else
        IsWorksheetEmpty = False
    End If
End Function


Function IsCellRange() As Boolean
    ' Returns True if the current <Selection> is a cell range.
    ' For Microsoft Excel by Max Schmeling

    If TypeName(Selection) <> "Range" Then
        IsCellRange = False
    Else
        IsCellRange = True
    End If
End Function
			
			
Sub ShowAllCells(Ws As Worksheet)
    ' Applies the following changes on the given Worksheet
    ' 1. Resets all filters (if any are set)
    ' 2. Unhide all hidden cells
    ' 3. Expand all collapsed groups
    ' Written for Microsoft Excel by Max Schmeling
    On Error Resume Next
    Ws.AutoFilter.ShowAllData
    Ws.Columns.EntireColumn.Hidden = False
    Ws.Rows.EntireRow.Hidden = False
    Ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
    On Error GoTo 0
End Sub


Function ColumnLetter(ColumnNumber As Variant) As String
    ' Return the letter for the given column number
    ' E.g.: ColumnLetter(3) = "C"
    ' For Microsoft Excel

    Dim n As Long
    Dim c As Byte
    Dim s As String
    
    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function


Sub Selection_ConvertFormulaToValue()
    ' Convert each formula in the current <Selection> into its value
    ' For Microsoft Excel by Max Schmeling

    Dim Calc_Setting As Integer
    Dim SelectedRange As Range
    
    ' Settings for performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    Calc_Setting = Application.Calculation
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandle
    
    Set SelectedRange = Selection
    
    SelectedRange.Copy
    SelectedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
EndProcess:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = Calc_Setting
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandle:
    MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", _
	vbCritical + vbOKOnly + vbDefaultButton1 + vbMsgBoxSetForeground, "Convert Formulas To Values"
    Err.Clear
    GoTo EndProcess
End Sub


Sub Selection_ConvertTextToNumber()
    ' Convert each text-number into a number in the current <Selection>
    ' For Microsoft Excel by Max Schmeling

    Dim Rng As Range
    Dim Calc_Setting As Integer
    Dim SelectedRange As Range
    
    ' Settings for performance
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = False
    Calc_Setting = Application.Calculation
    Application.Calculation = xlCalculationManual
    On Error GoTo EndProcess
    
    Set SelectedRange = Selection
    SelectedRange.TextToColumns DataType:=xlDelimited, _
				TextQualifier:=xlTextQualifierDoubleQuote, _
				ConsecutiveDelimiter:=False, _
                                Tab:=True, _
				Semicolon:=False, _ 
				Comma:=False, 
				Space:=False, 
				Other:=False, 
				FieldInfo:=Array(1, 1),
				TrailingMinusNumbers:=True
    SelectedRange.NumberFormat = "General"

EndProcess:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = Calc_Setting
    Application.DisplayAlerts = True
    On Error Goto 0
End Sub



Sub Selection_StripWhitespace()
    ' Strip leading, trailing and excess spaces (between words)
    ' off each cell in the current <Selection>
    ' For Microsoft Excel by Max Schmeling

    Dim Rng As Range

    For Each Rng In Selection
	If Not Rng.HasFormula And Not IsError(Rng) Then
        	Rng.Value = WorksheetFunction.Trim$(Rng.Value)
	End if
    Next Rng
End Sub



Sub Selection_ConvertToUpper()
    ' Convert each cell in the current <Selection> to UPPER case
    ' For Microsoft Excel by Max Schmeling dedicated to the community

    Dim Rng As Range

    For Each Rng In Selection
	If Not Rng.HasFormula And Not IsError(Rng) Then
        	Rng.Value = UCase(Rng.Value)
	End if
    Next Rng
End Sub



Sub Selection_ConvertToLower()
    ' Convert each cell in the current <Selection> to lower case
    ' For Microsoft Excel by Max Schmeling

    Dim Rng As Range

    For Each Rng In Selection
	If Not Rng.HasFormula And Not IsError(Rng) Then
        	Rng.Value = LCase(Rng.Value)
	End if
    Next Rng
End Sub



Sub Selection_ConvertToCapitalized()
    ' Convert each cell in the current <Selection> to lower case
    ' For Microsoft Excel by Max Schmeling

    Dim Rng As Range

    For Each Rng In Selection
	If Not Rng.HasFormula And Not IsError(Rng) Then
        	Rng.Value = WorksheetFunction.Proper(Rng.Value)
	End if
    Next Rng
End Sub
