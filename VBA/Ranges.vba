Public Function GetLastRow(Ws As Worksheet) As Long
    ' Returns the last row that contains anything. 0 if error
    On Error Resume Next
    GetLastRow = Ws.UsedRange.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    On Error GoTo 0
End Function


Public Function GetLastColumn(Ws As Worksheet) As Integer
    ' Returns the last column that contains anything. 0 if error
    On Error Resume Next
    GetLastColumn = Ws.UsedRange.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    On Error GoTo 0
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


Function GetUsedTableRange(Ws As Worksheet) As Range
    ' Returns the 'used-value table range' of the Range.UsedRange
    ' Ie. the header row of an assumed table in <Ws> will be determined
    ' using some sort of heuristics.
    '
    ' Attemp1 #1 and #2 are based on formatted excel tables (CTRL+T)
    ' All other attempts try to generate a range using heuristics
    ' and assume exactly one single unformatted (ie. not CTRL+T) table to
    ' exist one the sheet within the first <MAXSTARTROW>'s of the sheet.
    ' If a table-like shape cannot be found in these first rows of the sheet,
    ' the first row is assumed to be the header row.
    '
    ' RETURN VALUE:
    ' The range of the detected table.
    ' If the sheet contains exactly one formatted excel table (ie. CTRL+T),
    ' that table range will be returned. If the sheet contains multiple
    ' formatted excel tables, the one that is currently selected (ActiveCell)
    ' will be returned. If none, is selected a random one will be returned.
    ' The latter behaviour may be undesirable, so use it on your own risk.
    ' If the sheet does not contain one single unformatted table (but e.g.
    ' multiple pivot tables or some other contextual nonsense) and no
    ' formatted excel tables (CTRl+T) this function will return garbage.
    '
    ' REQUIREMENTS:
    ' - Function: IsWorksheetEmpty()
    ' - Function: GetEffectiveUsedRange()

    Dim MAXCOLUMNS As Integer
    MAXCOLUMNS = 2000 ' The maximum amount of columns to be checked for
                      ' to avoid tiring execution times. High number
                      ' will slow down execution time.
    Dim MAXSTARTROW As Long
    MAXSTARTROW = 15 ' The amount of rows to be checked for the first row
                     ' of the table. If table-like shape cannot be detected
                     ' within this amount of rows, a table cannot only
                     ' be detected if it is a formatted one (ie. CTRL+T)
                     ' A higher number will allow detecting a table that
                     ' starts in higher rows of the sheet, but will slow
                     ' down execution time.
    
    Dim RowRng As Range
    Dim UsdRng As Range
    Dim Rng As Range
    Dim i As Integer
    Dim Found As Boolean
    Found = False
    i = 0
    
    On Error GoTo ErrorExit
    
    If IsWorksheetEmpty(Ws) Then
        Set GetUsedTableRange = Nothing
        Exit Function
    End If
    
    ' ATTEMPT #1
    If Ws.ListObjects.Count = 1 Then
        ' There is only one table. Select that one
        Set GetUsedTableRange = Ws.ListObjects(1).Range
        Exit Function
    
    ' ATTEMPT #2
    ElseIf Ws.ListObjects.Count >= 2 Then
        ' There are multiple tables. Check if the user has one selected... take that one
        If Ws Is ActiveSheet Then
            If Not ActiveCell.ListObject Is Nothing Then
                Set GetUsedTableRange = ActiveCell.ListObject.Range
                Exit Function
            End If
            
            ' There are multiple tables but none is selected. Take a random one..
            Set GetUsedTableRange = Ws.ListObjects(1).Range
            Exit Function
        End If
    End If
    
    ' Get effective used range (ie. the minimum rectangular range that contains values)
    Set UsdRng = GetEffectiveUsedRange(Ws)
    If UsdRng Is Nothing Or UsdRng.Rows.Count <= 1 Then
        Set GetUsedTableRange = Nothing
        Exit Function
    End If
    
    ' ATTEMPT #3: First check if user has filter applied. If yes use the row of the autofilter
    If Ws.AutoFilterMode Then
    
        If Ws.AutoFilter.Range.Row >= UsdRng.Row Then
            i = Ws.AutoFilter.Range.Row - UsdRng.Row
            GoTo Finalize
        End If
        
    End If
    
    ' ATTEMPT #4: The user does not have a filter applied so we need to guess the header row. Criteria:
    ' 1. A header row does not contain empty cells
    ' 2. A column header cannot be a number
    
    ' To avoid crashes exit if the table has too many columns
    If UsdRng.Columns.Count >= MAXCOLUMNS Then
        Set GetUsedTableRange = UsdRng
        Exit Function
    End If
    
    ' Iterate over first 10 rows of the effective used range and find the header row
    ' Ie. the first row that does not have an empty cell in it. If this row cannot be
    ' found amongst the first 10 rows, assume the first row to be the header row.
    For Each RowRng In UsdRng.Rows
    
        ' Check if row contains empty cell (because we assume column headers are never empty)
        If WorksheetFunction.CountA(RowRng) = RowRng.Columns.Count Then
        
            ' Check if row does not contain any numbers (because we assume column headers are never just numbers)
            Found = True
            For Each Rng In RowRng.Cells
                If IsNumeric(Rng) Then
                    Found = False
                End If
            Next
            
            ' A column header row has been found.
            If Found Then
                Exit For
            End If
        End If
        
        ' After checking <MAXSTARTROW> rows, stop checking and assume the first row to be the header row
        If i > MAXSTARTROW Then
            i = 0
            Exit For
        End If
        
        i = i + 1
    Next
    
Finalize:
    ' If the header row is not the first row, resize the range to start at the header row
    Set GetUsedTableRange = UsdRng.Offset(i, 0).Resize(UsdRng.Rows.Count - i)
    Exit Function
    
ErrorExit:
    Set GetUsedTableRange = UsdRng
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
    Ws.ShowAllData
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


Function ExtractCellReferences(Rng As Range) As Variant
    ' Returns cell references contained in a formula in a given range as an array of strings
    ' Inspired by: https://www.get-digital-help.com/extract-cell-references-from-a-formula/
    Dim Results As Object
    Dim Pattern As String
    Dim Refs As Variant
    Dim i As Integer
    
    ' this pattern only allows internal references:
    'Pattern = "(\$?[A-Z]{1,3}\$?[0-9]{1,7})(:\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
    
    ' Alternative pattern which allows external references
    Pattern = "'?([a-zA-Z0-9\s\[\]\.])*'?!?\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?"
     
    With CreateObject("vbscript.regexp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = Pattern
        Set Results = .Execute(Rng.FormulaLocal)
    End With
     
    If Results.Count <> 0 Then
        With Results
            ReDim Refs(.Count - 1)
            For i = 0 To .Count - 1
                Refs(i) = .Item(i)
            Next
        End With
        ExtractCellReferences = Refs
    End If
End Function

																	
Sub Reset_Formats()
    ' Resets all format settings to default
    ' For Microsoft Excel
    
    Dim TargetRange As Range
    
    Set TargetRange = Selection
    
    ' Number Format
    TargetRange.NumberFormat = "General"
    
    ' Alignment
    With TargetRange
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Font
    With TargetRange.Font
        .Name = Application.StandardFont
        .FontStyle = "Standard"
        .Size = Application.StandardFontSize
        .Background = xlBackgroundAutomatic
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
'        .Underline = xlUnderlineStyleNone
'        .Bold = xlUnderlineStyleNone
'        .Italic = xlUnderlineStyleNone
    End With
    
    ' Borders
    TargetRange.Borders(xlDiagonalDown).LineStyle = xlNone
    TargetRange.Borders(xlDiagonalUp).LineStyle = xlNone
    TargetRange.Borders(xlEdgeLeft).LineStyle = xlNone
    TargetRange.Borders(xlEdgeTop).LineStyle = xlNone
    TargetRange.Borders(xlEdgeBottom).LineStyle = xlNone
    TargetRange.Borders(xlEdgeRight).LineStyle = xlNone
    TargetRange.Borders(xlInsideVertical).LineStyle = xlNone
    TargetRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ' Interior
    With TargetRange.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Protection
    TargetRange.Locked = True
    TargetRange.FormulaHidden = False
    
    ' Table formats
    Dim Tbl As ListObject
    For Each Tbl In TargetRange.Parent.ListObjects
        If Not Intersect(Tbl.Range, TargetRange) Is Nothing Then
            Tbl.TableStyle = ""
        End If
    Next Tbl
End Sub


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
