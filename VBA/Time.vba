Function ConvenientTimeUnit(Minutes As Double) As String
    ' 1 day sounds better than 1440 minutes, doesn't it?
    If Minutes < 1 Then
        ConvenientTimeUnit = "seconds"
    ElseIf Minutes < 60 Then
        ConvenientTimeUnit = "minutes"
    ElseIf Minutes < 1440 Then
        ConvenientTimeUnit = "hours"
    ElseIf Minutes < 10080 Then
        ConvenientTimeUnit = "days"
    ElseIf Minutes < 40320 Then
        ConvenientTimeUnit = "weeks"
    ElseIf Minutes < 482840 Then
        ConvenientTimeUnit = "months"
    Else
        ConvenientTimeUnit = "years"
    End If
    
    ' plural or singular
    If WorksheetFunction.RoundDown(ConvenientTime(Minutes), 0) = 1 Then
        ConvenientTimeUnit = Left$(ConvenientTimeUnit, Len(ConvenientTimeUnit) - 1)
    End If
End Function


Function ConvenientTime(Minutes As Double) As Double
    ' 1 day sounds better than 1440 minutes, doesn't it?
    If Minutes < 1 Then
        ConvenientTime = Minutes * 60 ' seconds
    ElseIf Minutes < 60 Then
        ConvenientTime = Minutes ' minutes
    ElseIf Minutes < 1440 Then
        ConvenientTime = Minutes / 60 ' hours
    ElseIf Minutes < 10080 Then
        ConvenientTime = Minutes / 1440 ' days
    ElseIf Minutes < 40320 Then
        ConvenientTime = Minutes / 10080 ' weeks
    ElseIf Minutes < 482840 Then
        ConvenientTime = Minutes / 40320 ' months
    Else
        ConvenientTime = Minutes / 482840
    End If
End Function
